#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""eval.py —— 简历生成链路评测脚本（读取 cases.jsonl，循环调用生成脚本，统计三项指标）

设计目标
--------
你的生成逻辑写在 web_app.py 里（Streamlit 页面 + generate_resume_data + parse_json），
没有独立 CLI。所以 eval.py 默认用「inprocess」方式：用桩替换 streamlit，import web_app，
再包一层可录像的 OpenAI 客户端，调用 web_app.generate_resume_data()，从而拿到「模型原始输出」
（用来算 JSON 解析成功率）和「原脚本自己的解析结果」。

三个核心指标
------------
1) JSON 解析成功率
   - strict  : json.loads(原文) 直接成功（说明模型严格守约，没包代码块/没加解释）
   - fenced  : 去掉 ```json 代码块后能解析
   - salvaged: 用括号配对/贪婪正则从混杂文本里抠出 JSON
   - parse_rate = (strict + fenced + salvaged) / 全部用例
2) 必填字段完整率
   - name / intent / contact / sections 四个顶层字段，外加 entry 的 5 个键
   - 占位符（[姓名] 这类）不算填上；但用例可以声明 required_fields，比如素材里本来就没姓名，
     就只要求 intent/contact/sections，避免冤枉模型
3) 幻觉检测（年份）
   - 用正则抓输出里所有 19xx / 20xx 年份，和素材里出现的年份比对
   - 多出来的年份每个扣 --year-penalty 分（默认 15），用例得分 = max(0, 100 - 扣分)

用法
----
    python eval.py                          # 默认：inprocess 适配器，跑全部用例，产物写 eval_results/
    python eval.py --limit 3                # 先花 3 条用例试水
    python eval.py --adapter mock           # 不发请求、不花钱，自测评测脚本本身是否工作
    python eval.py --dry-run                # 只校验用例/参数/适配器，不发请求
    python eval.py --cmd "python gen.py"    # 如果以后把生成逻辑拆成独立 CLI（素材走 stdin，JSON 走 stdout）
    python eval.py --adapter api            # 绕过 web_app，直接用 web_app.SYSTEM_PROMPT 请求 API

用例格式（cases.jsonl，一行一个 JSON 对象）
------------------------------------------
    {"id": "c01", "title": "信息齐全的应届生", "material": "口语化素材……",
     "expect": {"name": "林晓桐"},          # 素材里明确给了姓名才填，没给写 null
     "required_fields": ["name", "intent", "contact", "sections"],
     "expected_sections": ["教育经历", "项目经历"],
     "contact_labels": ["电话", "邮箱"],
     "allowed_years": ["2022"],             # 素材之外额外允许出现的年份（一般不用）
     "notes": "备注"}

    - material 必填；其余可选。required_fields 用来声明「本条真正必需的字段」，
      比如素材本来没姓名（c02/c16），就不要求 name，避免冤枉模型。
    - 想加用例：直接往 cases.jsonl 追加一行，跑 python eval.py --dry-run 校验格式。

产物
----
    eval_results/report.md      人读的评测报告（总览 + 逐条明细 + 口径说明）
    eval_results/summary.json   机读汇总，方便接 CI / 画趋势
    eval_results/results.jsonl  每条用例一行明细（含原始输出片段）
    eval_results/raw/<id>.txt   加 --dump-raw 时保存完整原始输出

注意：脚本不会打印 API Key 明文；Key 依次取 --api-key-env、环境变量 DEEPSEEK_API_KEY、
      .streamlit/secrets.toml 里的 DEEPSEEK_API_KEY。
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import random
import re
import statistics
import subprocess
import sys
import tempfile
import threading
import time
import types
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Sequence

VERSION = "1.0.0"

HERE = Path(__file__).resolve().parent
DEFAULT_CASES = HERE / "cases.jsonl"
DEFAULT_OUT_DIR = HERE / "eval_results"
DEFAULT_MODULE = "web_app.py"
DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"
DEFAULT_TEMPERATURE = 0.3
DEFAULT_MAX_TOKENS = 3000
DEFAULT_YEAR_PENALTY = 15.0

REQUIRED_TOP_FIELDS: tuple[str, ...] = ("name", "intent", "contact", "sections")
REQUIRED_ENTRY_KEYS: tuple[str, ...] = ("period", "org", "role", "title", "bullets")
ALLOWED_SECTIONS: tuple[str, ...] = ("教育经历", "项目经历", "实习经历", "在校经历")

# ---------------- 正则（口径都写在这里，方便改） ----------------
# 代码块：```json ... ``` 或 ``` ... ```（只去掉包裹，不算解析失败，但记 fenced）
_FENCE_RE = re.compile(r"^\s*```[A-Za-z0-9_.+\-]*\s*\n?(.*?)\n?\s*```\s*$", re.S)
# 年份：前后不能再接数字，避免把 2022.9 里的 2022 抓成 20229 之类
_YEAR_RE = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
# 年月：2022.9 / 2022年9月 / 2022-09（只用于校验 period 里的月份是否合法）
_YEAR_MONTH_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})\s*[./年]\s*(\d{1,2})(?!\d)")
# 占位符：[姓名] [学校名称] 这类，1~40 字
_PLACEHOLDER_RE = re.compile(r"\[[^\[\]]{1,40}\]")
# 必填要点领起词：**项目背景**
_LEADIN_RE = re.compile(r"^\*\*[\u4e00-\u9fa5A-Za-z0-9]{1,8}\*\*")
# 组织名（辅助检测，只提示不计分）：XX大学 / XX公司 / XX集团 ...
_ORG_RE = re.compile(r"[\u4e00-\u9fa5A-Za-z0-9]{2,12}(?:大学|学院|学校|公司|集团|科技|银行|医院|研究院|事务所)")
# 邮箱 / 网址：抓年份前先剔掉，避免 apple2023@example.com 被误判成幻觉年份
_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
_URL_RE = re.compile(r"https?://\S+|www\.\S+")

# 重试时认这些关键词（限流/超时/JSON 抖动都值得再试一次）
_RETRY_HINTS = ("json", "expecting", "timeout", "timed out", "rate limit", "429",
                "500", "502", "503", "connection", "temporarily", "overloaded")

_PRINT_LOCK = threading.Lock()


# ======================================================================
# 0. 小工具
# ======================================================================
def _setup_stdout() -> None:
    """Windows 控制台默认 GBK，中文/emoji 会炸，统一成 utf-8。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except Exception:
            pass


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(value).strip()
    return ""


def _norm(text: str) -> str:
    """全角数字/括号统一成半角，避免 ２０２２ 抓不到。"""
    return unicodedata.normalize("NFKC", text or "")


def years_in_text(text: str) -> set[str]:
    return set(_YEAR_RE.findall(_norm(text)))


def _as_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return list(value)
    if isinstance(value, str):
        return [value] if value.strip() else []
    return []


def flatten_strings(node: Any) -> list[str]:
    out: list[str] = []
    if isinstance(node, str):
        out.append(node)
    elif isinstance(node, dict):
        for value in node.values():
            out.extend(flatten_strings(value))
    elif isinstance(node, (list, tuple)):
        for value in node:
            out.extend(flatten_strings(value))
    return out


def _pct(value: float | None) -> str:
    if value is None:
        return "-"
    return f"{value * 100:.1f}%"


def _mark(flag: bool | None, ok: str = "OK", bad: str = "FAIL", none: str = "-") -> str:
    if flag is None:
        return none
    return ok if flag else bad


def _percentile(values: Sequence[float], q: float) -> float | None:
    if not values:
        return None
    ordered = sorted(values)
    if len(ordered) == 1:
        return float(ordered[0])
    return float(statistics.quantiles(ordered, n=100)[max(0, min(99, int(q) - 1))])


# ======================================================================
# 1. 用例
# ======================================================================
@dataclass(frozen=True)
class Case:
    id: str
    material: str
    title: str = ""
    expect_name: str | None = None
    allowed_years: tuple[str, ...] = ()
    required_fields: tuple[str, ...] = REQUIRED_TOP_FIELDS
    expected_sections: tuple[str, ...] = ()
    contact_labels: tuple[str, ...] = ()
    notes: str = ""

    @property
    def allowed_year_set(self) -> set[str]:
        return set(self.allowed_years) | years_in_text(self.material)


def load_cases(path: Path, limit: int | None = None) -> list[Case]:
    if not path.exists():
        raise SystemExit(f"[cases] 找不到用例文件：{path}")
    cases: list[Case] = []
    seen: set[str] = set()
    with path.open("r", encoding="utf-8-sig") as fh:
        for lineno, raw_line in enumerate(fh, 1):
            line = raw_line.strip()
            if not line or line.startswith("#") or line.startswith("//"):
                continue
            try:
                obj = json.loads(line)
            except json.JSONDecodeError as exc:
                raise SystemExit(f"[cases] 第 {lineno} 行不是合法 JSON：{exc}") from exc
            if not isinstance(obj, dict):
                raise SystemExit(f"[cases] 第 {lineno} 行必须是 JSON 对象，实际是 {type(obj).__name__}")

            cid = _as_text(obj.get("id")) or f"line{lineno}"
            material = _as_text(
                obj.get("material") or obj.get("input") or obj.get("raw") or obj.get("text")
            )
            if not material:
                raise SystemExit(f"[cases] 第 {lineno} 行（{cid}）缺少 material 素材字段")
            if cid in seen:
                raise SystemExit(f"[cases] 用例 id 重复：{cid}")
            seen.add(cid)

            expect = obj.get("expect") if isinstance(obj.get("expect"), dict) else {}
            expect_name = _as_text(expect.get("name") or obj.get("expected_name")) or None
            allowed = _as_list(obj.get("allowed_years") or expect.get("years"))
            allowed_years = tuple(
                y for y in (_as_text(x) for x in allowed if _as_text(x)) if _YEAR_RE.fullmatch(y)
            )

            required = _as_list(obj.get("required_fields") or expect.get("required_fields")) or list(
                REQUIRED_TOP_FIELDS
            )
            required_fields: list[str] = []
            for name in required:
                key = _as_text(name)
                if not key:
                    continue
                if key not in REQUIRED_TOP_FIELDS:
                    raise SystemExit(
                        f"[cases] {cid} 的 required_fields 里有未知字段 {key!r}，"
                        f"只支持 {REQUIRED_TOP_FIELDS}"
                    )
                if key not in required_fields:
                    required_fields.append(key)

            sections = _as_list(obj.get("expected_sections") or expect.get("sections"))
            labels = _as_list(obj.get("contact_labels") or expect.get("contact_labels"))

            cases.append(
                Case(
                    id=cid,
                    material=material,
                    title=_as_text(obj.get("title")),
                    expect_name=expect_name,
                    allowed_years=allowed_years,
                    required_fields=tuple(required_fields),
                    expected_sections=tuple(_as_text(s) for s in sections if _as_text(s)),
                    contact_labels=tuple(_as_text(s) for s in labels if _as_text(s)),
                    notes=_as_text(obj.get("notes")),
                )
            )

    if not cases:
        raise SystemExit(f"[cases] {path} 里没有有效用例")
    if limit is not None and limit > 0:
        cases = cases[:limit]

    for case in cases:
        if case.expect_name and case.expect_name not in case.material:
            print(
                f"[cases] 提示：{case.id} 的 expect.name={case.expect_name!r} 没在素材里出现，"
                "姓名准确率可能被误判",
                file=sys.stderr,
            )
    return cases


# ======================================================================
# 2. 解析与度量
# ======================================================================
def _json_object(text: str) -> tuple[dict | None, str | None]:
    try:
        value = json.loads(text)
    except json.JSONDecodeError as exc:
        return None, f"{exc.msg}（line {exc.lineno} col {exc.colno}）"
    if not isinstance(value, dict):
        return None, f"顶层类型是 {type(value).__name__}，不是 JSON 对象"
    return value, None


def _strip_fence(text: str) -> str:
    match = _FENCE_RE.match(text)
    return match.group(1).strip() if match else text


def _extract_object(text: str) -> str | None:
    """从第一个 { 开始做括号配对（跳过字符串内的括号/转义），比贪婪正则靠谱。"""
    start = text.find("{")
    if start < 0:
        return None
    depth = 0
    in_string = False
    escaped = False
    for idx in range(start, len(text)):
        char = text[idx]
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start: idx + 1]
    return None


def parse_output(raw: str) -> tuple[dict | None, str, str | None]:
    """返回 (data, mode, error)；mode ∈ strict / fenced / salvaged / failed。"""
    text = (raw or "").strip()
    if not text:
        return None, "failed", "空输出"

    data, err = _json_object(text)
    if data is not None:
        return data, "strict", None

    cleaned = _strip_fence(text)
    if cleaned != text:
        data, err = _json_object(cleaned)
        if data is not None:
            return data, "fenced", None

    # 整段就是合法 JSON，但顶层不是对象（数组/字符串/数字）：属于契约违规，不当成“能救回”
    try:
        whole = json.loads(cleaned)
    except json.JSONDecodeError:
        whole = None
    if whole is not None and not isinstance(whole, dict):
        return None, "failed", f"顶层类型是 {type(whole).__name__}，不是 JSON 对象"

    for candidate in (cleaned, text):
        snippet = _extract_object(candidate)
        if snippet:
            data, err = _json_object(snippet)
            if data is not None:
                return data, "salvaged", None

    # 最后兜底：和 web_app.parse_json 一样的贪婪正则
    match = re.search(r"\{.*\}", text, re.S)
    if match:
        data, err = _json_object(match.group(0))
        if data is not None:
            return data, "salvaged", None

    return None, "failed", err or "没有找到 JSON 对象"


def _is_placeholder(value: str) -> bool:
    return bool(_PLACEHOLDER_RE.fullmatch((value or "").strip()))


@dataclass
class FieldReport:
    present: dict[str, bool] = field(default_factory=dict)
    missing: list[str] = field(default_factory=list)
    placeholders: list[str] = field(default_factory=list)
    name_accuracy: bool | None = None


@dataclass
class HallucinationReport:
    output_years: list[str] = field(default_factory=list)
    allowed_years: list[str] = field(default_factory=list)
    extra_years: list[str] = field(default_factory=list)
    extra_orgs: list[str] = field(default_factory=list)
    score: float = 100.0

    @property
    def clean(self) -> bool:
        return not self.extra_years


def check_fields(data: dict | None, case: Case) -> FieldReport:
    present = {name: False for name in REQUIRED_TOP_FIELDS}
    missing: list[str] = []
    placeholders: list[str] = []

    if not isinstance(data, dict):
        report = FieldReport(present=present, missing=["<输出不是 JSON 对象>"])
        report.name_accuracy = False if case.expect_name else None
        return report

    # ---- name ----
    name = _as_text(data.get("name"))
    if not name:
        missing.append("name")
    elif _is_placeholder(name):
        placeholders.append("name")
    else:
        present["name"] = True

    # ---- intent ----
    intent = _as_text(data.get("intent"))
    if not intent:
        missing.append("intent")
    elif _is_placeholder(intent):
        placeholders.append("intent")
    else:
        present["intent"] = True

    # ---- contact ----
    contact = data.get("contact")
    if not isinstance(contact, list):
        missing.append("contact")
    else:
        labels: set[str] = set()
        valid = 0
        for idx, item in enumerate(contact):
            if not isinstance(item, dict):
                missing.append(f"contact[{idx}] 不是对象")
                continue
            label = _as_text(item.get("label"))
            value = _as_text(item.get("value"))
            if label:
                labels.add(label)
            else:
                missing.append(f"contact[{idx}].label")
            if not value:
                missing.append(f"contact[{idx}].value")
            elif _is_placeholder(value):
                placeholders.append(f"contact[{idx}].value")
            else:
                valid += 1
        if valid == 0:
            missing.append("contact 里没有填好的项")
        else:
            lacking = [label for label in case.contact_labels if label not in labels]
            if lacking:
                missing.append("contact 缺少标签：" + "、".join(lacking))
            else:
                present["contact"] = True

    # ---- sections ----
    sections = data.get("sections")
    if not isinstance(sections, list) or not sections:
        missing.append("sections")
    else:
        structure_ok = True
        total_bullets = 0
        for s_idx, section in enumerate(sections):
            if not isinstance(section, dict):
                missing.append(f"sections[{s_idx}] 不是对象")
                structure_ok = False
                continue
            title = _as_text(section.get("title"))
            if not title:
                missing.append(f"sections[{s_idx}].title")
                structure_ok = False
            entries = section.get("entries")
            if not isinstance(entries, list) or not entries:
                missing.append(f"sections[{s_idx}].entries 为空")
                structure_ok = False
                continue
            for e_idx, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    missing.append(f"sections[{s_idx}].entries[{e_idx}] 不是对象")
                    structure_ok = False
                    continue
                for key in REQUIRED_ENTRY_KEYS:
                    if key not in entry:
                        missing.append(f"{title or 'entry'}[{e_idx}] 缺少键 {key}")
                        structure_ok = False
                bullets = entry.get("bullets")
                if not isinstance(bullets, list):
                    missing.append(f"{title or 'entry'}[{e_idx}].bullets 不是数组")
                    structure_ok = False
                else:
                    total_bullets += sum(1 for b in bullets if _as_text(b))
        if structure_ok and total_bullets == 0:
            # 教育经历这类条目允许 bullets 为空，但整份简历不能一条要点都没有
            missing.append("所有条目都没有要点文本")
            structure_ok = False
        if structure_ok:
            present["sections"] = True

    report = FieldReport(present=present, missing=missing, placeholders=placeholders)
    if case.expect_name:
        report.name_accuracy = bool(
            name and not _is_placeholder(name) and (case.expect_name in name or name in case.expect_name)
        )
    return report


def check_hallucination(data: dict | None, case: Case, year_penalty: float) -> HallucinationReport:
    if not isinstance(data, dict):
        return HallucinationReport()

    text = "\n".join(flatten_strings(data))
    # 邮箱/网址里的数字不是简历内容（如 apple2023@example.com），先剔除再抓年份，避免误报
    normalized = _norm(_EMAIL_RE.sub(" ", _URL_RE.sub(" ", text)))
    output_years = sorted(set(_YEAR_RE.findall(normalized)), key=int)
    allowed_set = case.allowed_year_set
    extra_years = [year for year in output_years if year not in allowed_set]

    # 组织/学校名：只作提示（输出可能截断或换写法，误报率高，不计分）
    orgs = {_norm(m.group(0)) for m in _ORG_RE.finditer(normalized)}
    material_norm = _norm(case.material)
    extra_orgs = sorted(org for org in orgs if org not in material_norm)

    score = max(0.0, 100.0 - year_penalty * len(extra_years))
    return HallucinationReport(
        output_years=output_years,
        allowed_years=sorted(allowed_set, key=int),
        extra_years=extra_years,
        extra_orgs=extra_orgs,
        score=score,
    )


def collect_quality(data: dict | None) -> dict[str, Any]:
    if not isinstance(data, dict):
        return {"placeholders": 0, "bullets": 0, "long_bullets": 0,
                "missing_leadin": 0, "invalid_periods": [], "unknown_sections": []}

    text = "\n".join(flatten_strings(data))
    bullets: list[str] = []
    bad_periods: list[str] = []
    unknown_sections: set[str] = set()
    for section in data.get("sections") or []:
        if not isinstance(section, dict):
            continue
        title = _as_text(section.get("title"))
        if title and title not in ALLOWED_SECTIONS:
            unknown_sections.add(title)
        for entry in section.get("entries") or []:
            if not isinstance(entry, dict):
                continue
            for bullet in entry.get("bullets") or []:
                if isinstance(bullet, str) and bullet.strip():
                    bullets.append(bullet.strip())
            period = _as_text(entry.get("period"))
            for match in _YEAR_MONTH_RE.finditer(period):
                if not 1 <= int(match.group(2)) <= 12:
                    bad_periods.append(period)

    return {
        "placeholders": len(_PLACEHOLDER_RE.findall(text)),
        "bullets": len(bullets),
        "long_bullets": sum(1 for b in bullets if len(b) > 80),
        "missing_leadin": sum(1 for b in bullets if not _LEADIN_RE.match(b)),
        "invalid_periods": sorted(set(bad_periods)),
        "unknown_sections": sorted(unknown_sections),
    }


# ======================================================================
# 3. 适配器：怎么「调用你的脚本」
# ======================================================================
@dataclass
class Invocation:
    raw: str = ""
    func_data: dict | None = None      # 原脚本自己解析出的 dict（可能为 None）
    func_error: str | None = None      # 原脚本抛的异常
    usage: dict[str, Any] = field(default_factory=dict)
    mode: str = ""                     # 仅 mock 用，标记伪造行为


class _AnyStub:
    """streamlit 的万能桩：任何属性都返回可调用、可 with、可迭代的假对象。"""

    def __init__(self, name: str = "st") -> None:
        object.__setattr__(self, "_name", name)

    def __repr__(self) -> str:
        return f"<stub {self._name}>"

    def __call__(self, *args: Any, **kwargs: Any) -> "_AnyStub":
        if len(args) == 1 and not kwargs and callable(args[0]):  # @st.cache_xxx 这类装饰器
            return args[0]
        return _AnyStub(f"{self._name}()")

    def __getattr__(self, item: str) -> "_AnyStub":
        if item.startswith("__"):
            raise AttributeError(item)
        return _AnyStub(f"{self._name}.{item}")

    def __enter__(self) -> "_AnyStub":
        return self

    def __exit__(self, *exc_info: Any) -> bool:
        return False

    def __iter__(self):
        return iter(())

    def __bool__(self) -> bool:
        return False

    def __contains__(self, item: Any) -> bool:
        return False

    def __getitem__(self, item: Any) -> "_AnyStub":
        return _AnyStub(f"{self._name}[{item!r}]")

    def __setitem__(self, key: Any, value: Any) -> None:
        return None

    def get(self, *args: Any, **kwargs: Any) -> None:
        return None

    def text_input(self, *args: Any, **kwargs: Any) -> str:
        return ""

    def file_uploader(self, *args: Any, **kwargs: Any) -> None:
        return None


def _install_streamlit_stub() -> None:
    """在 import web_app 之前把 streamlit 换掉，避免执行页面代码/依赖运行环境。"""
    existing = sys.modules.get("streamlit")
    if getattr(existing, "__eval_stub__", False):
        return

    module = types.ModuleType("streamlit")
    module.__eval_stub__ = True  # type: ignore[attr-defined]

    def identity_decorator(*dargs: Any, **dkwargs: Any):
        if len(dargs) == 1 and not dkwargs and callable(dargs[0]):
            return dargs[0]
        return lambda func: func

    module.cache_resource = identity_decorator  # type: ignore[attr-defined]
    module.cache_data = identity_decorator  # type: ignore[attr-defined]
    module.secrets = _AnyStub("secrets")  # type: ignore[attr-defined]
    module.session_state = _AnyStub("session_state")  # type: ignore[attr-defined]
    module.sidebar = _AnyStub("sidebar")  # type: ignore[attr-defined]

    def module_getattr(name: str) -> _AnyStub:
        if name.startswith("__"):
            raise AttributeError(name)
        return _AnyStub(name)

    module.__getattr__ = module_getattr  # type: ignore[attr-defined]  # PEP 562
    sys.modules["streamlit"] = module


class _RecordingCompletions:
    def __init__(self, inner: Any, box: dict[str, Any]) -> None:
        self._inner = inner
        self._box = box

    def create(self, *args: Any, **kwargs: Any):
        response = self._inner.create(*args, **kwargs)
        try:
            self._box["raw"] = response.choices[0].message.content or ""
        except Exception:
            self._box["raw"] = ""
        try:
            usage = getattr(response, "usage", None)
            if usage is not None:
                self._box["usage"] = {
                    "prompt_tokens": getattr(usage, "prompt_tokens", None),
                    "completion_tokens": getattr(usage, "completion_tokens", None),
                    "total_tokens": getattr(usage, "total_tokens", None),
                }
        except Exception:
            pass
        return response

    def __getattr__(self, name: str) -> Any:
        return getattr(self._inner, name)


class _RecordingChat:
    def __init__(self, inner: Any, box: dict[str, Any]) -> None:
        self._inner = inner
        self._box = box

    @property
    def completions(self) -> _RecordingCompletions:
        return _RecordingCompletions(self._inner.completions, self._box)

    def __getattr__(self, name: str) -> Any:
        return getattr(self._inner, name)


class _RecordingClient:
    def __init__(self, inner: Any, box: dict[str, Any]) -> None:
        self._inner = inner
        self._box = box

    @property
    def chat(self) -> _RecordingChat:
        return _RecordingChat(self._inner.chat, self._box)

    def __getattr__(self, name: str) -> Any:
        return getattr(self._inner, name)


def _callable_positional_count(func: Any) -> int:
    import inspect

    try:
        params = list(inspect.signature(func).parameters.values())
    except (TypeError, ValueError):
        return 2
    return sum(
        1
        for p in params
        if p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)
    )


class InProcessInvoker:
    """默认适配器：import 你的 web_app，调用它的 generate_resume_data()。"""

    kind = "inprocess"

    def __init__(self, module_path: Path, api_key: str, base_url: str, model: str,
                 temperature: float, max_tokens: int) -> None:
        self.module_path = module_path
        self.module = load_pipeline_module(module_path)
        func = getattr(self.module, "generate_resume_data", None)
        if not callable(func):
            raise SystemExit(
                f"[adapter] {module_path.name} 里没有 generate_resume_data()，"
                "请用 --adapter api 或 --cmd 指定调用方式"
            )
        self.func = func
        self.arity = _callable_positional_count(func)
        self.client = _make_client(api_key, base_url)
        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens

    def invoke(self, case: Case) -> Invocation:
        box: dict[str, Any] = {}
        proxy = _RecordingClient(self.client, box)
        data: dict | None = None
        error: str | None = None
        try:
            result = self.func(proxy, case.material) if self.arity >= 2 else self.func(case.material)
            if isinstance(result, dict):
                data = result
        except Exception as exc:  # 原脚本解析失败时仍要拿到 raw 输出
            error = f"{type(exc).__name__}: {exc}"
        raw = str(box.get("raw") or "")
        return Invocation(raw=raw, func_data=data, func_error=error,
                          usage=box.get("usage") or {})


FALLBACK_SYSTEM_PROMPT = """你是一位简历排版顾问。请把求职者素材整理成结构化简历 JSON，只使用素材中的事实，
缺失的信息写 [占位符]。只输出一个 JSON 对象，不要代码块、不要解释。字段：
{"name": "...", "intent": "...", "contact": [{"label": "...", "value": "..."}],
 "sections": [{"title": "教育经历", "icon": "🎓",
   "entries": [{"period": "", "org": "", "role": "", "title": "", "bullets": ["..."]}]}]}
板块只能是：教育经历 / 项目经历 / 实习经历 / 在校经历，素材没有就省略。"""


class ApiInvoker:
    """绕过 web_app 的函数，直接用它的 SYSTEM_PROMPT 请求 API（模型/参数可覆盖）。"""

    kind = "api"

    def __init__(self, module_path: Path | None, api_key: str, base_url: str,
                 model: str, temperature: float, max_tokens: int) -> None:
        self.client = _make_client(api_key, base_url)
        self.module = None
        self.system_prompt = FALLBACK_SYSTEM_PROMPT
        if module_path is not None and module_path.exists():
            try:
                module = load_pipeline_module(module_path)
            except Exception as exc:
                print(f"[adapter] 读取 SYSTEM_PROMPT 失败，改用内置提示词：{exc}", file=sys.stderr)
            else:
                self.module = module
                prompt = getattr(module, "SYSTEM_PROMPT", None)
                if isinstance(prompt, str) and prompt.strip():
                    self.system_prompt = prompt
        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens

    def invoke(self, case: Case) -> Invocation:
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": f"这是求职者的原始素材：\n\n{case.material}"},
            ],
            response_format={"type": "json_object"},
            temperature=self.temperature,
            max_tokens=self.max_tokens,
        )
        raw = (response.choices[0].message.content or "").strip()
        usage = {}
        try:
            usage = {
                "prompt_tokens": response.usage.prompt_tokens,
                "completion_tokens": response.usage.completion_tokens,
                "total_tokens": response.usage.total_tokens,
            }
        except Exception:
            pass
        return Invocation(raw=raw, usage=usage)


class CliInvoker:
    """把素材喂给外部命令（stdin 或 {input_file}），命令的 stdout 当成模型原始输出。"""

    kind = "cli"

    def __init__(self, cmd: str, timeout: float, cwd: Path | None) -> None:
        self.cmd = cmd
        self.timeout = timeout
        self.cwd = cwd
        self.module = None

    def invoke(self, case: Case) -> Invocation:
        cmd = self.cmd
        tmp_path: Path | None = None
        if "{input_file}" in cmd:
            handle = tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8")
            handle.write(case.material)
            handle.close()
            tmp_path = Path(handle.name)
            cmd = cmd.replace("{input_file}", str(tmp_path))
        try:
            proc = subprocess.run(
                cmd,
                shell=True,
                cwd=str(self.cwd) if self.cwd else None,
                input=None if tmp_path else case.material,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=self.timeout,
            )
        except subprocess.TimeoutExpired as exc:
            return Invocation(func_error=f"命令超时（{self.timeout:.0f}s）：{exc}")
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

        raw = (proc.stdout or "").strip()
        error = None
        if proc.returncode != 0:
            error = f"退出码 {proc.returncode}：{(proc.stderr or '').strip()[:400]}"
        return Invocation(raw=raw, func_error=error)


MOCK_BEHAVIOR: dict[str, str] = {
    "c01": "clean",
    "c02": "clean_no_name",
    "c03": "fence_json",
    "c04": "missing_contact",
    "c05": "placeholder_name",
    "c06": "truncated",
    "c07": "extra_year",
    "c08": "array_output",
    "c09": "empty_sections",
    "c10": "bad_entry_keys",
    "c11": "fence_prose",
    "c12": "clean",
    "c13": "clean",
    "c14": "clean",
    "c15": "clean",
    "c16": "clean",
    "c17": "clean",
    "c18": "clean",
    "c19": "clean",
    "c20": "clean",
}


class MockInvoker:
    """离线假模型：验证评测脚本自身（不用 API、不花钱）。按用例 id 注入典型故障。"""

    kind = "mock"
    module = None

    def invoke(self, case: Case) -> Invocation:
        data = self._base(case)
        behavior = MOCK_BEHAVIOR.get(case.id, "clean")

        if behavior == "clean_no_name":
            data["name"] = "[姓名]"
        elif behavior == "placeholder_name":
            data["name"] = "[姓名]"
        elif behavior == "missing_contact":
            data.pop("contact", None)
        elif behavior == "empty_sections":
            data["sections"] = []
        elif behavior == "bad_entry_keys":
            for section in data["sections"]:
                for entry in section["entries"]:
                    entry.pop("role", None)
                    entry.pop("title", None)
        elif behavior == "extra_year":
            data["sections"][0]["entries"][0]["bullets"].append("**荣誉奖项** 2018 年拿过校级奖学金")
        elif behavior == "fence_json":
            return Invocation(raw="```json\n" + json.dumps(data, ensure_ascii=False, indent=2) + "\n```",
                              mode="fence_json")
        elif behavior == "fence_prose":
            return Invocation(raw="好的，下面是整理后的 JSON：\n```json\n"
                                  + json.dumps(data, ensure_ascii=False, indent=2)
                                  + "\n```\n希望对你有帮助。", mode="fence_prose")
        elif behavior == "truncated":
            text = json.dumps(data, ensure_ascii=False)
            return Invocation(raw=text[: max(20, len(text) // 2)], mode="truncated")
        elif behavior == "array_output":
            return Invocation(raw=json.dumps([data], ensure_ascii=False), mode="array_output")

        return Invocation(raw=json.dumps(data, ensure_ascii=False), mode=behavior)

    def _base(self, case: Case) -> dict[str, Any]:
        years = sorted(years_in_text(case.material), key=int)
        period = f"{years[0]}.9-{years[-1]}.6" if len(years) >= 2 else ""
        sections = []
        for title in case.expected_sections or ("教育经历", "项目经历"):
            if title == "教育经历":
                entry = {"period": period, "org": "[学校名称]", "role": "[专业] | 本科",
                         "title": "", "bullets": ["**学分绩点** 素材里没写具体数字，这里不编造"]}
            elif title == "项目经历":
                entry = {"period": "", "org": "", "role": "[角色]",
                         "title": "[项目名称]", "bullets": ["**项目背景** 用 Python 做了一个小工具"]}
            elif title == "实习经历":
                entry = {"period": period, "org": "[公司名称]", "role": "[岗位]",
                         "title": "", "bullets": ["**主要工作** 整理需求并跟进落地"]}
            else:
                entry = {"period": "", "org": "", "role": "",
                         "title": "[职务]", "bullets": ["**组织工作** 组织过班级活动"]}
            sections.append({"title": title, "icon": "", "entries": [entry]})
        return {
            "name": case.expect_name or "[姓名]",
            "intent": "AI 助理工程师",
            "contact": [
                {"label": "电话", "value": "13800000000"},
                {"label": "邮箱", "value": "mock@example.com"},
            ],
            "sections": sections,
        }


def _make_client(api_key: str, base_url: str):
    from openai import OpenAI

    return OpenAI(api_key=api_key, base_url=base_url, timeout=90.0, max_retries=2)


def load_pipeline_module(path: Path):
    """加载 web_app.py，但先把 streamlit 换成桩，避免执行页面 UI。"""
    if not path.exists():
        raise SystemExit(f"[adapter] 找不到模块：{path}")
    _install_streamlit_stub()
    name = path.stem or "pipeline"
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise SystemExit(f"[adapter] 无法加载模块：{path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except ImportError as exc:
        sys.modules.pop(name, None)
        raise SystemExit(
            f"[adapter] 加载 {path.name} 失败，可能缺少依赖"
            f"（先 pip install -r requirements.txt）：{exc}"
        ) from exc
    except Exception:
        sys.modules.pop(name, None)
        raise
    return module


def _resolve_module_path(value: str) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    for base in (Path.cwd(), HERE):
        candidate = base / value
        if candidate.exists():
            return candidate
    return HERE / value


def _resolve_api_key(args: argparse.Namespace, project_dir: Path | None = None) -> tuple[str | None, str]:
    """返回 (key, 来源说明)，绝不回显明文。"""
    env_name = args.api_key_env or "DEEPSEEK_API_KEY"
    key = os.getenv(env_name, "").strip()
    if key:
        return key, f"环境变量 {env_name}"

    candidates: list[Path] = [HERE / ".streamlit" / "secrets.toml"]
    if project_dir is not None:
        candidate = project_dir / ".streamlit" / "secrets.toml"
        if candidate not in candidates:
            candidates.append(candidate)

    for secrets_path in candidates:
        if not secrets_path.exists():
            continue
        try:
            import tomllib

            with secrets_path.open("rb") as fh:
                data = tomllib.load(fh)
            key = str(data.get("DEEPSEEK_API_KEY") or "").strip()
        except Exception:
            match = re.search(r'DEEPSEEK_API_KEY\s*=\s*"([^"]+)"', secrets_path.read_text(encoding="utf-8"))
            key = match.group(1).strip() if match else ""
        if key:
            return key, f"{secrets_path.name}（非环境变量，未回显）"

    return None, "未找到"


def _retryable(inv: Invocation) -> bool:
    if not inv.raw and not inv.func_data:
        return True
    error = (inv.func_error or "").lower()
    return bool(error) and any(hint in error for hint in _RETRY_HINTS)


def invoke_with_retry(invoker: Any, case: Case, retries: int) -> Invocation:
    attempts = max(1, retries + 1)
    last_error: str | None = None
    for attempt in range(attempts):
        try:
            inv = invoker.invoke(case)
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            if attempt < attempts - 1:
                time.sleep(min(8.0, 1.5 * (2 ** attempt)) + random.random() * 0.3)
                continue
            return Invocation(func_error=last_error)

        if attempt < attempts - 1 and _retryable(inv):
            time.sleep(min(8.0, 1.5 * (2 ** attempt)) + random.random() * 0.3)
            continue
        return inv

    return Invocation(func_error=last_error or "调用失败")


# ======================================================================
# 4. 单条用例执行
# ======================================================================
@dataclass
class CaseResult:
    id: str
    title: str
    parse_ok: bool
    mode: str
    strict_ok: bool
    required_ok: bool
    required_fields: tuple[str, ...]
    field_present: dict[str, bool]
    missing: list[str]
    placeholders: list[str]
    extra_years: list[str]
    extra_orgs: list[str]
    allowed_years: list[str]
    output_years: list[str]
    hall_score: float
    name_accuracy: bool | None
    latency_ms: float
    error: str | None
    render_error: str | None
    quality: dict[str, Any]
    raw_excerpt: str
    usage: dict[str, Any] = field(default_factory=dict)
    parsed_data: dict | None = field(default=None, repr=False)
    render_checked: bool = False
    raw_full: str = field(default="", repr=False)

    @property
    def hall_clean(self) -> bool:
        return not self.extra_years

    @property
    def ok(self) -> bool:
        return self.parse_ok and self.required_ok and self.hall_clean

    def to_row(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "title": self.title,
            "ok": self.ok,
            "parse_ok": self.parse_ok,
            "parse_mode": self.mode,
            "strict_json": self.strict_ok,
            "required_ok": self.required_ok,
            "required_fields": list(self.required_fields),
            "field_present": self.field_present,
            "missing": self.missing,
            "placeholders": self.placeholders,
            "output_years": self.output_years,
            "allowed_years": self.allowed_years,
            "extra_years": self.extra_years,
            "hallucination_score": round(self.hall_score, 1),
            "name_accuracy": self.name_accuracy,
            "latency_ms": round(self.latency_ms, 1),
            "error": self.error,
            "render_error": self.render_error,
            "extra_orgs": self.extra_orgs,
            "quality": self.quality,
            "usage": self.usage,
            "raw_excerpt": self.raw_excerpt,
        }


def run_case(invoker: Any, case: Case, args: argparse.Namespace) -> CaseResult:
    started = time.perf_counter()
    invocation = invoke_with_retry(invoker, case, args.retries)
    latency_ms = (time.perf_counter() - started) * 1000.0

    data, mode, parse_error = parse_output(invocation.raw)
    if data is None and invocation.func_data is not None:
        # 原脚本自己解析成功了（说明它的 parse_json 比我们宽松），按可用数据处理
        data, mode, parse_error = invocation.func_data, "func", None
    parse_ok = data is not None

    if parse_ok:
        field_report = check_fields(data, case)
        hall = check_hallucination(data, case, args.year_penalty)
        quality = collect_quality(data)
        required_ok = all(field_report.present.get(f, False) for f in case.required_fields)
    else:
        field_report = FieldReport(present={f: False for f in REQUIRED_TOP_FIELDS},
                                   missing=["<JSON 解析失败>"],
                                   name_accuracy=False if case.expect_name else None)
        hall = HallucinationReport()
        quality = collect_quality(None)
        required_ok = False

    error = invocation.func_error or (None if parse_ok else (parse_error or "解析失败"))
    return CaseResult(
        id=case.id,
        title=case.title,
        parse_ok=parse_ok,
        mode=mode,
        strict_ok=mode == "strict",
        required_ok=required_ok,
        required_fields=case.required_fields,
        field_present=field_report.present,
        missing=field_report.missing,
        placeholders=field_report.placeholders,
        extra_years=hall.extra_years,
        extra_orgs=hall.extra_orgs,
        allowed_years=hall.allowed_years,
        output_years=hall.output_years,
        hall_score=hall.score,
        name_accuracy=field_report.name_accuracy,
        latency_ms=latency_ms,
        error=error,
        render_error=None,
        quality=quality,
        raw_excerpt=(invocation.raw or "")[:500],
        usage=invocation.usage,
        parsed_data=data if parse_ok else None,
        raw_full=invocation.raw or "",
    )


def check_render(module: Any, data: dict | None) -> str | None:
    """用原项目的 normalize + build_resume_bytes 真的生成一次 docx，验证端到端可排版。"""
    if module is None or data is None:
        return "未解析出 JSON，跳过"
    normalize = getattr(module, "normalize", None)
    build = getattr(module, "build_resume_bytes", None)
    if not callable(normalize) or not callable(build):
        return "模块缺少 normalize / build_resume_bytes，跳过"
    try:
        normalized = normalize(data)
        if not normalized.get("sections"):
            return "normalize 后没有任何经历板块"
        build(normalized, None)
        return None
    except Exception as exc:
        return f"{type(exc).__name__}: {exc}"


def run_all(invoker: Any, cases: list[Case], args: argparse.Namespace) -> list[CaseResult]:
    results_by_id: dict[str, CaseResult] = {}

    def work(case: Case) -> CaseResult:
        return run_case(invoker, case, args)

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futures = {pool.submit(work, case): case for case in cases}
        done = 0
        total = len(cases)
        for future in as_completed(futures):
            case = futures[future]
            try:
                result = future.result()
            except Exception as exc:  # 理论上不会到这里，兜底
                result = _failed_result(case, f"{type(exc).__name__}: {exc}")
            results_by_id[case.id] = result
            done += 1
            if not args.quiet:
                line = (f"[{done:>2}/{total}] {case.id:<4} "
                        f"JSON={'OK ' if result.parse_ok else 'FAIL'}  "
                        f"字段={'OK ' if result.required_ok else 'FAIL'}  "
                        f"年份={'OK ' if result.hall_clean else 'FAIL'}  "
                        f"{result.latency_ms:.0f}ms")
                with _PRINT_LOCK:
                    print(line, flush=True)

    results = [results_by_id[case.id] for case in cases if case.id in results_by_id]

    module = getattr(invoker, "module", None)
    if not args.no_render and module is not None:
        for result in results:
            if not result.parse_ok:
                result.render_error = "未解析出 JSON，跳过"
                continue
            result.render_error = check_render(module, result.parsed_data)
            result.render_checked = True
    return results


def _failed_result(case: Case, error: str) -> CaseResult:
    return CaseResult(
        id=case.id, title=case.title, parse_ok=False, mode="failed", strict_ok=False,
        required_ok=False, required_fields=case.required_fields,
        field_present={f: False for f in REQUIRED_TOP_FIELDS},
        missing=["<执行异常>"], placeholders=[], extra_years=[], extra_orgs=[],
        allowed_years=[], output_years=[], hall_score=100.0,
        name_accuracy=False if case.expect_name else None, latency_ms=0.0,
        error=error, render_error="未执行渲染", quality=collect_quality(None),
        raw_excerpt="",
    )


# ======================================================================
# 5. 汇总与输出
# ======================================================================
def summarize(results: list[CaseResult], cases: list[Case], args: argparse.Namespace,
              invoker_kind: str, elapsed_sec: float) -> dict[str, Any]:
    total = len(results)
    parsed = [r for r in results if r.parse_ok]
    parsed_n = len(parsed)

    strict_n = sum(1 for r in results if r.strict_ok)
    repaired_n = sum(1 for r in results if r.parse_ok and not r.strict_ok)
    parse_n = parsed_n

    field_rates = {
        name: (sum(1 for r in parsed if r.field_present.get(name)) / parsed_n if parsed_n else 0.0)
        for name in REQUIRED_TOP_FIELDS
    }
    required_rate = (sum(1 for r in parsed if r.required_ok) / parsed_n) if parsed_n else 0.0
    name_cases = [r for r in parsed if r.name_accuracy is not None]
    name_accuracy_rate = (
        sum(1 for r in name_cases if r.name_accuracy) / len(name_cases) if name_cases else None
    )

    hall_clean_rate = (sum(1 for r in parsed if r.hall_clean) / parsed_n) if parsed_n else 0.0
    avg_hall_score = (sum(r.hall_score for r in parsed) / parsed_n) if parsed_n else 0.0
    extra_year_total = sum(len(r.extra_years) for r in parsed)

    render_checked = [r for r in results if r.render_checked]
    render_rate = (
        sum(1 for r in render_checked if r.render_error is None) / len(render_checked)
        if render_checked else None
    )

    latencies = [r.latency_ms for r in results if r.latency_ms > 0]
    quality_bullets = [r.quality.get("bullets", 0) for r in parsed]
    placeholders = [r.quality.get("placeholders", 0) for r in parsed]

    json_rate = parse_n / total if total else 0.0
    composite = (0.3 * json_rate + 0.4 * required_rate + 0.3 * hall_clean_rate) * 100

    failures = []
    for result, case in zip(results, cases):
        reasons: list[str] = []
        if not result.parse_ok:
            reasons.append(f"JSON 解析失败：{result.error or '空输出'}")
        else:
            lacking = [f for f in case.required_fields if not result.field_present.get(f)]
            if lacking:
                reasons.append("缺少必填字段：" + "、".join(lacking))
            if result.extra_years:
                reasons.append("出现素材中没有的年份：" + "、".join(result.extra_years))
        if result.render_checked and result.render_error:
            reasons.append("Word 渲染失败：" + result.render_error)
        if reasons:
            failures.append({"id": case.id, "title": case.title, "reasons": reasons})

    return {
        "meta": {
            "eval_version": VERSION,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "adapter": invoker_kind,
            "model": None if invoker_kind == "mock" else args.model,
            "temperature": None if invoker_kind == "mock" else args.temperature,
            "cases_file": Path(args.cases).name,
            "n_cases": total,
            "workers": args.workers,
            "retries": args.retries,
            "year_penalty": args.year_penalty,
            "elapsed_sec": round(elapsed_sec, 2),
        },
        "metrics": {
            "json_parse_rate": json_rate,
            "strict_json_rate": strict_n / total if total else 0.0,
            "repaired_json_rate": repaired_n / total if total else 0.0,
            "required_field_complete_rate": required_rate,
            "field_rates": {k: round(v, 4) for k, v in field_rates.items()},
            "name_accuracy_rate": name_accuracy_rate,
            "hallucination_free_rate": hall_clean_rate,
            "avg_hallucination_score": avg_hall_score,
            "extra_year_total": extra_year_total,
            "composite_score": composite,
            "error_rate": sum(1 for r in results if r.error) / total if total else 0.0,
            "render_ok_rate": render_rate,
            "render_checked": len(render_checked),
            "avg_placeholders": (sum(placeholders) / parsed_n) if parsed_n else 0.0,
            "avg_bullets": (sum(quality_bullets) / parsed_n) if parsed_n else 0.0,
            "long_bullet_total": sum(r.quality.get("long_bullets", 0) for r in parsed),
            "missing_leadin_total": sum(r.quality.get("missing_leadin", 0) for r in parsed),
            "invalid_period_total": sum(len(r.quality.get("invalid_periods", [])) for r in parsed),
            "unknown_section_total": sum(len(r.quality.get("unknown_sections", [])) for r in parsed),
            "latency_ms_p50": _percentile(latencies, 50),
            "latency_ms_p95": _percentile(latencies, 95),
        },
        "failures": failures,
        "cases": [r.to_row() for r in results],
    }


def _render_markdown(summary: dict[str, Any], results: list[CaseResult],
                     cases: list[Case]) -> str:
    meta = summary["meta"]
    metrics = summary["metrics"]
    case_by_id = {case.id: case for case in cases}

    lines: list[str] = []
    lines.append("# 简历生成评测报告")
    lines.append("")
    lines.append(f"- 时间：{meta['timestamp']}")
    lines.append(f"- 用例：{meta['n_cases']} 条（`{meta['cases_file']}`）")
    lines.append(f"- 适配器：`{meta['adapter']}`" + (
        "（真实调用 web_app.generate_resume_data）" if meta["adapter"] == "inprocess" else ""))
    if meta["adapter"] != "mock":
        lines.append(f"- 模型：`{meta['model']}`，temperature={meta['temperature']}")
    else:
        lines.append("- 模型：离线假模型（mock，仅用于验证评测脚本本身）")
    lines.append(f"- 总耗时：{meta['elapsed_sec']}s")
    lines.append("")

    lines.append("## 总览")
    lines.append("")
    lines.append("| 指标 | 数值 | 说明 |")
    lines.append("| --- | --- | --- |")
    lines.append(f"| JSON 解析成功率 | **{_pct(metrics['json_parse_rate'])}** | 分母=全部用例 |")
    lines.append(f"| 　严格 JSON（无代码块/无解释） | {_pct(metrics['strict_json_rate'])} | 原文直接 json.loads 成功 |")
    lines.append(f"| 　经修复后解析成功 | {_pct(metrics['repaired_json_rate'])} | 去代码块 / 抠括号找回 |")
    lines.append(f"| 必填字段完整率 | **{_pct(metrics['required_field_complete_rate'])}** | 分母=解析成功的用例 |")
    for name in REQUIRED_TOP_FIELDS:
        lines.append(f"| 　其中 {name} | {_pct(metrics['field_rates'].get(name))} | 单个字段出现率 |")
    if metrics["name_accuracy_rate"] is not None:
        lines.append(f"| 　姓名与素材一致 | {_pct(metrics['name_accuracy_rate'])} | 仅在素材给了姓名的用例上统计 |")
    lines.append(f"| 年份幻觉免疫率 | **{_pct(metrics['hallucination_free_rate'])}** | 分母=解析成功的用例 |")
    lines.append(f"| 　平均幻觉得分 | {metrics['avg_hallucination_score']:.1f} / 100 | 每多一个年份扣 {meta['year_penalty']:g} 分 |")
    lines.append(f"| 　多出年份总次数 | {metrics['extra_year_total']} | 越低越好 |")
    lines.append(f"| 综合得分 | **{metrics['composite_score']:.1f} / 100** | 0.3×JSON + 0.4×字段 + 0.3×幻觉 |")
    lines.append(f"| Word 渲染成功率 | {_pct(metrics['render_ok_rate'])} | normalize + build_resume_bytes 实跑 |")
    lines.append(f"| 调用出错率 | {_pct(metrics['error_rate'])} | 异常/超时/空输出 |")
    if metrics["latency_ms_p50"] is not None:
        lines.append(f"| 延迟 P50 / P95 | {metrics['latency_ms_p50']:.0f}ms / {metrics['latency_ms_p95']:.0f}ms | 端到端单条 |")
    lines.append("")

    lines.append("## 逐条明细")
    lines.append("")
    lines.append("| id | 解析 | 必填 | 姓名 | 年份幻觉 | 幻觉分 | 用时 | 备注 |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- | --- |")
    for result in results:
        case = case_by_id.get(result.id)
        required = case.required_fields if case else REQUIRED_TOP_FIELDS
        note = result.mode if result.mode not in ("strict",) else ""
        if result.error:
            note = result.error[:40].replace("|", "/")
        elif result.extra_years:
            note = "多出 " + "、".join(result.extra_years)
        else:
            lacking = [name for name in required if not result.field_present.get(name)]
            if lacking:
                note = "缺 " + "、".join(lacking)[:40]
        lines.append(
            "| {id} | {parse} | {req} | {name} | {hall} | {score:.0f} | {ms:.0f}ms | {note} |".format(
                id=result.id,
                parse=_mark(result.parse_ok),
                req=_mark(result.required_ok),
                name=_mark(result.name_accuracy),
                hall=_mark(result.hall_clean),
                score=result.hall_score,
                ms=result.latency_ms,
                note=note,
            )
        )
    lines.append("")

    if summary["failures"]:
        lines.append("## 需要修的问题")
        lines.append("")
        for item in summary["failures"]:
            title = f"（{item['title']}）" if item["title"] else ""
            lines.append(f"- **{item['id']}**{title}")
            for reason in item["reasons"]:
                lines.append(f"  - {reason}")
        lines.append("")

    lines.append("## 指标口径（改口径就改这里）")
    lines.append("")
    lines.append("- **JSON 解析成功率**：`strict` = `json.loads(原文)` 成功；`fenced` = 去掉 ```json 包裹后成功；")
    lines.append("  `salvaged` = 括号配对/贪婪正则从混杂文本里找回对象。三者任一成功即算解析成功。")
    lines.append("- **必填字段完整率**：`name` / `intent` / `contact` / `sections` 四个顶层字段非空且不是纯占位符；")
    lines.append("  `sections` 还要求每个 entry 带齐 `period/org/role/title/bullets` 五个键，且整份简历至少有一条要点。")
    lines.append("  用例可用 `required_fields` 声明本条真正必需的字段（素材本来没有姓名时就不要求 name）。")
    lines.append("- **年份幻觉**：正则 `(?<!\\d)(?:19|20)\\d{2}(?!\\d)` 抓输出里的年份，凡不在素材年份集合里的都算多出，")
    lines.append(f"  每个扣 {meta['year_penalty']:g} 分，单条得分 = `max(0, 100 - 扣分)`；幻觉免疫率 = 无多出年份的用例占比。")
    lines.append("- **辅助指标**（不计入综合分）：占位符数量、要点条数、超过 80 字的要点、缺少 `**领起词**` 的要点、")
    lines.append("  非法月份（如 `2022.13`）、组织名未在素材出现（误报率高，仅提示）。")
    lines.append("")
    lines.append("## 复现")
    lines.append("")
    lines.append("```bash")
    lines.append("python eval.py --limit 3          # 先试 3 条")
    lines.append("python eval.py                    # 全量")
    lines.append("python eval.py --adapter mock      # 不花钱自测评测器")
    lines.append("```")
    lines.append("")
    return "\n".join(lines)


def write_outputs(out_dir: Path, summary: dict[str, Any], results: list[CaseResult],
                  args: argparse.Namespace, cases: list[Case]) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)

    rows = [json.dumps(result.to_row(), ensure_ascii=False) for result in results]
    (out_dir / "results.jsonl").write_text("\n".join(rows) + "\n", encoding="utf-8")
    (out_dir / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_dir / "report.md").write_text(
        _render_markdown(summary, results, cases), encoding="utf-8"
    )

    if args.dump_raw:
        raw_dir = out_dir / "raw"
        raw_dir.mkdir(exist_ok=True)
        for result in results:
            content = result.raw_full or result.raw_excerpt
            (raw_dir / f"{result.id}.txt").write_text(content, encoding="utf-8")


def print_report(summary: dict[str, Any], out_dir: Path, args: argparse.Namespace) -> None:
    metrics = summary["metrics"]
    meta = summary["meta"]
    print()
    print("=" * 62)
    print(" 评测汇总")
    print("=" * 62)
    print(f" 用例 {meta['n_cases']} 条 | 适配器 {meta['adapter']} | 耗时 {meta['elapsed_sec']}s")
    print("-" * 62)
    print(f" JSON 解析成功率     {_pct(metrics['json_parse_rate']):>7}   "
          f"（严格 {_pct(metrics['strict_json_rate'])} / 修复 {_pct(metrics['repaired_json_rate'])}）")
    print(f" 必填字段完整率      {_pct(metrics['required_field_complete_rate']):>7}   "
          f"name {_pct(metrics['field_rates']['name'])} · intent {_pct(metrics['field_rates']['intent'])} · "
          f"contact {_pct(metrics['field_rates']['contact'])} · sections {_pct(metrics['field_rates']['sections'])}")
    print(f" 年份幻觉免疫率      {_pct(metrics['hallucination_free_rate']):>7}   "
          f"平均分 {metrics['avg_hallucination_score']:.1f} · 多出 {metrics['extra_year_total']} 处")
    print(f" 综合得分            {metrics['composite_score']:>6.1f} / 100")
    print(f" Word 渲染成功率     {_pct(metrics['render_ok_rate']):>7}   "
          f"出错率 {_pct(metrics['error_rate'])}")
    if metrics["latency_ms_p50"] is not None:
        print(f" 延迟 P50 / P95      {metrics['latency_ms_p50']:.0f}ms / {metrics['latency_ms_p95']:.0f}ms")
    print("-" * 62)
    print(f" 报告：{out_dir / 'report.md'}")
    print(f" 明细：{out_dir / 'results.jsonl'}")
    print()

    if summary["failures"] and args.show_failures:
        print(" 失败用例：")
        for item in summary["failures"]:
            print(f"  - {item['id']}: " + "；".join(item["reasons"]))
        print()

    # 阈值检查，方便接 CI
    checks = [
        ("min_json_rate", metrics["json_parse_rate"], "JSON 解析成功率"),
        ("min_field_rate", metrics["required_field_complete_rate"], "必填字段完整率"),
        ("min_hallucination_free_rate", metrics["hallucination_free_rate"], "年份幻觉免疫率"),
    ]
    failed = False
    for attr, value, label in checks:
        threshold = getattr(args, attr, None)
        if threshold is None:
            continue
        if value < threshold:
            print(f" [阈值未通过] {label} {_pct(value)} < {_pct(threshold)}")
            failed = True
    if failed:
        raise SystemExit(1)


# ======================================================================
# 6. CLI
# ======================================================================
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="简历生成链路评测：JSON 解析率 / 必填字段完整率 / 年份幻觉检测",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--cases", type=Path, default=DEFAULT_CASES, help="用例 JSONL 路径")
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR, help="报告输出目录")
    parser.add_argument("--adapter", choices=("auto", "inprocess", "api", "cli", "mock"),
                        default="auto", help="怎么调用生成脚本")
    parser.add_argument("--module", default=DEFAULT_MODULE, help="生成逻辑所在模块（相对 eval.py 或 cwd）")
    parser.add_argument("--cmd", default="", help="cli 适配器的命令，素材走 stdin；可用 {input_file} 占位")
    parser.add_argument("--api-key-env", default="DEEPSEEK_API_KEY", help="读取 Key 的环境变量名")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="OpenAI 兼容接口地址")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="模型名")
    parser.add_argument("--temperature", type=float, default=DEFAULT_TEMPERATURE, help="采样温度")
    parser.add_argument("--max-tokens", type=int, default=DEFAULT_MAX_TOKENS, help="最大输出 token")
    parser.add_argument("--limit", type=int, default=0, help="只跑前 N 条（0=全部）")
    parser.add_argument("--workers", type=int, default=4, help="并发数")
    parser.add_argument("--retries", type=int, default=2, help="失败/限流重试次数")
    parser.add_argument("--timeout", type=float, default=180.0, help="cli 适配器单次超时（秒）")
    parser.add_argument("--year-penalty", type=float, default=DEFAULT_YEAR_PENALTY,
                        help="每多出一个年份扣的分")
    parser.add_argument("--no-render", action="store_true", help="跳过 docx 渲染检查")
    parser.add_argument("--dump-raw", action="store_true", help="保存每条完整原始输出到 raw/")
    parser.add_argument("--show-failures", action="store_true", help="控制台列出失败用例")
    parser.add_argument("--dry-run", action="store_true", help="只校验用例/参数/适配器，不发请求")
    parser.add_argument("--quiet", action="store_true", help="不打印逐条进度")
    parser.add_argument("--min-json-rate", type=float, default=None,
                        help="低于该比例（0~1）则退出码 1")
    parser.add_argument("--min-field-rate", type=float, default=None,
                        help="低于该比例（0~1）则退出码 1")
    parser.add_argument("--min-hallucination-free-rate", type=float, default=None,
                        help="低于该比例（0~1）则退出码 1")
    return parser


def build_invoker(args: argparse.Namespace, cases: list[Case]) -> Any:
    adapter = args.adapter
    module_path = _resolve_module_path(args.module)

    if adapter == "mock":
        return MockInvoker()

    if adapter == "cli":
        if not args.cmd.strip():
            raise SystemExit("[adapter] --adapter cli 需要同时提供 --cmd")
        return CliInvoker(args.cmd, args.timeout, Path.cwd())

    api_key, source = _resolve_api_key(args, module_path.parent)
    if not api_key:
        raise SystemExit(
            "[adapter] 找不到 DeepSeek API Key：请设置环境变量 DEEPSEEK_API_KEY，"
            "或写入 .streamlit/secrets.toml；也可以先用 --adapter mock 离线自测"
        )
    if adapter == "auto":
        has_generator = module_path.exists() and re.search(
            r"def\s+generate_resume_data\s*\(", module_path.read_text(encoding="utf-8", errors="replace")
        )
        adapter = "inprocess" if has_generator else "api"

    if adapter == "inprocess":
        print(f"[adapter] inprocess：{module_path.name} · key 来自 {source}", file=sys.stderr)
        return InProcessInvoker(module_path, api_key, args.base_url, args.model,
                                args.temperature, args.max_tokens)

    print(f"[adapter] api：{args.model} · key 来自 {source}", file=sys.stderr)
    return ApiInvoker(module_path if module_path.exists() else None, api_key, args.base_url,
                      args.model, args.temperature, args.max_tokens)


def main(argv: list[str] | None = None) -> int:
    _setup_stdout()
    parser = build_parser()
    args = parser.parse_args(argv)

    started = time.perf_counter()
    cases = load_cases(args.cases, limit=args.limit or None)

    if args.dry_run:
        print(f"[cases] {len(cases)} 条用例：{args.cases}")
        for case in cases:
            extra = []
            if case.expect_name:
                extra.append(f"name={case.expect_name}")
            extra.append("years=" + (",".join(sorted(case.allowed_year_set, key=int)) or "无"))
            extra.append("required=" + "/".join(case.required_fields))
            print(f"  - {case.id:<4} {case.title or '（无标题）'}  [{' | '.join(extra)}]")
        try:
            invoker = build_invoker(args, cases)
            print(f"[adapter] 可用：{invoker.kind}")
        except SystemExit as exc:
            print(f"[adapter] 不可用：{exc}", file=sys.stderr)
            return 1
        print("[dry-run] 未发送任何请求。")
        return 0

    invoker = build_invoker(args, cases)
    results = run_all(invoker, cases, args)
    elapsed = time.perf_counter() - started
    summary = summarize(results, cases, args, invoker.kind, elapsed)
    write_outputs(args.out_dir, summary, results, args, cases)
    print_report(summary, args.out_dir, args)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
