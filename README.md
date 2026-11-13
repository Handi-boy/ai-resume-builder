# AI 简历生成器

把大白话经历交给 DeepSeek 整理成结构化内容，再排成可直接投递的 A4 Word 简历。

版式固定：深蓝色块标题 + 右侧横线、双栏信息表、时间左对齐 / 机构右对齐、加粗领起词。

## 特性

- 口语化素材 → 结构化 JSON → 一键下载 .docx
- 排版由程序控制，模型只负责写内容，不会把版式搞乱
- 缺失信息自动写成 `[占位符]`，界面提示补全
- 只使用用户提供的事实，不编造公司、时间、数字
- 内置超时与重试，Key 失效 / 余额不足有明确提示

## 本地运行

先 `pip install -r requirements.txt`，再 `streamlit run web_app.py`。

API Key 放在 `.streamlit/secrets.toml`（该文件已被 `.gitignore` 排除）：

`DEEPSEEK_API_KEY = "sk-..."`
