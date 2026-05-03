# Auto-Lab-Report-Agent

<div align="center">

**中山大学工科实验报告全自动生成智能体**

[![PyPI - Python Version](https://img.shields.io/badge/python-3.7%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](./LICENSE)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/lzh20060217/Auto-Lab-Report-Agent/pulls)

*Stop wrestling with Word formatting. Let AI craft your lab reports while you focus on the actual experiment.*

</div>

---

## 🔥 TL;DR

你还在手动调 Word 表格宽度、对齐、字体、行距吗？你还在为 Markdown 表格太小、打印后写不下数据而抓狂吗？

**把这个项目丢给大模型，上传讲义 PDF，一键生成排版完美的实验报告。**
- 3×2 隐形表格完美对齐基本信息
- 0.8cm 行高，手写数据再也不拥挤
- 6cm 图片占位框，贴示波器截图刚好
- AI 自动写总结，400+ 字学术深度

---

## 📸 效果预览

> 左侧为生成后的 Word 文档，右侧为 PDF 导出效果。排版直接拿去打印交作业。

| DOCX | PDF |
|------|-----|
| 基础信息隐形表格对齐 | 页眉页脚完整保留 |
| 表格行高 0.8cm 防拥挤 | 图片框 6cm 定格不跨页 |
| 黑体/宋体/Times New Roman 混排 | 页码居中 `- N -` 格式 |

---

## 🎯 痛点解决

### 痛点 1：Word 调格式调到凌晨三点

> 标题居中对齐 → 不小心偏了 → 用空格凑 → 打印出来歪了 → 重来。

本项目的 Python 脚本以 **像素级精度** 控制每个段落的字体、字号、缩进、对齐、行距。一次配置，永久复用。

### 痛点 2：Markdown 表格太小，手写根本写不下

> Markdown 生成的表格行高约 0.3cm，铅笔字写上去直接溢出到下一行。

本项目强制设置 **所有数据表格最小行高为 0.8cm**，相当于 Word 中的"至少 23pt"，给你充足的空白空间用笔填写实验数据。

### 痛点 3：图片占位留一整页空白，打印出来像缺页

> `[在此插入图片]` 占一整页 → 打印出来 → 助教："你是不是少交了一页？"

本项目使用 **1 行 × 1 列单格表格** 作为图片占位框，固定高度 6cm，灰色边框清晰标记粘贴范围，不会跨页留白，更不会破坏文档结构。

### 痛点 4：实验总结每次都要编，编得还很水

> "通过本次实验，我学到了很多知识……" ← 这种总结谁看谁尴尬。

本项目内置讲义知识库，AI 自动生成 **400 字以上**具有专业理论深度的实验总结，涵盖物理概念定性分析、参数工程折衷、误差反思和课程串联。

---

## 🏗️ 工作原理

```
┌──────────────┐     ┌─────────────────┐     ┌──────────────────┐
│ 实验讲义 PDF  │ ──▶ │ 大模型 (LLM)     │ ──▶ │ python-docx 排版  │
│ (RZ9653 等)  │     │ 提取&扩展内容    │     │ 引擎            │
└──────────────┘     └─────────────────┘     └────────┬─────────┘
                                                      │
                                              ┌───────▼─────────┐
                                              │  .docx + .pdf   │
                                              │  排版完美的报告  │
                                              └─────────────────┘
```

1. **内容层：** 大模型读取讲义 → 提取实验目的、原理、步骤 → 学术化扩展
2. **排版层：** `sysu_lab_skill.py` 提供所有排版函数（标题、正文、表格、图片框、页眉页脚）
3. **输出层：** `python-docx` 生成 `.docx`，`docx2pdf` 调用 Word COM 接口导出 `.pdf`

---

## 📂 项目结构

```
wordcode/
├── SYSU_Lab_Report_Skill.md     # ★ 核心 Skill 文件（给 Agent 的系统指令）
├── README.md                    # 本文件
├── requirements.txt             # Python 依赖
├── main.py                      # 命令行交互入口
├── generate_optimized.py        # 实验报告生成主脚本
│
├── skill/
│   ├── lab_report_skill.py      # 语义层 System Prompt
│   └── sysu_lab_skill.py        # ★ 排版铁律引擎（所有排版函数）
│
├── src/
│   ├── config.py                # 全局配置
│   ├── template.py              # 12 个实验的讲义数据库
│   ├── report_generator.py      # Markdown 报告生成器
│   └── converter.py             # Markdown → DOCX/PDF 转换器
│
├── presets/
│   └── experiment1.py           # 实验一预设
│
└── output/
    ├── 实验1_实验报告_promax版.docx
    ├── 实验1_实验报告_promax版.pdf
    ├── 实验1_实验报告_完美版.docx
    ├── 实验1_实验报告_完美版.pdf
    └── ...                      # 各版本示例输出
```

---

## 🚀 How to Use

### 方式一：Agent 对话模式（推荐）

1. 打开你喜欢的 AI Agent 工具（如 Trae、Cursor、ChatGPT with Code Interpreter）
2. 新建对话，上传你的实验讲义 PDF / PPTX
3. 艾特（@）`SYSU_Lab_Report_Skill.md` 文件
4. 告诉 Agent 你的实验编号和课程：

```
@SYSU_Lab_Report_Skill.md 
请根据我上传的高频电路实验讲义，生成实验一（小信号调谐放大器）的实验报告。
姓名：张三，学号：24350001
```

5. Agent 会自动调用 `python-docx` 生成排版完美的 `.docx` 和 `.pdf`

### 方式二：本地命令行模式

```bash
# 1. 克隆仓库
git clone https://github.com/lzh20060217/Auto-Lab-Report-Agent.git
cd Auto-Lab-Report-Agent/wordcode

# 2. 安装依赖
pip install -r requirements.txt

# 3. 生成报告
python generate_optimized.py

# 4. 打开 output/ 目录，直接打印交作业
```

### 方式三：指定实验编号

```bash
# 生成实验一，保存所有格式
python main.py --exp 1 --all

# 列出所有可用实验
python main.py --list

# 查看 Skill 信息
python main.py --skill
```

---

## 🧬 自定义适配

### 适配你自己的课程

1. 修改 `skill/sysu_lab_skill.py` 中的 `SYSULabReportSkill` 类常量
2. 在 `src/template.py` 中添加你课程的实验数据模板
3. 参照 `generate_optimized.py` 的写法，复制一份改成你的课程名

### 适配别的学校

1. 修改页眉文字 → `setup_header()`
2. 修改基础信息表格 → `add_info_table()`
3. 修改课程名/学院名 → `src/config.py`

---

## 📐 排版铁律速查

| 元素 | 要求 |
|------|------|
| 基础信息 | 3×2 隐形表格，五号 10.5pt 宋体，4:6 列宽 |
| 页眉 | 小五号 9pt 宋体 + 底部细横线 |
| 大标题 | 一号 26pt 宋体加粗居中 |
| 实验标题 | 小三号 15pt 黑体加粗居中，单行不折行 |
| 正文 | 小四 12pt 宋体，首行缩进 2 字符，1.5 倍行距 |
| 西文/数字 | Times New Roman |
| 数据表格行高 | ≥ 0.8cm (atLeast) |
| 图片占位框 | 1×1 表格，固定高度 6cm |
| 图注 | 五号 10.5pt，深灰色 #555555，居中 |
| 页脚 | 居中 `- PAGE -` 域代码 |

---

## 🔗 Links

- [python-docx 文档](https://python-docx.readthedocs.io/)
- [docx2pdf](https://pypi.org/project/docx2pdf/)
- [RZ9653 高频电子线路实验平台](http://www.njrz.com/)

---

## 📄 License

MIT © 2025 [lzh20060217](https://github.com/lzh20060217)

---

<div align="center">

**⭐ 如果这个项目帮你省了一晚上调格式的时间，请给个 Star！**

*Made with ❤️ by a fellow SYSU student who was tired of adjusting Word margins at 3 AM.*

</div>
