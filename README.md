# 论文万花筒 / Paper Kaleidoscope

中文 | [English](#english)

## 中文

**论文万花筒（Paper Kaleidoscope）** 是一个长期维护的双入口项目，用来把真实论文 PDF 的第一页整理成适合展示、汇报和 PPT 使用的论文首页排布图。

项目坚持一个边界：只使用真实论文 PDF 的第一页，不生成、不仿造、不用占位封面。

- `program/`：面向普通用户的 Python 程序入口。输入论文网站链接列表，输出真实论文首页拼图。
- `skill/`：面向 agent 的 workflow skill。用于引导“搜索相关文献 -> 用户确认最终文献 -> 自动生成排布图”的完整流程。

### 仓库结构

- `program/run_paper_collage_from_sites.py`：程序入口
- `program/paper_collage_core.py`：核心逻辑
- `program/requirements.txt`：程序依赖
- `program/sample_paper_sites.csv`：CSV 示例输入
- `program/sample_paper_sites.txt`：TXT 示例输入
- `skill/SKILL.md`：skill 入口
- `skill/templates/`：供 agent 或用户快速填写的输入模板

### 程序输入

支持两种输入格式。

#### CSV

至少包含以下任一列：

- `article_url`
- `url`
- `paper_url`
- `website`

可选列：

- `refs`
- `slug`
- `pdf_url`

#### TXT

- 每行一个论文页面链接
- 空行会被忽略
- 以 `#` 开头的行会被视为注释

### 安装

```powershell
python -m pip install -r program/requirements.txt
```

如果目标站点需要机构登录，可额外安装：

```powershell
python -m pip install playwright
playwright install chromium
```

### 运行示例

```powershell
python program/run_paper_collage_from_sites.py `
  --input program/sample_paper_sites.csv `
  --output output/paper_kaleidoscope_16x9.png `
  --pdf-output output/paper_kaleidoscope_16x9.pdf `
  --pptx-output output/paper_kaleidoscope_16x9_editable.pptx
```

TXT 用法相同：

```powershell
python program/run_paper_collage_from_sites.py `
  --input program/sample_paper_sites.txt `
  --output output/paper_kaleidoscope_16x9.png
```

### 输出

- 扁平拼图 PNG
- 扁平拼图 PDF
- 可拖动编辑的 PPTX
- 若缺失真实 PDF，则输出 `missing_pdfs.txt`

### Program / Skill 区分

如果你只是想直接跑一个程序，请使用 `program/`。

如果你希望 agent 帮你完成从选文献到生成拼图的整套流程，请使用 `skill/`，并让 agent 按 `skill/SKILL.md` 执行。

## English

**Paper Kaleidoscope** is a long-term dual-entry project for turning real paper PDF first pages into PPT-ready visual collages for talks, reports, and research presentations.

The project keeps one strict boundary: it only uses real first pages from real paper PDFs. It does not generate fake covers, imitate paper pages, or create placeholders.

- `program/`: the standalone Python program for end users. It takes a list of paper website URLs and generates a collage from real first-page PDFs.
- `skill/`: the agent workflow skill for the full flow from literature search to final collage generation.

### Repository layout

- `program/run_paper_collage_from_sites.py`: program entry point
- `program/paper_collage_core.py`: core implementation
- `program/requirements.txt`: program dependencies
- `program/sample_paper_sites.csv`: sample CSV input
- `program/sample_paper_sites.txt`: sample TXT input
- `skill/SKILL.md`: skill entry point
- `skill/templates/`: reusable templates for agent or user input files

### Program input

Two input formats are supported.

#### CSV

Requires at least one of:

- `article_url`
- `url`
- `paper_url`
- `website`

Optional columns:

- `refs`
- `slug`
- `pdf_url`

#### TXT

- one paper page URL per line
- blank lines are ignored
- lines starting with `#` are treated as comments

### Installation

```powershell
python -m pip install -r program/requirements.txt
```

If publisher access requires institution login, also install:

```powershell
python -m pip install playwright
playwright install chromium
```

### Usage

```powershell
python program/run_paper_collage_from_sites.py `
  --input program/sample_paper_sites.csv `
  --output output/paper_kaleidoscope_16x9.png `
  --pdf-output output/paper_kaleidoscope_16x9.pdf `
  --pptx-output output/paper_kaleidoscope_16x9_editable.pptx
```

TXT works the same way:

```powershell
python program/run_paper_collage_from_sites.py `
  --input program/sample_paper_sites.txt `
  --output output/paper_kaleidoscope_16x9.png
```

### Outputs

- flattened collage PNG
- flattened collage PDF
- editable PPTX with draggable paper-page objects
- `missing_pdfs.txt` if required PDFs are still unavailable

### Program vs. Skill

Use `program/` if you want a direct command-line workflow.

Use `skill/` if you want an agent-guided workflow from paper discovery to final collage output.
