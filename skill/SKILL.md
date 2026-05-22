---
name: paper-kaleidoscope-workflow
description: Use when a user wants an end-to-end Paper Kaleidoscope workflow for searching papers, selecting a final list, downloading real article PDFs, and generating a PPT-ready collage or editable PPTX from those papers.
---

# Paper Kaleidoscope Workflow

## Overview

Guide the user from topic discovery to a finished paper-first-page collage. The workflow must preserve one hard rule: use real paper PDFs only, never fabricate a cover page.

## When to Use

- User wants "论文万花筒"
- User wants "搜索文献然后做论文首页拼图"
- User wants help turning a research topic into a paper collage
- User needs an agent-guided flow from paper search to final PNG/PDF/PPTX
- User wants a CSV/TXT input file prepared and then run through Paper Kaleidoscope

Do not use this skill if the user already has the final paper list and only wants a tiny one-off script tweak.

## Workflow

1. Confirm the research topic and intended output style.
2. Search and summarize candidate papers from primary or publisher pages.
3. Ask the user to choose the final paper list before any collage run.
4. Write the selected paper URLs into either:
   - `program/sample_paper_sites.csv` style CSV, or
   - one-URL-per-line TXT
5. Run:

```powershell
python program/run_paper_collage_from_sites.py `
  --input <papers.csv-or-txt> `
  --output output/paper_kaleidoscope_16x9.png `
  --pdf-output output/paper_kaleidoscope_16x9.pdf `
  --pptx-output output/paper_kaleidoscope_16x9_editable.pptx
```

6. If institution login blocks download, rerun with:

```powershell
python program/run_paper_collage_from_sites.py `
  --input <papers.csv-or-txt> `
  --interactive-login `
  --login-timeout 1200 `
  --output output/paper_kaleidoscope_16x9.png
```

7. Deliver:
   - final PNG
   - optional PDF
   - optional editable PPTX
   - missing PDF report if anything still fails

## Input File Rules

### CSV

Required column set: one of `article_url`, `url`, `paper_url`, `website`

Optional columns:

- `refs`
- `slug`
- `pdf_url`

### TXT

- one paper URL per line
- ignore blank lines
- ignore lines starting with `#`

## Non-Negotiables

- Real paper PDF first page only
- No AI-generated paper covers
- No placeholder pages
- If a download fails, either:
  - use interactive login, or
  - stop and surface the missing paper list

## Recommended Agent Behavior

- When searching papers, cite the paper title and source URL clearly
- Let the user choose the final subset; do not silently decide the final reading list
- Prefer publisher or PMC links over secondary mirrors
- Keep the final input file under the project workspace so the user can review it
