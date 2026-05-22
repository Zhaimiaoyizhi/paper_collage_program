# Program Entry

This directory contains the standalone Python program for Paper Kaleidoscope.

## Main entry

```powershell
python program/run_paper_collage_from_sites.py --help
```

## Inputs

- CSV with `article_url`/`url`/`paper_url`/`website`
- TXT with one paper page URL per line

## Outputs

- collage PNG
- collage PDF
- editable PPTX
- optional missing-PDF report

## Notes

- The program only uses real paper PDFs.
- It supports optional interactive institution login via Playwright.
