from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROGRAM_DIR = Path(__file__).resolve().parent
if str(PROGRAM_DIR) not in sys.path:
    sys.path.insert(0, str(PROGRAM_DIR))

from paper_collage_core import (  # noqa: E402
    InteractiveLoginConfig,
    MissingPdfError,
    load_sources_from_input_file,
    run_collage_pipeline,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Create a real-paper collage from a CSV or TXT file of article URLs."
    )
    parser.add_argument("--input", required=True, help="CSV or TXT file containing paper website URLs.")
    parser.add_argument("--output", default="output/real_paper_cover_collage_16x9.png")
    parser.add_argument("--pdf-output", default=None)
    parser.add_argument("--pptx-output", default=None)
    parser.add_argument("--cache-dir", default="pdf_cache")
    parser.add_argument("--image-dir", default="page1_cache")
    parser.add_argument("--canvas-width", type=int, default=1920)
    parser.add_argument("--canvas-height", type=int, default=1080)
    parser.add_argument("--page-width", type=int, default=390)
    parser.add_argument("--page-height", type=int, default=520)
    parser.add_argument("--gap-x", type=int, default=70)
    parser.add_argument("--gap-y", type=int, default=52)
    parser.add_argument("--background", default="white")
    parser.add_argument("--allow-missing", action="store_true")
    parser.add_argument("--skip-download", action="store_true")
    parser.add_argument("--timeout", type=int, default=45)
    parser.add_argument("--interactive-login", action="store_true")
    parser.add_argument("--login-profile-dir", default=".playwright_login_profile")
    parser.add_argument("--login-timeout", type=int, default=900)
    parser.add_argument("--login-poll-interval", type=int, default=10)
    parser.epilog = (
        "Accepted input formats:\n"
        "  CSV: requires article_url/url/paper_url/website column; supports optional refs, slug, pdf_url\n"
        "  TXT: one paper URL per non-empty line; lines starting with # are ignored"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    input_path = Path(args.input)
    sources = load_sources_from_input_file(input_path)
    interactive_login = InteractiveLoginConfig(
        enabled=args.interactive_login,
        profile_dir=Path(args.login_profile_dir),
        timeout_seconds=args.login_timeout,
        poll_interval_seconds=args.login_poll_interval,
    )
    return run_collage_pipeline(
        sources,
        cache_dir=Path(args.cache_dir),
        image_dir=Path(args.image_dir),
        output_path=Path(args.output),
        pdf_output_path=Path(args.pdf_output) if args.pdf_output else Path(args.output).with_suffix(".pdf"),
        pptx_output_path=Path(args.pptx_output) if args.pptx_output else None,
        canvas_width=args.canvas_width,
        canvas_height=args.canvas_height,
        page_width=args.page_width,
        page_height=args.page_height,
        gap_x=args.gap_x,
        gap_y=args.gap_y,
        background=args.background,
        allow_missing=args.allow_missing,
        skip_download=args.skip_download,
        timeout=args.timeout,
        interactive_login=interactive_login,
    )


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except MissingPdfError as exc:
        for source in exc.missing:
            print(f"- {source.slug}: {source.article_url}", file=sys.stderr)
        raise SystemExit(1)
