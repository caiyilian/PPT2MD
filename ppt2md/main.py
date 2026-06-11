"""CLI entry point for ppt2md."""

import argparse
import sys
from pathlib import Path

from ppt2md.parser.presentation import open_presentation, get_slide_count


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="ppt2md",
        description="Convert PowerPoint (.pptx) files to Markdown format.",
    )
    parser.add_argument(
        "input",
        help="Path to the input .pptx file or directory containing .pptx files",
    )
    parser.add_argument(
        "-o", "--output",
        help="Output directory (default: same as input file)",
    )
    parser.add_argument(
        "--include-notes",
        action="store_true",
        help="Include speaker notes in output",
    )
    parser.add_argument(
        "--include-hidden",
        action="store_true",
        help="Include hidden slides",
    )
    parser.add_argument(
        "--skip-empty",
        action="store_true",
        help="Skip empty slides",
    )
    parser.add_argument(
        "--formula-as-image",
        action="store_true",
        help="Export formulas as images instead of LaTeX",
    )
    parser.add_argument(
        "--keep-original-format",
        action="store_true",
        help="Keep original formatting info in comments",
    )
    parser.add_argument(
        "--image-dpi",
        type=int,
        default=96,
        help="DPI for image export (default: 96)",
    )
    parser.add_argument(
        "--no-frontmatter",
        action="store_true",
        help="Skip generating YAML frontmatter",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose output",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug output",
    )

    args = parser.parse_args(argv)
    input_path = Path(args.input)

    if not input_path.exists():
        print("Error: {} does not exist".format(input_path), file=sys.stderr)
        return 1

    if input_path.suffix.lower() != ".pptx":
        print("Error: {} is not a .pptx file".format(input_path), file=sys.stderr)
        return 1

    prs = open_presentation(str(input_path))
    count = get_slide_count(prs)
    print("ppt2md: {} contains {} slide(s)".format(input_path.name, count))
    return 0


if __name__ == "__main__":
    sys.exit(main())
