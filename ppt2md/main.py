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

    args = parser.parse_args(argv)
    input_path = Path(args.input)

    if not input_path.exists():
        print(f"Error: {input_path} does not exist", file=sys.stderr)
        return 1

    if input_path.suffix.lower() != ".pptx":
        print(f"Error: {input_path} is not a .pptx file", file=sys.stderr)
        return 1

    prs = open_presentation(str(input_path))
    count = get_slide_count(prs)
    print(f"ppt2md: {input_path.name} contains {count} slide(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
