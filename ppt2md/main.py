"""CLI entry point for ppt2md."""

import argparse
import sys


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
    print(f"ppt2md: Converting {args.input} to Markdown")
    return 0


if __name__ == "__main__":
    sys.exit(main())
