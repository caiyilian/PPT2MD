"""Batch conversion of multiple PPTX files."""

import os
from pathlib import Path


def find_pptx_files(input_path):
    """Find all .pptx files in a directory.

    Args:
        input_path: Path to directory.

    Returns:
        list of Path objects.
    """
    if input_path.is_file():
        if input_path.suffix.lower() == ".pptx":
            return [input_path]
        return []

    pptx_files = []
    for item in sorted(input_path.iterdir()):
        if item.is_file() and item.suffix.lower() == ".pptx":
            pptx_files.append(item)
    return pptx_files


def get_output_dir_for_file(pptx_path, output_base=None):
    """Get the output directory for a single PPTX file.

    Args:
        pptx_path: Path to the PPTX file.
        output_base: Base output directory.

    Returns:
        Path: Output directory for the markdown and images.
    """
    if output_base:
        base = Path(output_base)
    else:
        base = pptx_path.parent

    stem = pptx_path.stem
    return base / "{}_output".format(stem)


def batch_convert_summary(results):
    """Generate a summary report of batch conversion.

    Args:
        results: list of dict with file, success, error info.

    Returns:
        str: Summary report.
    """
    total = len(results)
    success = sum(1 for r in results if r.get("success"))
    failed = total - success

    lines = [
        "Batch conversion complete:",
        "  Total: {}".format(total),
        "  Success: {}".format(success),
        "  Failed: {}".format(failed),
    ]

    if failed > 0:
        lines.append("\nFailed files:")
        for r in results:
            if not r.get("success"):
                lines.append("  - {}: {}".format(r["file"], r.get("error", "Unknown error")))

    return "\n".join(lines)
