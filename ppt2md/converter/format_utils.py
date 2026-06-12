"""Text formatting utilities for Markdown output."""

from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _get_baseline(run):
    """Get baseline shift from run XML (positive=superscript, negative=subscript)."""
    try:
        rPr = run._r.find("{{{}}}rPr".format(A_NS))
        if rPr is not None:
            baseline = rPr.get("baseline")
            if baseline is not None:
                return int(baseline)
    except Exception:
        pass
    return 0


def apply_inline_formatting(text, bold=False, italic=False, underline=False,
                             strikethrough=False, superscript=False, subscript=False):
    """Apply Markdown inline formatting to text."""
    if not text:
        return text
    if bold:
        text = "**{}**".format(text)
    if italic:
        text = "*{}*".format(text)
    if underline:
        text = "<u>{}</u>".format(text)
    if strikethrough:
        text = "~~{}~~".format(text)
    if superscript:
        text = "^{}".format(text)
    if subscript:
        text = "_{}".format(text)
    return text


def format_paragraph_runs(paragraph):
    """Convert a paragraph's runs to Markdown, preserving inline formatting.

    Handles bold, italic, underline, strikethrough, superscript, subscript.
    """
    parts = []
    for run in paragraph.runs:
        text = run.text
        if not text:
            continue
        bold = run.font.bold
        italic = run.font.italic
        underline = run.font.underline
        strikethrough = getattr(run.font, "strikethrough", None)

        baseline = _get_baseline(run)
        superscript = baseline > 0
        subscript = baseline < 0

        formatted = apply_inline_formatting(
            text,
            bold=bool(bold) if bold is not None else False,
            italic=bool(italic) if italic is not None else False,
            underline=bool(underline) if underline is not None else False,
            strikethrough=bool(strikethrough) if strikethrough is not None else False,
            superscript=superscript,
            subscript=subscript,
        )
        parts.append(formatted)
    return "".join(parts)
