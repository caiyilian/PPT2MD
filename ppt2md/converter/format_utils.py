"""Text formatting utilities for Markdown output."""


def apply_inline_formatting(text, bold=False, italic=False, underline=False, strikethrough=False):
    """Apply Markdown inline formatting to text.

    Args:
        text: The text to format.
        bold: Wrap in ** for bold.
        italic: Wrap in * for italic.
        underline: Wrap in <u> for underline.
        strikethrough: Wrap in ~~ for strikethrough.

    Returns:
        Formatted text string.
    """
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
    return text


def format_paragraph_runs(paragraph):
    """Convert a paragraph's runs to Markdown, preserving inline formatting.

    Args:
        paragraph: python-pptx Paragraph object.

    Returns:
        Markdown-formatted string for the paragraph.
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
        formatted = apply_inline_formatting(
            text,
            bold=bool(bold) if bold is not None else False,
            italic=bool(italic) if italic is not None else False,
            underline=bool(underline) if underline is not None else False,
            strikethrough=bool(strikethrough) if strikethrough is not None else False,
        )
        parts.append(formatted)
    return "".join(parts)
