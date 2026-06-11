"""List and bullet detection for Markdown output."""

import re


def is_ordered_list(text):
    """Check if text looks like an ordered list item.

    Args:
        text: Paragraph text.

    Returns:
        bool: True if text starts with a number followed by a dot.
    """
    return bool(re.match(r"^\d+\.\s", text))


def format_list_item(text, level=0):
    """Format text as a Markdown list item with proper indentation.

    Args:
        text: The list item text.
        level: Indentation level (0-based).

    Returns:
        str: Formatted list item.
    """
    indent = "  " * level
    return "{}- {}".format(indent, text)


def format_ordered_item(text, level=0):
    """Format text as a Markdown ordered list item.

    Args:
        text: The list item text (with number prefix stripped).
        level: Indentation level.

    Returns:
        str: Formatted ordered list item.
    """
    indent = "  " * level
    return "{}1. {}".format(indent, text)


def format_paragraph_as_list(text, level=0):
    """Detect if paragraph is a list item and format accordingly.

    Args:
        text: The paragraph text.
        level: Paragraph indentation level from python-pptx.

    Returns:
        tuple of (is_list, formatted_text).
    """
    if is_ordered_list(text):
        # Strip leading number and dot
        cleaned = re.sub(r"^\d+\.\s*", "", text)
        return True, format_ordered_item(cleaned, level)
    elif text.strip():
        return True, format_list_item(text, level)
    return False, text
