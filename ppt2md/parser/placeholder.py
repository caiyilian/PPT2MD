"""Placeholder type detection and Markdown mapping."""

from pptx.enum.shapes import PP_PLACEHOLDER


# Placeholder types to skip (metadata/footers)
SKIP_TYPES = {
    PP_PLACEHOLDER.SLIDE_NUMBER,
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.HEADER,
}


def get_placeholder_type(shape):
    """Determine the Markdown role of a placeholder shape.

    Args:
        shape: python-pptx Shape that is a placeholder.

    Returns:
        str: 'title', 'subtitle', 'body', 'object', or 'skip'.
    """
    if not shape.is_placeholder:
        return "object"

    ph_type = shape.placeholder_format.type

    if ph_type in SKIP_TYPES:
        return "skip"

    if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
        return "title"
    if ph_type == PP_PLACEHOLDER.SUBTITLE:
        return "subtitle"
    if ph_type == PP_PLACEHOLDER.BODY:
        return "body"

    return "object"


def placeholder_to_markdown_prefix(role):
    """Return the Markdown prefix for a placeholder role.

    Args:
        role: One of 'title', 'subtitle', 'body', 'object', 'skip'.

    Returns:
        str: Markdown prefix (e.g. '## ' for title).
    """
    if role == "title":
        return "## "
    if role == "subtitle":
        return "### "
    if role in ("body", "object"):
        return ""
    return ""
