"""Master/layout placeholder filtering."""


# Default placeholder texts in various languages
DEFAULT_PLACEHOLDER_TEXTS = {
    # English
    "click to add title",
    "click to add subtitle",
    "click to add text",
    "click to add footer",
    "click to add header",
    # Chinese
    "单击此处添加标题",
    "单击此处添加副标题",
    "单击此处添加文本",
    "单击此处添加页脚",
    "单击此处添加页眉",
    # Common variations
    "title",
    "subtitle",
    "text",
    "enter title here",
    "enter subtitle here",
    "enter text here",
}


def is_default_placeholder(text):
    """Check if text is a default placeholder that should be filtered.

    Args:
        text: The text content.

    Returns:
        bool: True if text is a default placeholder.
    """
    if not text:
        return False
    normalized = text.strip().lower()
    return normalized in DEFAULT_PLACEHOLDER_TEXTS


def filter_shape_text(shape):
    """Filter out default placeholder text from a shape.

    Args:
        shape: python-pptx Shape object.

    Returns:
        tuple of (should_skip, filtered_text).
    """
    if not shape.has_text_frame:
        return False, ""

    texts = []
    for para in shape.text_frame.paragraphs:
        text = para.text.strip()
        if text and not is_default_placeholder(text):
            texts.append(text)

    if not texts:
        return True, ""

    return False, "\n".join(texts)


def should_skip_shape(shape):
    """Determine if a shape should be skipped entirely.

    Args:
        shape: python-pptx Shape object.

    Returns:
        bool: True if shape should be skipped.
    """
    if not shape.has_text_frame:
        return False

    for para in shape.text_frame.paragraphs:
        text = para.text.strip()
        if text and not is_default_placeholder(text):
            return False

    return True


def is_empty_slide(slide):
    """Check if a slide has no meaningful content.

    A slide is considered empty if it has no text, images, tables, or charts.

    Args:
        slide: python-pptx Slide object.

    Returns:
        bool: True if slide is empty.
    """
    for shape in slide.shapes:
        # Check for images
        if hasattr(shape, "image") and shape.image is not None:
            return False

        # Check for tables
        if hasattr(shape, "has_table") and shape.has_table:
            return False

        # Check for charts
        if hasattr(shape, "has_chart") and shape.has_chart:
            return False

        # Check for meaningful text (not just default placeholders)
        if shape.has_text_frame:
            skip, _ = filter_shape_text(shape)
            if not skip:
                return False

    return True
