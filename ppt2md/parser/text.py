"""Text frame extraction from PPTX slides."""


def extract_text_from_slide(slide):
    """Extract all text from a slide.

    Args:
        slide: python-pptx Slide object.

    Returns:
        list of dict with shape_index, text, and placeholder info.
    """
    results = []
    for i, shape in enumerate(slide.shapes):
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            placeholder_idx = None
            if shape.is_placeholder:
                placeholder_idx = shape.placeholder_format.idx
            results.append({
                "shape_index": i,
                "text": text,
                "placeholder_idx": placeholder_idx,
            })
    return results


def extract_text_from_shape(shape):
    """Extract text from a single shape.

    Args:
        shape: python-pptx Shape object with text_frame.

    Returns:
        list of paragraph text strings.
    """
    if not shape.has_text_frame:
        return []
    texts = []
    for para in shape.text_frame.paragraphs:
        text = para.text.strip()
        if text:
            texts.append(text)
    return texts
