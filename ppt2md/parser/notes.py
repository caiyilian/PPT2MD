"""Speaker notes extraction from PPTX slides."""


def extract_notes_from_slide(slide):
    """Extract speaker notes from a slide.

    Args:
        slide: python-pptx Slide object.

    Returns:
        str: Notes text, or empty string if no notes.
    """
    if not slide.has_notes_slide:
        return ""

    notes_slide = slide.notes_slide
    if not hasattr(notes_slide, "notes_text_frame"):
        return ""

    return notes_slide.notes_text_frame.text.strip()


def format_notes_markdown(notes_text):
    """Format speaker notes as Markdown.

    Args:
        notes_text: The notes text.

    Returns:
        str: Formatted notes section, or empty string if no notes.
    """
    if not notes_text:
        return ""
    return "<!-- Notes -->\n### Speaker Notes\n\n{}".format(notes_text)
