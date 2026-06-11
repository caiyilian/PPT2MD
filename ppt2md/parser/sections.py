"""Section detection from PPTX presentations."""


def get_sections(prs):
    """Extract section information from a presentation.

    Args:
        prs: python-pptx Presentation object.

    Returns:
        list of dict with section name and slide indices.
    """
    sections = []
    try:
        for section in prs.sections:
            sections.append({
                "name": section.name,
                "start_slide_idx": section.start_slide_idx if hasattr(section, "start_slide_idx") else 0,
            })
    except (AttributeError, TypeError):
        pass

    return sections


def format_section_markdown(section):
    """Format a section as Markdown heading.

    Args:
        section: dict with section info.

    Returns:
        str: Markdown section heading.
    """
    name = section.get("name", "Untitled Section")
    return "# Section: {}".format(name)
