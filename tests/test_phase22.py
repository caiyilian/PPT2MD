"""Tests for Phase 22: Section Detection."""

from pptx import Presentation

from ppt2md.parser.sections import get_sections, format_section_markdown


def test_get_sections_empty():
    prs = Presentation()
    sections = get_sections(prs)
    assert isinstance(sections, list)


def test_format_section_markdown():
    section = {"name": "Introduction"}
    result = format_section_markdown(section)
    assert result == "# Section: Introduction"


def test_format_section_markdown_no_name():
    section = {}
    result = format_section_markdown(section)
    assert "Untitled Section" in result
