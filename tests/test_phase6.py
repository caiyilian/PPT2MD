"""Tests for Phase 6: List and Bullet Detection."""

from ppt2md.converter.list_utils import (
    is_ordered_list,
    format_list_item,
    format_ordered_item,
    format_paragraph_as_list,
)


def test_is_ordered_list():
    assert is_ordered_list("1. First item") is True
    assert is_ordered_list("2. Second item") is True
    assert is_ordered_list("10. Tenth item") is True


def test_is_not_ordered_list():
    assert is_ordered_list("Just text") is False
    assert is_ordered_list("- bullet") is False
    assert is_ordered_list("") is False


def test_format_list_item_level_0():
    assert format_list_item("item") == "- item"


def test_format_list_item_level_1():
    assert format_list_item("item", level=1) == "  - item"


def test_format_list_item_level_2():
    assert format_list_item("item", level=2) == "    - item"


def test_format_ordered_item():
    assert format_ordered_item("item") == "1. item"


def test_format_ordered_item_indented():
    assert format_ordered_item("item", level=1) == "  1. item"


def test_format_paragraph_as_list_bullet():
    is_list, result = format_paragraph_as_list("bullet item")
    assert is_list is True
    assert result == "- bullet item"


def test_format_paragraph_as_list_ordered():
    is_list, result = format_paragraph_as_list("1. ordered item")
    assert is_list is True
    assert result == "1. ordered item"


def test_format_paragraph_as_list_empty():
    is_list, result = format_paragraph_as_list("")
    assert is_list is False
