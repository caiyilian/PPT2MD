"""Tests for Phase 27: Error Handling."""

from ppt2md.converter.error_handler import (
    handle_file_error,
    handle_parse_error,
    handle_image_error,
    handle_formula_error,
    get_exit_code,
    EXIT_SUCCESS,
    EXIT_PARTIAL_FAILURE,
    EXIT_ALL_FAILED,
)


def test_handle_file_error():
    result = handle_file_error(Exception("not found"), "test.pptx")
    assert result == EXIT_ALL_FAILED


def test_handle_parse_error():
    result = handle_parse_error(Exception("bad xml"), "test.pptx")
    assert result == EXIT_PARTIAL_FAILURE


def test_handle_image_error():
    msg = handle_image_error(Exception("corrupt"), 5)
    assert "shape 5" in msg


def test_handle_formula_error():
    result = handle_formula_error(Exception("bad formula"), 3)
    assert "Formula" in result or "formula" in result


def test_get_exit_code_no_errors():
    assert get_exit_code([]) == EXIT_SUCCESS


def test_get_exit_code_few_errors():
    assert get_exit_code(["e1", "e2"]) == EXIT_PARTIAL_FAILURE


def test_get_exit_code_many_errors():
    assert get_exit_code(["e1", "e2", "e3"]) == EXIT_ALL_FAILED
