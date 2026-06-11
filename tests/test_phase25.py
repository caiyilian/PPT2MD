"""Tests for Phase 25: Full CLI Integration."""

from ppt2md.main import main


def test_cli_help():
    try:
        main(["--help"])
    except SystemExit:
        pass  # --help exits with 0


def test_cli_basic():
    import pytest
    with pytest.raises(SystemExit) as exc_info:
        main(["--help"])
    assert exc_info.value.code == 0


def test_cli_nonexistent():
    result = main(["nonexistent.pptx"])
    assert result == 1
