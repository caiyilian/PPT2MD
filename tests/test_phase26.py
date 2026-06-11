"""Tests for Phase 26: Batch Conversion."""

from pptx import Presentation

from ppt2md.converter.batch import find_pptx_files, get_output_dir_for_file, batch_convert_summary


def test_find_pptx_files_single(tmp_path):
    pptx = tmp_path / "test.pptx"
    Presentation().save(str(pptx))
    files = find_pptx_files(pptx)
    assert len(files) == 1
    assert files[0] == pptx


def test_find_pptx_files_directory(tmp_path):
    for i in range(3):
        pptx = tmp_path / "test{}.pptx".format(i)
        Presentation().save(str(pptx))
    (tmp_path / "other.txt").touch()
    files = find_pptx_files(tmp_path)
    assert len(files) == 3


def test_find_pptx_files_empty(tmp_path):
    files = find_pptx_files(tmp_path)
    assert len(files) == 0


def test_get_output_dir():
    from pathlib import Path
    pptx = Path("/data/presentation.pptx")
    result = get_output_dir_for_file(pptx)
    assert result.name == "presentation_output"


def test_batch_convert_summary_success():
    results = [
        {"file": "a.pptx", "success": True},
        {"file": "b.pptx", "success": True},
    ]
    summary = batch_convert_summary(results)
    assert "Success: 2" in summary
    assert "Failed: 0" in summary


def test_batch_convert_summary_mixed():
    results = [
        {"file": "a.pptx", "success": True},
        {"file": "b.pptx", "success": False, "error": "Parse error"},
    ]
    summary = batch_convert_summary(results)
    assert "Failed: 1" in summary
    assert "Parse error" in summary
