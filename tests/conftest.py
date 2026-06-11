"""Pytest fixtures for ppt2md tests."""

import os
import tempfile
from pathlib import Path

import pytest


@pytest.fixture
def scratch_dir():
    """Path to the scratch directory for test outputs."""
    scratch = Path(__file__).parent.parent / "scratch"
    scratch.mkdir(exist_ok=True)
    return scratch


@pytest.fixture
def output_dir(scratch_dir):
    """Create a temporary output directory inside scratch/."""
    out = scratch_dir / "output"
    out.mkdir(exist_ok=True)
    return out


@pytest.fixture
def test_files_dir(scratch_dir):
    """Path to the test files directory."""
    tf = scratch_dir / "test_files"
    tf.mkdir(exist_ok=True)
    return tf
