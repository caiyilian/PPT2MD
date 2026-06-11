"""Error handling for PPTX to Markdown conversion."""

import sys
import traceback


# Exit codes
EXIT_SUCCESS = 0
EXIT_PARTIAL_FAILURE = 1
EXIT_ALL_FAILED = 2


class ConversionError(Exception):
    """Base exception for conversion errors."""
    pass


class FileError(ConversionError):
    """Error opening or reading a file."""
    pass


class ParseError(ConversionError):
    """Error parsing PPTX content."""
    pass


class ImageError(ConversionError):
    """Error processing an image."""
    pass


class FormulaError(ConversionError):
    """Error converting a formula."""
    pass


def handle_file_error(error, filename, verbose=False):
    """Handle file-related errors.

    Args:
        error: The exception.
        filename: Name of the file being processed.
        verbose: Whether to show detailed error info.

    Returns:
        int: Exit code.
    """
    print("Error processing {}: {}".format(filename, str(error)), file=sys.stderr)
    if verbose:
        traceback.print_exc()
    return EXIT_ALL_FAILED


def handle_parse_error(error, filename, verbose=False):
    """Handle parsing errors.

    Args:
        error: The exception.
        filename: Name of the file being processed.
        verbose: Whether to show detailed error info.

    Returns:
        int: Exit code.
    """
    print("Parse error in {}: {}".format(filename, str(error)), file=sys.stderr)
    if verbose:
        traceback.print_exc()
    return EXIT_PARTIAL_FAILURE


def handle_image_error(error, shape_index, verbose=False):
    """Handle image processing errors (skip and continue).

    Args:
        error: The exception.
        shape_index: Index of the problematic shape.
        verbose: Whether to show detailed error info.

    Returns:
        str: Warning message.
    """
    msg = "Warning: Skipping image at shape {} ({})".format(shape_index, str(error))
    print(msg, file=sys.stderr)
    if verbose:
        traceback.print_exc()
    return msg


def handle_formula_error(error, shape_index, verbose=False):
    """Handle formula conversion errors (fallback to raw XML).

    Args:
        error: The exception.
        shape_index: Index of the problematic shape.
        verbose: Whether to show detailed error info.

    Returns:
        str: Fallback text.
    """
    msg = "Warning: Formula conversion failed at shape {} ({})".format(shape_index, str(error))
    print(msg, file=sys.stderr)
    if verbose:
        traceback.print_exc()
    return "[Formula: conversion failed]"


def get_exit_code(errors):
    """Determine the overall exit code based on errors.

    Args:
        errors: list of error messages.

    Returns:
        int: EXIT_SUCCESS, EXIT_PARTIAL_FAILURE, or EXIT_ALL_FAILED.
    """
    if not errors:
        return EXIT_SUCCESS
    return EXIT_PARTIAL_FAILURE if len(errors) < 3 else EXIT_ALL_FAILED
