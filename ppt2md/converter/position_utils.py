"""Position and size utilities for Markdown output."""

from ppt2md.parser.image import get_image_extension, extract_images_from_slide, CONTENT_TYPE_TO_EXT


EMU_PER_CM = 360000
EMU_PER_INCH = 914400


def emu_to_cm(emu):
    """Convert EMU (English Metric Units) to centimeters.

    Args:
        emu: Value in EMU.

    Returns:
        float: Value in centimeters.
    """
    if emu is None:
        return 0.0
    return emu / EMU_PER_CM


def emu_to_px(emu, dpi=96):
    """Convert EMU to pixels.

    Args:
        emu: Value in EMU.
        dpi: Dots per inch for conversion.

    Returns:
        float: Value in pixels.
    """
    if emu is None:
        return 0.0
    return emu / EMU_PER_INCH * dpi


def format_position_comment(shape):
    """Generate a Markdown HTML comment with position/size info.

    Args:
        shape: python-pptx Shape object.

    Returns:
        str: HTML comment like '<!-- position: x=10.28cm, y=6.57cm, width=28.81cm, height=10.28cm -->'
    """
    x = emu_to_cm(shape.left)
    y = emu_to_cm(shape.top)
    w = emu_to_cm(shape.width)
    h = emu_to_cm(shape.height)
    return "<!-- position: x={:.2f}cm, y={:.2f}cm, width={:.2f}cm, height={:.2f}cm -->".format(
        x, y, w, h
    )


def format_image_markdown(image_info, shape):
    """Generate Markdown for an image with position comment.

    Args:
        image_info: dict from extract_images_from_slide.
        shape: python-pptx Shape object.

    Returns:
        str: Markdown image reference with position comment.
    """
    position = format_position_comment(shape)
    return "{}\n![image](images/{})".format(position, image_info["filename"])
