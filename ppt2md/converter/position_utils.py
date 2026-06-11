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


def get_crop_info(shape):
    """Extract crop values from a shape (if it has image fill).

    Args:
        shape: python-pptx Shape object.

    Returns:
        dict with crop_left, crop_right, crop_top, crop_bottom as ratios (0.0-1.0),
        or None if no crop info available.
    """
    if not hasattr(shape, "image") or shape.image is None:
        return None

    try:
        crop_left = getattr(shape, "crop_left", None)
        crop_right = getattr(shape, "crop_right", None)
        crop_top = getattr(shape, "crop_top", None)
        crop_bottom = getattr(shape, "crop_bottom", None)

        if crop_left is None and crop_right is None and crop_top is None and crop_bottom is None:
            return None

        cl = crop_left or 0
        cr = crop_right or 0
        ct = crop_top or 0
        cb = crop_bottom or 0

        if cl == 0 and cr == 0 and ct == 0 and cb == 0:
            return None

        return {
            "crop_left": cl,
            "crop_right": cr,
            "crop_top": ct,
            "crop_bottom": cb,
        }
    except Exception:
        return None


def format_position_comment(shape):
    """Generate a Markdown HTML comment with position/size info and optional crop.

    Args:
        shape: python-pptx Shape object.

    Returns:
        str: HTML comment with position and optional crop info.
    """
    x = emu_to_cm(shape.left)
    y = emu_to_cm(shape.top)
    w = emu_to_cm(shape.width)
    h = emu_to_cm(shape.height)
    parts = ["position: x={:.2f}cm, y={:.2f}cm, width={:.2f}cm, height={:.2f}cm".format(x, y, w, h)]

    crop = get_crop_info(shape)
    if crop:
        parts.append("crop: left={}, right={}, top={}, bottom={}".format(
            crop["crop_left"], crop["crop_right"], crop["crop_top"], crop["crop_bottom"]
        ))

    return "<!-- {} -->".format(", ".join(parts))


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
