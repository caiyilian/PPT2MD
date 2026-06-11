"""Background image extraction from PPTX slides."""

import os


def extract_background_image(slide, output_dir, slide_index):
    """Extract background image from a slide if present.

    Args:
        slide: python-pptx Slide object.
        output_dir: Directory to save background images.
        slide_index: 1-based slide number for naming.

    Returns:
        dict with background info, or None if no background image.
    """
    background = slide.background
    if background is None:
        return None

    fill = background.fill
    if fill is None:
        return None

    # Check for picture fill
    try:
        if fill.type is not None and hasattr(fill, "picture"):
            picture = fill.picture
            if picture is not None and hasattr(picture, "image"):
                image = picture.image
                content_type = image.content_type
                ext = ".png"
                if "jpeg" in content_type:
                    ext = ".jpg"
                elif "gif" in content_type:
                    ext = ".gif"

                filename = "slide_{:02d}_bg{}".format(slide_index, ext)
                os.makedirs(output_dir, exist_ok=True)
                filepath = os.path.join(output_dir, filename)

                with open(filepath, "wb") as f:
                    f.write(image.blob)

                return {
                    "filename": filename,
                    "path": filepath,
                    "content_type": content_type,
                }
    except (AttributeError, TypeError):
        pass

    return None


def format_background_markdown(bg_info):
    """Format background image info as Markdown.

    Args:
        bg_info: dict from extract_background_image.

    Returns:
        str: Markdown image reference for background.
    """
    if bg_info is None:
        return ""
    return "![background](images/{})".format(bg_info["filename"])
