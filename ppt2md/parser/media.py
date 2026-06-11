"""Media and OLE object extraction from PPTX slides."""

import os


MEDIA_EXTENSIONS = {
    ".mp3": ".mp3",
    ".wav": ".wav",
    ".wma": ".wma",
    ".mp4": ".mp4",
    ".avi": ".avi",
    ".wmv": ".wmv",
    ".mov": ".mov",
    ".mid": ".mid",
    ".midi": ".midi",
}


def extract_media_from_slide(slide, output_dir, slide_index):
    """Extract media files (audio/video) from a slide.

    Args:
        slide: python-pptx Slide object.
        output_dir: Directory to save media files.
        slide_index: 1-based slide number for naming.

    Returns:
        list of dict with media info.
    """
    os.makedirs(output_dir, exist_ok=True)
    media_files = []
    media_count = 0

    for shape in slide.shapes:
        # Check for media shapes (audio/video)
        if hasattr(shape, "media") and shape.media is not None:
            media_count += 1
            media = shape.media
            content_type = media.content_type
            ext = MEDIA_EXTENSIONS.get("." + content_type.split("/")[-1], ".bin")
            filename = "slide_{:02d}_media_{:02d}{}".format(slide_index, media_count, ext)
            filepath = os.path.join(output_dir, filename)

            with open(filepath, "wb") as f:
                f.write(media.blob)

            media_files.append({
                "filename": filename,
                "path": filepath,
                "content_type": content_type,
                "shape_index": slide.shapes.index(shape),
            })

    return media_files


def detect_ole_objects(slide):
    """Detect OLE embedded objects in a slide.

    Args:
        slide: python-pptx Slide object.

    Returns:
        list of dict with OLE object info.
    """
    ole_objects = []
    for i, shape in enumerate(slide.shapes):
        # OLE objects typically have shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
        # or can be detected by checking the shape's XML for OLE references
        if hasattr(shape, "shape_type"):
            shape_type_str = str(shape.shape_type)
            if "OLE" in shape_type_str or "EMBEDDED" in shape_type_str:
                ole_objects.append({
                    "shape_index": i,
                    "name": shape.name,
                    "type": "embedded_object",
                })

    return ole_objects


def format_media_markdown(media_info):
    """Format media info as Markdown.

    Args:
        media_info: dict from extract_media_from_slide.

    Returns:
        str: Markdown representation.
    """
    content_type = media_info.get("content_type", "")
    if "audio" in content_type:
        return "[Audio: {}](media/{})".format(content_type, media_info["filename"])
    elif "video" in content_type:
        return "[Video: {}](media/{})".format(content_type, media_info["filename"])
    return "[Media: {}](media/{})".format(content_type, media_info["filename"])


def format_ole_markdown(ole_info):
    """Format OLE object info as Markdown.

    Args:
        ole_info: dict from detect_ole_objects.

    Returns:
        str: Markdown representation.
    """
    return "[Embedded Object: {}]".format(ole_info["name"])
