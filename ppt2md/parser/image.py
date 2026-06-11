"""Image extraction from PPTX slides."""

import os


CONTENT_TYPE_TO_EXT = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/tiff": ".tiff",
    "image/x-emf": ".emf",
    "image/x-wmf": ".wmf",
}


def get_image_extension(content_type):
    """Map content type to file extension.

    Args:
        content_type: MIME type string.

    Returns:
        str: File extension with dot, or '.bin' for unknown.
    """
    return CONTENT_TYPE_TO_EXT.get(content_type, ".bin")


def extract_images_from_slide(slide, output_dir, slide_index):
    """Extract all images from a slide and save to output_dir.

    Args:
        slide: python-pptx Slide object.
        output_dir: Path to images output directory.
        slide_index: 1-based slide number for naming.

    Returns:
        list of dict with image info (filename, path, content_type).
    """
    os.makedirs(output_dir, exist_ok=True)
    images = []
    img_count = 0

    for shape in slide.shapes:
        if not hasattr(shape, "image") or shape.image is None:
            continue

        img_count += 1
        image = shape.image
        content_type = image.content_type
        ext = get_image_extension(content_type)
        filename = "slide_{:02d}_img_{:02d}{}".format(slide_index, img_count, ext)
        filepath = os.path.join(output_dir, filename)

        with open(filepath, "wb") as f:
            f.write(image.blob)

        images.append({
            "filename": filename,
            "path": filepath,
            "content_type": content_type,
            "shape_index": slide.shapes.index(shape),
        })

    return images
