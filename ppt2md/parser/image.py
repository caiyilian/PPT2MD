"""Image extraction from PPTX slides."""

import os
import hashlib


CONTENT_TYPE_TO_EXT = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/tiff": ".tiff",
    "image/x-emf": ".emf",
    "image/x-wmf": ".wmf",
}

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def get_image_extension(content_type):
    """Map content type to file extension.

    Args:
        content_type: MIME type string.

    Returns:
        str: File extension with dot, or '.bin' for unknown.
    """
    return CONTENT_TYPE_TO_EXT.get(content_type, ".bin")


def extract_images_from_slide(slide, output_dir, slide_index, seen_rids=None):
    """Extract all images from a slide and save to output_dir.

    Args:
        slide: python-pptx Slide object.
        output_dir: Path to images output directory.
        slide_index: 1-based slide number for naming.
        seen_rids: dict mapping rId to filename for deduplication.

    Returns:
        list of dict with image info (filename, path, content_type).
    """
    if seen_rids is None:
        seen_rids = {}

    os.makedirs(output_dir, exist_ok=True)
    images = []
    img_count = [0]

    def save_blob(blob, content_type, path, kind):
        image_key = hashlib.sha1(blob).hexdigest()
        if image_key in seen_rids:
            filename = seen_rids[image_key]
            images.append({
                "filename": filename,
                "path": os.path.join(output_dir, filename),
                "content_type": content_type,
                "shape_index": path[0],
                "shape_path": list(path),
                "kind": kind,
                "deduplicated": True,
            })
            return

        img_count[0] += 1
        ext = get_image_extension(content_type)
        filename = "slide_{:02d}_img_{:02d}{}".format(slide_index, img_count[0], ext)
        filepath = os.path.join(output_dir, filename)

        with open(filepath, "wb") as f:
            f.write(blob)

        seen_rids[image_key] = filename
        images.append({
            "filename": filename,
            "path": filepath,
            "content_type": content_type,
            "shape_index": path[0],
            "shape_path": list(path),
            "kind": kind,
            "deduplicated": False,
        })

    def walk_shapes(shapes, path=()):
        for idx, shape in enumerate(shapes):
            current_path = path + (idx,)
            if hasattr(shape, "shapes"):
                walk_shapes(shape.shapes, current_path)

            if not hasattr(shape, "image") or shape.image is None:
                blips = shape.element.findall(
                    ".//{{{}}}spPr/{{{}}}blipFill/{{{}}}blip".format(P_NS, A_NS, A_NS)
                )
                for blip in blips:
                    rid = blip.get("{{{}}}embed".format(R_NS))
                    if not rid:
                        continue
                    try:
                        part = shape.part.related_part(rid)
                    except KeyError:
                        continue
                    save_blob(part.blob, part.content_type, current_path, "fill")
                continue

            image = shape.image
            save_blob(image.blob, image.content_type, current_path, "picture")

    walk_shapes(slide.shapes)
    return images
