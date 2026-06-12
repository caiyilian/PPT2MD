"""Convert Markdown with embedded metadata back to PPTX."""

import json
import os
import re
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE


def parse_metadata_blocks(md_content):
    """Extract PPTX_META JSON blocks from markdown content.

    Returns list of slide metadata dicts.
    """
    pattern = r'<!-- PPTX_META_START\n(.*?)\nPPTX_META_END -->'
    matches = re.findall(pattern, md_content, re.DOTALL)
    slides = []
    for match in matches:
        try:
            slide_meta = json.loads(match)
            slides.append(slide_meta)
        except json.JSONDecodeError:
            continue
    return slides


def parse_frontmatter(md_content):
    """Extract YAML frontmatter from markdown."""
    match = re.match(r'^---\n(.*?)\n---', md_content, re.DOTALL)
    if not match:
        return {}
    fm = {}
    for line in match.group(1).split('\n'):
        if ':' in line:
            key, _, value = line.partition(':')
            value = value.strip().strip('"')
            fm[key.strip()] = value
    return fm


def md_to_pptx(md_path, images_dir=None, output_path=None):
    """Convert a markdown file with embedded metadata back to PPTX.

    Args:
        md_path: Path to the .md file.
        images_dir: Path to images directory (default: same dir as md).
        output_path: Output .pptx path (default: same name as md).

    Returns:
        Path to the created PPTX file.
    """
    md_path = Path(md_path)
    if images_dir is None:
        images_dir = md_path.parent / "images"
    else:
        images_dir = Path(images_dir)
    if output_path is None:
        output_path = md_path.with_suffix(".pptx")
    else:
        output_path = Path(output_path)

    with open(md_path, "r", encoding="utf-8") as f:
        md_content = f.read()

    frontmatter = parse_frontmatter(md_content)
    slides_meta = parse_metadata_blocks(md_content)

    prs = Presentation()
    prs.slide_width = Emu(12192000)  # Standard 33.867cm
    prs.slide_height = Emu(6858000)  # Standard 19.05cm

    # Set document properties
    if frontmatter.get("title"):
        prs.core_properties.title = frontmatter["title"]
    if frontmatter.get("author"):
        prs.core_properties.author = frontmatter["author"]

    slide_layout = prs.slide_layouts[6]  # Blank layout

    for slide_meta in slides_meta:
        slide = prs.slides.add_slide(slide_layout)

        shapes_meta = slide_meta.get("shapes", [])
        for shape_meta in shapes_meta:
            _add_shape_from_metadata(slide, shape_meta, images_dir)

    prs.save(str(output_path))
    return output_path


def _add_shape_from_metadata(slide, meta, images_dir):
    """Add a shape to a slide based on metadata."""
    shape_type = meta.get("type", "")

    x = meta.get("position", {}).get("x", 0)
    y = meta.get("position", {}).get("y", 0)
    w = meta.get("size", {}).get("width", 100000)
    h = meta.get("size", {}).get("height", 100000)
    rotation = meta.get("rotation", 0)

    # Handle images
    if "image" in meta and meta["image"]:
        _add_image_shape(slide, meta, images_dir, x, y, w, h)
        return

    # Handle tables
    if "table" in meta and meta["table"]:
        _add_table_shape(slide, meta, x, y, w, h)
        return

    # Handle charts (skip - complex to recreate)
    if "chart" in meta and meta["chart"]:
        return

    # Handle lines/arrows
    if shape_type.startswith("LINE"):
        _add_line_shape(slide, meta, x, y, w, h)
        return

    # Handle groups
    if shape_type.startswith("GROUP"):
        return  # Groups are complex, skip for now

    # Handle auto shapes and text boxes
    _add_auto_shape(slide, meta, x, y, w, h, rotation)


def _add_auto_shape(slide, meta, x, y, w, h, rotation):
    """Add an auto shape (rectangle, rounded rect, etc.) with text."""
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

    auto_type_str = meta.get("auto_shape_type", "")
    shape_type_map = {
        "ROUNDED_RECTANGLE": MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        "RECTANGLE": MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        "OVAL": MSO_AUTO_SHAPE_TYPE.OVAL,
        "RIGHT_ARROW": MSO_SHAPE_TYPE.AUTO_SHAPE,
    }

    try:
        auto_type = shape_type_map.get(auto_type_str, MSO_AUTO_SHAPE_TYPE.RECTANGLE)
        shape = slide.shapes.add_shape(auto_type, x, y, w, h)
    except Exception:
        shape = slide.shapes.add_shape(1, x, y, w, h)  # Fallback to rectangle

    shape.rotation = rotation

    # Apply fill
    fill_meta = meta.get("fill")
    if fill_meta:
        _apply_fill(shape, fill_meta)

    # Apply line
    line_meta = meta.get("line")
    if line_meta:
        _apply_line(shape, line_meta)

    # Apply text
    text_meta = meta.get("text")
    if text_meta:
        _apply_text(shape, text_meta)


def _add_image_shape(slide, meta, images_dir, x, y, w, h):
    """Add an image shape."""
    image_meta = meta.get("image", {})
    content_type = image_meta.get("content_type", "")

    # Try to find the image file
    slide_num = meta.get("_slide_num", 1)
    shape_name = meta.get("name", "")

    ext = ".png"
    if "jpeg" in content_type:
        ext = ".jpg"
    elif "gif" in content_type:
        ext = ".gif"

    # Search for matching image in images directory
    if images_dir.exists():
        for img_file in images_dir.iterdir():
            if img_file.suffix.lower() in [ext, ".png", ".jpg", ".jpeg"]:
                try:
                    slide.shapes.add_picture(str(img_file), x, y, w, h)
                    return
                except Exception:
                    continue

    # If no image found, add a placeholder rectangle
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.text_frame.text = "[Image: {}]".format(content_type)


def _add_line_shape(slide, meta, x, y, w, h):
    """Add a line/arrow shape."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(1, x, y, w, h)

    line_meta = meta.get("line", {})
    _apply_line(shape, line_meta)


def _add_table_shape(slide, meta, x, y, w, h):
    """Add a table shape."""
    table_meta = meta.get("table", {})
    rows = table_meta.get("rows", 1)
    cols = table_meta.get("cols", 1)

    table_shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    # Table content would need to be in metadata too


def _apply_fill(shape, fill_meta):
    """Apply fill to a shape."""
    if fill_meta.get("type") == "none":
        shape.fill.background()
    elif fill_meta.get("type") == "solid":
        color = fill_meta.get("color")
        if color:
            try:
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor.from_string(color)
            except Exception:
                pass


def _apply_line(shape, line_meta):
    """Apply line properties to a shape."""
    width = line_meta.get("width", 0)
    if width > 0:
        shape.line.width = width

    color = line_meta.get("color")
    if color:
        try:
            shape.line.color.rgb = RGBColor.from_string(color)
        except Exception:
            pass


def _apply_text(shape, text_meta):
    """Apply text content and formatting to a shape."""
    paragraphs = text_meta.get("paragraphs", [])
    if not paragraphs:
        return

    tf = shape.text_frame
    # Clear default paragraph
    if tf.paragraphs:
        tf.paragraphs[0].text = ""

    for i, para_meta in enumerate(paragraphs):
        if i == 0:
            para = tf.paragraphs[0]
        else:
            para = tf.add_paragraph()

        level = para_meta.get("level", 0)
        para.level = level

        for run_meta in para_meta.get("runs", []):
            text = run_meta.get("text", "")
            run = para.add_run()
            run.text = text

            if run_meta.get("font_size"):
                run.font.size = run_meta["font_size"]
            if run_meta.get("bold"):
                run.font.bold = True
            if run_meta.get("italic"):
                run.font.italic = True
            if run_meta.get("underline"):
                run.font.underline = True
            if run_meta.get("font_name"):
                run.font.name = run_meta["font_name"]
            if run_meta.get("font_color"):
                try:
                    run.font.color.rgb = RGBColor.from_string(run_meta["font_color"])
                except Exception:
                    pass
