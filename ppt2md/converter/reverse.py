"""Convert Markdown with embedded metadata back to PPTX."""

import json
import os
import re
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_CONNECTOR_TYPE
from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def parse_metadata_blocks(md_content, tag="PPTX_META"):
    """Extract JSON metadata blocks from markdown content.

    Args:
        md_content: Markdown content string.
        tag: Tag name to look for (default: PPTX_META).

    Returns list of parsed dicts.
    """
    pattern = r'<!-- {}_START\n(.*?)\n{}_END -->'.format(tag, tag)
    matches = re.findall(pattern, md_content, re.DOTALL)
    results = []
    for match in matches:
        try:
            results.append(json.loads(match))
        except json.JSONDecodeError:
            continue
    return results


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


def _parse_shape_type_str(type_str):
    """Parse shape type string like 'AUTO_SHAPE (1)' or 'LINE (9)'.

    Returns (name, int_value) or None.
    """
    if not type_str:
        return None
    m = re.match(r'(\w+)\s*\((\d+)\)', type_str)
    if m:
        return m.group(1), int(m.group(2))
    return None


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
    slides_meta = parse_metadata_blocks(md_content, "PPTX_META")

    pres_meta_list = parse_metadata_blocks(md_content, "PPTX_PRESENTATION_META")
    pres_meta = pres_meta_list[0] if pres_meta_list else {}

    prs = Presentation()
    prs.slide_width = Emu(int(pres_meta.get("slide_width", 12192000)))
    prs.slide_height = Emu(int(pres_meta.get("slide_height", 6858000)))

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
    parsed = _parse_shape_type_str(meta.get("type", ""))
    type_name = parsed[0] if parsed else ""
    type_val = parsed[1] if parsed else 0

    x = meta.get("position", {}).get("x", 0)
    y = meta.get("position", {}).get("y", 0)
    w = meta.get("size", {}).get("width", 100000)
    h = meta.get("size", {}).get("height", 100000)
    rotation = meta.get("rotation", 0)

    # Handle images (MSO_SHAPE_TYPE.PICTURE = 13)
    if type_val == 13 or ("image" in meta and meta["image"]):
        _add_image_shape(slide, meta, images_dir, x, y, w, h)
        return

    # Handle tables
    if "table" in meta and meta["table"]:
        _add_table_shape(slide, meta, x, y, w, h)
        return

    # Handle charts (skip - complex to recreate)
    if "chart" in meta and meta["chart"]:
        return

    # Handle lines/arrows (MSO_SHAPE_TYPE.LINE = 9)
    if type_val == 9 or type_name == "LINE":
        _add_line_shape(slide, meta, x, y, w, h, rotation)
        return

    # Handle groups (MSO_SHAPE_TYPE.GROUP = 6)
    if type_val == 6 or type_name == "GROUP":
        return  # Groups are complex, skip for now

    # Handle text boxes (MSO_SHAPE_TYPE.TEXT_BOX = 17)
    if type_val == 17 or type_name == "TEXT_BOX":
        _add_text_box_shape(slide, meta, x, y, w, h, rotation)
        return

    # Handle auto shapes and everything else
    _add_auto_shape(slide, meta, x, y, w, h, rotation)


def _add_text_box_shape(slide, meta, x, y, w, h, rotation):
    """Add a text box shape (no fill, no line by default)."""
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, h)
    shape.rotation = rotation

    # Text boxes have no fill and no line by default
    shape.fill.background()

    line_meta = meta.get("line")
    if line_meta:
        width = line_meta.get("width", 0)
        color = line_meta.get("color")
        if width > 0 and color:
            _apply_line(shape, line_meta)
        elif color is None or width == 0:
            # No visible line
            try:
                shape.line.fill.background()
            except Exception:
                pass
    else:
        try:
            shape.line.fill.background()
        except Exception:
            pass

    text_meta = meta.get("text")
    if text_meta:
        _apply_text(shape, text_meta)


def _add_auto_shape(slide, meta, x, y, w, h, rotation):
    """Add an auto shape (rectangle, rounded rect, etc.) with proper fill/line."""
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

    auto_type_str = meta.get("auto_shape_type", "")
    auto_type_val = None
    if auto_type_str:
        parsed = re.match(r'.*?\((\d+)\)', auto_type_str)
        if parsed:
            auto_type_val = int(parsed.group(1))

    auto_shape_map = {
        1: MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        5: MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        9: MSO_AUTO_SHAPE_TYPE.OVAL,
        33: MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW,
        34: MSO_AUTO_SHAPE_TYPE.LEFT_ARROW,
    }

    auto_type = auto_shape_map.get(auto_type_val, MSO_AUTO_SHAPE_TYPE.RECTANGLE)
    try:
        shape = slide.shapes.add_shape(auto_type, x, y, w, h)
    except Exception:
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, h)

    shape.rotation = rotation

    # Apply fill
    fill_meta = meta.get("fill")
    if fill_meta:
        _apply_fill(shape, fill_meta)
    else:
        # No fill metadata → transparent
        try:
            shape.fill.background()
        except Exception:
            pass

    # Apply line
    line_meta = meta.get("line")
    if line_meta:
        width = line_meta.get("width", 0)
        color = line_meta.get("color")
        if width > 0 and color:
            _apply_line(shape, line_meta)
        elif color is None or width == 0:
            try:
                shape.line.fill.background()
            except Exception:
                pass
    else:
        try:
            shape.line.fill.background()
        except Exception:
            pass

    # Apply text
    text_meta = meta.get("text")
    if text_meta:
        _apply_text(shape, text_meta)


def _add_image_shape(slide, meta, images_dir, x, y, w, h):
    """Add an image shape."""
    image_meta = meta.get("image", {})
    content_type = image_meta.get("content_type", "")

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
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, h)
    shape.fill.background()
    shape.text_frame.text = "[Image: {}]".format(content_type)


def _add_line_shape(slide, meta, x, y, w, h, rotation):
    """Add a line/connector shape with arrowheads."""
    from pptx.enum.shapes import MSO_CONNECTOR_TYPE

    try:
        shape = slide.shapes.add_connector(
            MSO_CONNECTOR_TYPE.STRAIGHT, x, y, x + w, y + h
        )
    except Exception:
        from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, max(w, 1), max(h, 1))
        shape.fill.background()

    # Apply line properties including arrowheads
    line_meta = meta.get("line")
    if line_meta:
        _apply_line(shape, line_meta)

        # Apply arrowheads via XML
        _apply_arrowheads(shape, line_meta)


def _apply_arrowheads(shape, line_meta):
    """Apply arrowhead (tailEnd/headEnd) to a connector shape's line XML."""
    try:
        spPr = _get_spPr(shape)
        if spPr is None:
            return
        ln = _make_ln(spPr, line_meta.get("width", 0))

        for end_key in ["tailEnd", "headEnd"]:
            arrow_info = line_meta.get(end_key)
            if not arrow_info:
                continue
            arr_type = arrow_info.get("type")
            if not arr_type:
                continue
            tag = "{{{}}}{}".format(A_NS, end_key)
            existing = ln.find(tag)
            if existing is not None:
                ln.remove(existing)
            elem = ln.makeelement(tag, {})
            elem.set("type", arr_type)
            if arrow_info.get("w"):
                elem.set("w", arrow_info["w"])
            if arrow_info.get("len"):
                elem.set("len", arrow_info["len"])
            ln.append(elem)
    except Exception:
        pass


def _add_table_shape(slide, meta, x, y, w, h):
    """Add a table shape."""
    table_meta = meta.get("table", {})
    rows = table_meta.get("rows", 1)
    cols = table_meta.get("cols", 1)

    slide.shapes.add_table(rows, cols, x, y, w, h)


def _apply_fill(shape, fill_meta):
    """Apply fill to a shape.

    Handles solid (RGB), scheme (theme), and no-fill types.
    """
    fill_type = fill_meta.get("type")
    color = fill_meta.get("color")

    if fill_type == "none":
        try:
            shape.fill.background()
        except Exception:
            pass
        return

    if fill_type == "solid":
        if color:
            try:
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor.from_string(color)
            except Exception:
                pass
        return

    if fill_type == "scheme":
        _apply_scheme_fill(shape, color)
        return


def _apply_scheme_fill(shape, theme_color):
    """Apply a theme/scheme color fill via XML.

    Example: <a:solidFill><a:schemeClr val="accent6"/></a:solidFill>
    """
    if not theme_color:
        return
    try:
        spPr = _get_spPr(shape)
        if spPr is None:
            return
        # Remove existing fill elements
        fill_tags = ['{}Fill'.format(A_NS), '{}fill'.format(A_NS),
                     '{}NoFill'.format(A_NS), '{}noFill'.format(A_NS),
                     '{}solidFill'.format(A_NS), '{}blipFill'.format(A_NS),
                     '{}gradFill'.format(A_NS), '{}pattFill'.format(A_NS)]
        for child in list(spPr):
            if child.tag in fill_tags:
                spPr.remove(child)
        # Add scheme fill as first child of spPr (after xfrm, before geom)
        solid = etree.Element('{{{}}}solidFill'.format(A_NS))
        scheme = etree.SubElement(solid, '{{{}}}schemeClr'.format(A_NS))
        scheme.set('val', theme_color)
        spPr.insert(1, solid)  # After xfrm (index 0)
    except Exception:
        pass


def _is_hex_color(color_str):
    """Check if a color string is a valid hex color (6 hex digits)."""
    if not color_str or not isinstance(color_str, str):
        return False
    return bool(re.match(r'^[0-9A-Fa-f]{6}$', color_str))


def _apply_line(shape, line_meta):
    """Apply line properties to a shape.

    Handles connectors and regular shapes. Creates a:ln XML element
    if needed for arrowhead support.
    """
    width = line_meta.get("width", 0)
    color = line_meta.get("color")

    if width > 0:
        try:
            shape.line.width = width
        except Exception:
            pass

    if _is_hex_color(color):
        try:
            shape.line.color.rgb = RGBColor.from_string(color)
        except Exception:
            pass
    elif color is not None and not _is_hex_color(color):
        # Theme color (e.g. "tx1", "accent1") - set via XML
        _apply_theme_line_color(shape, color, width)

    if color is None and width == 0:
        try:
            shape.line.fill.background()
        except Exception:
            pass

    # Ensure a:ln exists for arrowhead support if not already present
    if width <= 0 and not _is_hex_color(color):
        _ensure_line_xml(shape)


def _apply_theme_line_color(shape, theme_color, width):
    """Apply a theme/scheme color to line via XML."""
    try:
        spPr = _get_spPr(shape)
        if spPr is None:
            return
        ln = spPr.find('{{{}}}ln'.format(A_NS))
        if ln is None:
            ln = _make_ln(spPr, width)
        # Remove existing fill
        for child in list(ln):
            tag = child.tag
            if 'Fill' in tag or 'fill' in tag or 'NoFill' in tag:
                ln.remove(child)
        solid = ln.makeelement('{{{}}}solidFill'.format(A_NS), {})
        scheme = solid.makeelement('{{{}}}schemeClr'.format(A_NS), {'val': theme_color})
        solid.append(scheme)
        ln.append(solid)
    except Exception:
        pass


def _get_spPr(shape):
    """Get the spPr element from a shape (handles both sp and cxnSp)."""
    return shape._element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}spPr')


def _make_ln(spPr, width=0):
    """Create an a:ln element under spPr if it doesn't exist."""
    ln = spPr.find('{{{}}}ln'.format(A_NS))
    if ln is not None:
        return ln
    ln = etree.SubElement(spPr, '{{{}}}ln'.format(A_NS))
    if width > 0:
        ln.set('w', str(width))
    return ln


def _ensure_line_xml(shape):
    """Ensure a minimal a:ln element exists for arrowhead support."""
    try:
        spPr = _get_spPr(shape)
        if spPr is None:
            return
        ln = spPr.find('{{{}}}ln'.format(A_NS))
        if ln is None:
            etree.SubElement(spPr, '{{{}}}ln'.format(A_NS))
    except Exception:
        pass


def _apply_text(shape, text_meta):
    """Apply text content and formatting to a shape."""
    paragraphs = text_meta.get("paragraphs", [])
    if not paragraphs:
        return

    tf = shape.text_frame
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
