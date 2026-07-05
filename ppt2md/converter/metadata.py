"""Extract comprehensive shape metadata for round-trip conversion."""

from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _safe_color(color_obj):
    """Extract color value from python-pptx color object."""
    if color_obj is None:
        return None
    try:
        return {"type": "rgb", "value": str(color_obj.rgb)}
    except AttributeError:
        pass
    try:
        return {"type": "theme", "value": str(color_obj.theme_color), "shade": getattr(color_obj, '_color', None) and getattr(color_obj._color, 'shade', None)}
    except AttributeError:
        pass
    return None


def _get_fill_info(shape):
    """Extract fill information from shape XML."""
    spPr = shape.element.find(".//{{{}}}spPr".format(P_NS))
    if spPr is None:
        return None

    solid = spPr.find("{{{}}}solidFill".format(A_NS))
    if solid is not None:
        srgb = solid.find("{{{}}}srgbClr".format(A_NS))
        if srgb is not None:
            return {"type": "solid", "color": srgb.get("val")}
        scheme = solid.find("{{{}}}schemeClr".format(A_NS))
        if scheme is not None:
            info = {"type": "scheme", "color": scheme.get("val")}
            modifiers = _get_color_modifiers(scheme)
            if modifiers:
                info["modifiers"] = modifiers
            return info

    no_fill = spPr.find("{{{}}}noFill".format(A_NS))
    if no_fill is not None:
        return {"type": "none"}

    return None


def _get_color_modifiers(scheme_elem):
    """Extract color modifiers (lumMod, lumOff, satMod, shade, tint) from a schemeClr element."""
    modifiers = {}
    if scheme_elem is None:
        return modifiers
    for child in scheme_elem:
        tag = child.tag
        val = child.get("val")
        if val is not None:
            local = tag.split("}")[-1] if "}" in tag else tag
            modifiers[local] = val
    return modifiers


def _get_line_info(shape):
    """Extract line/border information from shape XML."""
    ln = shape.element.find(".//{{{}}}spPr/{{{}}}ln".format(P_NS, A_NS))
    if ln is None:
        ln = shape.element.find(".//{{{}}}ln".format(A_NS))
    if ln is None:
        return None

    width = ln.get("w")
    info = {"width": int(width) if width else 0}

    solid = ln.find(".//{{{}}}solidFill".format(A_NS))
    if solid is not None:
        srgb = solid.find("{{{}}}srgbClr".format(A_NS))
        if srgb is not None:
            info["color"] = srgb.get("val")
        scheme = solid.find("{{{}}}schemeClr".format(A_NS))
        if scheme is not None:
            info["color"] = scheme.get("val")

    no_fill = ln.find("{{{}}}noFill".format(A_NS))
    if no_fill is not None:
        info["color"] = None

    # Line style (dash)
    prstDash = ln.find(".//{{{}}}prstDash".format(A_NS))
    if prstDash is not None:
        info["dash"] = prstDash.get("val")

    # Arrow heads
    for end in ["tailEnd", "headEnd"]:
        tag = "{{{}}}{}".format(A_NS, end)
        arrow = ln.find(tag)
        if arrow is not None:
            info[end] = {
                "type": arrow.get("type"),
                "w": arrow.get("w"),
                "len": arrow.get("len"),
            }

    # Connector geometry preset (for LINE type shapes)
    spPr = shape.element.find(".//{{{}}}spPr".format(P_NS))
    if spPr is not None:
        prstGeom = spPr.find("{{{}}}prstGeom".format(A_NS))
        if prstGeom is not None:
            prst = prstGeom.get("prst")
            if prst and prst != "line":
                info["connector_geom"] = prst
        # Flip flags on xfrm
        xfrm = spPr.find("{{{}}}xfrm".format(A_NS))
        if xfrm is not None:
            for attr in ("flipH", "flipV"):
                val = xfrm.get(attr)
                if val:
                    info[attr] = val

    return info


def _get_text_info(shape):
    """Extract detailed text information with formatting."""
    if not shape.has_text_frame:
        return None

    paragraphs = []
    for para in shape.text_frame.paragraphs:
        # Only extract alignment if explicitly set on the paragraph
        algn = None
        try:
            A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            pPr = para._p.find('{{{}}}pPr'.format(A_NS))
            if pPr is not None and pPr.get('algn') is not None:
                algn = str(para.alignment)
        except Exception:
            pass
        para_info = {
            "level": para.level or 0,
            "alignment": algn,
            "runs": [],
        }

        for run in para.runs:
            run_info = {"text": run.text}
            font = run.font
            try:
                if font.size:
                    run_info["font_size"] = font.size
            except: pass
            try:
                if font.bold:
                    run_info["bold"] = True
            except: pass
            try:
                if font.italic:
                    run_info["italic"] = True
            except: pass
            try:
                if font.underline:
                    run_info["underline"] = True
            except: pass
            try:
                if getattr(font, 'strikethrough', None):
                    run_info["strikethrough"] = True
            except: pass

            # Superscript/subscript via baseline
            try:
                rPr = run._r.find("{{{}}}rPr".format(A_NS))
                if rPr is not None:
                    baseline = rPr.get("baseline")
                    if baseline:
                        val = int(baseline)
                        if val > 0:
                            run_info["superscript"] = True
                        elif val < 0:
                            run_info["subscript"] = True
            except: pass

            # Font name
            try:
                if font.name:
                    run_info["font_name"] = font.name
            except: pass

            # Font color
            try:
                fc = font.color
                if fc and fc.rgb:
                    run_info["font_color"] = str(fc.rgb)
                elif fc and fc.theme_color:
                    run_info["font_color"] = "theme:{}".format(str(fc.theme_color))
            except:
                # Check XML for scheme color as fallback
                try:
                    rPr = run._r.find("{{{}}}rPr".format(A_NS))
                    if rPr is not None:
                        sf = rPr.find("{{{}}}solidFill".format(A_NS))
                        if sf is not None:
                            sc = sf.find("{{{}}}schemeClr".format(A_NS))
                            if sc is not None:
                                run_info["font_color"] = "theme:{}".format(sc.get("val"))
                except:
                    pass

            para_info["runs"].append(run_info)

        paragraphs.append(para_info)

    return {"paragraphs": paragraphs}


def _get_body_props(shape):
    """Extract bodyPr properties (wrap, autofit, insets) from shape XML."""
    try:
        # Try both A_NS and P_NS for txBody
        txBody = shape._element.find('.//{%s}txBody' % A_NS)
        if txBody is None:
            txBody = shape._element.find('.//{%s}txBody' % P_NS)
        if txBody is None:
            return None
        # bodyPr is always in A_NS
        bodyPr = txBody.find('{%s}bodyPr' % A_NS)
        if bodyPr is None:
            return None

        props = {}
        # wrap attribute
        wrap = bodyPr.get('wrap')
        if wrap is not None:
            props['wrap'] = wrap

        # inset attributes (margins)
        for attr in ('lIns', 'rIns', 'tIns', 'bIns'):
            val = bodyPr.get(attr)
            if val is not None:
                props[attr] = int(val)

        # autofit elements
        spAutoFit = bodyPr.find('{%s}spAutoFit' % A_NS)
        normAutofit = bodyPr.find('{%s}normAutofit' % A_NS)
        noAutofit = bodyPr.find('{%s}noAutofit' % A_NS)

        if spAutoFit is not None:
            props['autofit'] = 'spAutoFit'
        elif normAutofit is not None:
            props['autofit'] = 'normAutofit'
            # normAutofit may have fontScale and lnSpcReduction
            fontScale = normAutofit.get('fontScale')
            lnSpcReduction = normAutofit.get('lnSpcReduction')
            if fontScale:
                props['fontScale'] = int(fontScale)
            if lnSpcReduction:
                props['lnSpcReduction'] = int(lnSpcReduction)
        elif noAutofit is not None:
            props['autofit'] = 'noAutofit'

        return props if props else None
    except Exception:
        return None


def _get_group_info(shape):
    """Extract group shape info including children and coordinate space."""
    try:
        # Get group coordinate space (chOff/chExt)
        grpSpPr = shape._element.find('{%s}grpSpPr' % P_NS)
        coord_space = {}
        if grpSpPr is not None:
            xfrm = grpSpPr.find('{%s}xfrm' % A_NS)
            if xfrm is not None:
                chOff = xfrm.find('{%s}chOff' % A_NS)
                chExt = xfrm.find('{%s}chExt' % A_NS)
                if chOff is not None:
                    coord_space['chOffX'] = int(chOff.get('x'))
                    coord_space['chOffY'] = int(chOff.get('y'))
                if chExt is not None:
                    coord_space['chExtCX'] = int(chExt.get('cx'))
                    coord_space['chExtCY'] = int(chExt.get('cy'))

        # Extract children metadata
        children = []
        if hasattr(shape, 'shapes'):
            for child in shape.shapes:
                child_meta = extract_shape_metadata(child)
                children.append(child_meta)

        result = {"children": children}
        if coord_space:
            result["coord_space"] = coord_space
        return result
    except Exception:
        return None


def extract_shape_metadata(shape):
    """Extract comprehensive metadata for a shape.

    Returns dict with all properties needed for round-trip conversion.
    """
    meta = {
        "name": shape.name,
        "type": str(shape.shape_type),
        "position": {
            "x": shape.left,
            "y": shape.top,
        },
        "size": {
            "width": shape.width,
            "height": shape.height,
        },
        "rotation": shape.rotation or 0,
    }

    # Handle GROUP shapes (MSO_SHAPE_TYPE.GROUP = 6)
    if shape.shape_type == 6:
        meta["group"] = _get_group_info(shape)
        return meta

    # Auto shape type
    try:
        if shape.shape_type == 1:  # MSO_SHAPE_TYPE.AUTO_SHAPE
            meta["auto_shape_type"] = str(shape.auto_shape_type)
    except (AttributeError, ValueError):
        pass

    # Placeholder info
    if shape.is_placeholder:
        meta["placeholder"] = {
            "idx": shape.placeholder_format.idx,
            "type": str(shape.placeholder_format.type),
        }

    # Fill
    fill_info = _get_fill_info(shape)
    if fill_info:
        meta["fill"] = fill_info

    # Line
    line_info = _get_line_info(shape)
    if line_info:
        meta["line"] = line_info

    # Body properties (wrap, autofit)
    body_props = _get_body_props(shape)
    if body_props:
        meta["body_props"] = body_props

    # Text
    text_info = _get_text_info(shape)
    if text_info:
        meta["text"] = text_info

    # Image
    if hasattr(shape, "image") and shape.image is not None:
        meta["image"] = {
            "content_type": shape.image.content_type,
            "blob_size": len(shape.image.blob),
        }

    # Table
    if hasattr(shape, "has_table") and shape.has_table:
        table = shape.table
        meta["table"] = {
            "rows": len(table.rows),
            "cols": len(table.columns),
        }

    # Chart
    if hasattr(shape, "has_chart") and shape.has_chart:
        meta["chart"] = {
            "type": str(shape.chart.chart_type),
        }

    return meta


def _get_fallback_xml(sp):
    """Get the Fallback sp XML from the AlternateContent parent."""
    from lxml import etree
    MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
    P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    try:
        parent = sp.getparent()
        while parent is not None:
            local = parent.tag.split('}')[1] if '}' in parent.tag else ''
            if local == 'AlternateContent':
                fallback = parent.find('{%s}Fallback' % MC_NS)
                if fallback is not None:
                    fb_sp = fallback.find('{%s}sp' % P_NS)
                    if fb_sp is not None:
                        return etree.tostring(fb_sp).decode()
                break
            parent = parent.getparent()
    except Exception:
        pass
    return None


def extract_alternate_content_shapes(slide):
    """Extract shapes from inside mc:AlternateContent elements.

    python-pptx does not process shapes inside AlternateContent wrappers
    (used for OMML formulas and other forward-compatible features).
    This function finds those shapes and extracts their metadata.

    Args:
        slide: python-pptx Slide object.

    Returns:
        list of shape metadata dicts for shapes inside AlternateContent.
    """
    from lxml import etree
    from ppt2md.parser.formula import find_omml_elements, convert_omml_to_latex

    P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    shapes_meta = []
    try:
        spTree = slide._element.find('.//{%s}spTree' % P_NS)
        if spTree is None:
            return shapes_meta

        # Process direct children that are AlternateContent
        for child in list(spTree):
            local = child.tag.split('}')[1] if '}' in child.tag else ''
            if local != 'AlternateContent':
                continue

            # Find the Choice branch (has the actual shape with OMML)
            choice = child.find('{%s}Choice' % MC_NS)
            if choice is None:
                continue

            # Find the sp element inside Choice
            sp = choice.find('{%s}sp' % P_NS)
            if sp is None:
                continue

            # Get shape name
            name_el = sp.find('.//{%s}cNvPr' % P_NS)
            shape_name = name_el.get('name', 'Formula') if name_el is not None else 'Formula'

            # Get position from spPr/xfrm
            spPr = sp.find('{%s}spPr' % P_NS)
            if spPr is None:
                continue
            xfrm = spPr.find('{%s}xfrm' % A_NS)
            if xfrm is None:
                continue
            off = xfrm.find('{%s}off' % A_NS)
            ext = xfrm.find('{%s}ext' % A_NS)
            if off is None or ext is None:
                continue

            x = int(off.get('x', 0))
            y = int(off.get('y', 0))
            w = int(ext.get('cx', 100000))
            h = int(ext.get('cy', 100000))

            # Extract OMML XML for direct injection during reverse
            A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'
            txBody = sp.find('{%s}txBody' % P_NS)
            omml_xml_parts = []
            if txBody is not None:
                for a14_m in txBody.iter('{%s}m' % A14_NS):
                    omml_xml_parts.append(etree.tostring(a14_m).decode())

            # Extract OMML and convert to LaTeX
            latex_parts = []
            omml_list = find_omml_elements(sp)
            for omml in omml_list:
                latex = convert_omml_to_latex(omml)
                if latex:
                    latex_parts.append(latex)

            formula_text = ' '.join(latex_parts) if latex_parts else ''

            # Also extract any regular text from the shape
            text_parts = []
            if txBody is not None:
                for t_el in txBody.iter('{%s}t' % A_NS):
                    if t_el.text:
                        text_parts.append(t_el.text)

            meta = {
                "name": shape_name,
                "type": "AUTO_SHAPE (1)",
                "position": {"x": x, "y": y},
                "size": {"width": w, "height": h},
                "rotation": 0,
                "auto_shape_type": "RECTANGLE (1)",
                "fill": {"type": "none"},
                "text": {"paragraphs": []},
                "_is_formula": True,
                "_omml_xml": omml_xml_parts,
                "_fallback_xml": _get_fallback_xml(sp),
            }

            # Build text content: combine regular text with formula
            para_info = {"level": 0, "alignment": None, "runs": []}

            if text_parts:
                run_info = {"text": ''.join(text_parts)}
                para_info["runs"].append(run_info)

            if formula_text:
                # Placeholder text that indicates a formula
                formula_display = "$%s$" % formula_text
                run_info_f = {"text": formula_display}
                para_info["runs"].append(run_info_f)

            if para_info["runs"]:
                meta["text"]["paragraphs"].append(para_info)

            shapes_meta.append(meta)

    except Exception:
        pass

    return shapes_meta

    # Auto shape type
    try:
        if shape.shape_type == 1:  # MSO_SHAPE_TYPE.AUTO_SHAPE
            meta["auto_shape_type"] = str(shape.auto_shape_type)
    except (AttributeError, ValueError):
        pass

    # Placeholder info
    if shape.is_placeholder:
        meta["placeholder"] = {
            "idx": shape.placeholder_format.idx,
            "type": str(shape.placeholder_format.type),
        }

    # Fill
    fill_info = _get_fill_info(shape)
    if fill_info:
        meta["fill"] = fill_info

    # Line
    line_info = _get_line_info(shape)
    if line_info:
        meta["line"] = line_info

    # Body properties (wrap, autofit)
    body_props = _get_body_props(shape)
    if body_props:
        meta["body_props"] = body_props

    # Text
    text_info = _get_text_info(shape)
    if text_info:
        meta["text"] = text_info

    # Image
    if hasattr(shape, "image") and shape.image is not None:
        meta["image"] = {
            "content_type": shape.image.content_type,
            "blob_size": len(shape.image.blob),
        }

    # Table
    if hasattr(shape, "has_table") and shape.has_table:
        table = shape.table
        meta["table"] = {
            "rows": len(table.rows),
            "cols": len(table.columns),
        }

    # Chart
    if hasattr(shape, "has_chart") and shape.has_chart:
        meta["chart"] = {
            "type": str(shape.chart.chart_type),
        }

    return meta
