"""Extract comprehensive shape metadata for round-trip conversion."""

from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


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


def _get_fill_info(shape, theme_color_map=None, image_filename_map=None, shape_path=None):
    """Extract fill information from shape XML.

    Args:
        shape: python-pptx Shape object.
        theme_color_map: Optional dict mapping scheme names to hex colors,
                        e.g. {'accent1': '5B9BD5', 'dk1': '000000'}.
                        When provided, scheme fills are resolved to absolute hex.
    """
    spPr = shape.element.find(".//{{{}}}spPr".format(P_NS))
    if spPr is None:
        return None

    solid = spPr.find("{{{}}}solidFill".format(A_NS))
    if solid is not None:
        srgb = solid.find("{{{}}}srgbClr".format(A_NS))
        if srgb is not None:
            return {
                "type": "solid",
                "color": srgb.get("val"),
                "xml": etree.tostring(solid, encoding="unicode"),
            }
        scheme = solid.find("{{{}}}schemeClr".format(A_NS))
        if scheme is not None:
            scheme_name = scheme.get("val")
            info = {
                "type": "scheme",
                "color": scheme_name,
                "xml": etree.tostring(solid, encoding="unicode"),
            }
            modifiers = _get_color_modifiers(scheme)
            if modifiers:
                info["modifiers"] = modifiers
            # Resolve scheme to absolute hex if theme color map is available
            if theme_color_map and scheme_name in theme_color_map:
                resolved = _resolve_scheme_color(
                    theme_color_map[scheme_name], modifiers
                )
                if resolved:
                    info["_resolved"] = resolved
            return info

    grad_fill = spPr.find("{{{}}}gradFill".format(A_NS))
    if grad_fill is not None:
        return {
            "type": "gradient",
            "xml": etree.tostring(grad_fill, encoding="unicode"),
        }

    no_fill = spPr.find("{{{}}}noFill".format(A_NS))
    if no_fill is not None:
        return {"type": "none", "xml": etree.tostring(no_fill, encoding="unicode")}

    blip_fill = spPr.find("{{{}}}blipFill".format(A_NS))
    if blip_fill is not None:
        info = {"type": "blip"}
        if image_filename_map and shape_path is not None:
            filename = image_filename_map.get((tuple(shape_path), "fill"))
            if filename:
                info["filename"] = filename
        return info

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


def _resolve_scheme_color(hex_color, modifiers):
    """Resolve a theme color with modifiers to an absolute hex color.

    Supports lumMod, lumOff modifiers. Returns hex string without '#'.
    """
    if not hex_color or not modifiers:
        return hex_color
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)

        # Convert to HSL
        import math
        r_norm = r / 255.0
        g_norm = g / 255.0
        b_norm = b / 255.0

        max_c = max(r_norm, g_norm, b_norm)
        min_c = min(r_norm, g_norm, b_norm)
        l = (max_c + min_c) / 2.0

        if max_c == min_c:
            s = 0.0
            h = 0.0
        else:
            d = max_c - min_c
            s = d / (2.0 - max_c - min_c) if l > 0.5 else d / (max_c + min_c)
            if max_c == r_norm:
                h = ((g_norm - b_norm) / d + (6.0 if g_norm < b_norm else 0.0)) / 6.0
            elif max_c == g_norm:
                h = ((b_norm - r_norm) / d + 2.0) / 6.0
            else:
                h = ((r_norm - g_norm) / d + 4.0) / 6.0

        if 'lumMod' in modifiers:
            l *= int(modifiers['lumMod']) / 100000.0
        if 'lumOff' in modifiers:
            l += int(modifiers['lumOff']) / 100000.0

        l = max(0.0, min(1.0, l))

        def hsl_to_rgb(h, s, l):
            if s == 0.0:
                return (int(l * 255), int(l * 255), int(l * 255))
            if l < 0.5:
                q = l * (1.0 + s)
            else:
                q = l + s - l * s
            p = 2.0 * l - q

            def hue_to_rgb(t):
                if t < 0: t += 1.0
                if t > 1: t -= 1.0
                if t < 1.0/6.0: return p + (q - p) * 6.0 * t
                if t < 1.0/2.0: return q
                if t < 2.0/3.0: return p + (q - p) * (2.0/3.0 - t) * 6.0
                return p

            r2 = int(hue_to_rgb(h + 1.0/3.0) * 255)
            g2 = int(hue_to_rgb(h) * 255)
            b2 = int(hue_to_rgb(h - 1.0/3.0) * 255)
            return (r2, g2, b2)

        r2, g2, b2 = hsl_to_rgb(h, s, l)
        return "%02X%02X%02X" % (r2, g2, b2)
    except Exception:
        return hex_color


def _get_line_info(shape):
    """Extract line/border information from shape XML."""
    ln = shape.element.find(".//{{{}}}spPr/{{{}}}ln".format(P_NS, A_NS))
    if ln is None:
        ln = shape.element.find(".//{{{}}}ln".format(A_NS))
    if ln is None:
        return None

    width = ln.get("w")
    info = {
        "width": int(width) if width else 0,
        "xml": etree.tostring(ln, encoding="unicode"),
    }

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
                    run_info["font_size"] = int(font.size)
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
            # Superscript/subscript via baseline
            try:
                rPr = run._r.find("{{{}}}rPr".format(A_NS))
                if rPr is not None:
                    strike = rPr.get("strike")
                    if strike and strike != "noStrike":
                        run_info["strikethrough"] = True
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
        for attr in ('wrap', 'vert', 'anchor', 'rtlCol'):
            val = bodyPr.get(attr)
            if val is not None:
                props[attr] = val

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


def _get_raw_relationships(shape, image_filename_map=None, shape_path=None):
    """Map relationship ids used in raw shape XML to extracted asset filenames."""
    if not image_filename_map or shape_path is None:
        return None

    rels = {}
    for blip in shape.element.findall(".//{{{}}}blip".format(A_NS)):
        rid = blip.get("{{{}}}embed".format(R_NS)) or blip.get("{{{}}}link".format(R_NS))
        if not rid:
            continue
        filename = image_filename_map.get((tuple(shape_path), "rId", rid))
        if filename:
            rels[rid] = {
                "type": "image",
                "filename": filename,
            }

    return rels or None


def _get_group_info(shape, theme_color_map=None, image_filename_map=None, shape_path=None):
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
            for idx, child in enumerate(shape.shapes):
                child_path = tuple(shape_path + (idx,)) if shape_path is not None else None
                child_meta = extract_shape_metadata(
                    child,
                    theme_color_map,
                    image_filename_map,
                    child_path,
                )
                children.append(child_meta)

        result = {"children": children}
        if coord_space:
            result["coord_space"] = coord_space
        return result
    except Exception:
        return None


def extract_shape_metadata(shape, theme_color_map=None, image_filename_map=None, shape_path=None):
    """Extract comprehensive metadata for a shape.

    Args:
        shape: python-pptx Shape object.
        theme_color_map: Optional dict mapping scheme names to hex colors.

    Returns dict with all properties needed for round-trip conversion.
    """
    meta = {
        "name": shape.name,
        "type": str(shape.shape_type),
        "raw_xml": etree.tostring(shape.element, encoding="unicode"),
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
    raw_relationships = _get_raw_relationships(shape, image_filename_map, shape_path)
    if raw_relationships:
        meta["raw_relationships"] = raw_relationships

    # Handle GROUP shapes (MSO_SHAPE_TYPE.GROUP = 6)
    if shape.shape_type == 6:
        meta["group"] = _get_group_info(
            shape,
            theme_color_map,
            image_filename_map,
            tuple(shape_path) if shape_path is not None else None,
        )
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
    fill_info = _get_fill_info(shape, theme_color_map, image_filename_map, shape_path)
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
        if image_filename_map and shape_path is not None:
            filename = image_filename_map.get(tuple(shape_path))
            if filename:
                meta["image"]["filename"] = filename

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
            seen_latex = set()
            omml_list = find_omml_elements(sp)
            for omml in omml_list:
                latex = convert_omml_to_latex(omml)
                if latex and latex not in seen_latex:
                    latex_parts.append(latex)
                    seen_latex.add(latex)

            formula_text = ' '.join(latex_parts) if latex_parts else ''

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
            # Preserve original element order for correct rendering
            para_info = {"level": 0, "alignment": None, "runs": []}
            ordered_elements = []  # tracks original order: {"type": "text", "idx": N} or {"type": "omml", "idx": N}

            if txBody is not None:
                ap = txBody.find('{%s}p' % A_NS)
                if ap is not None:
                    for child in ap:
                        local = child.tag.split('}')[1] if '}' in child.tag else child.tag
                        if local == 'r':
                            t_el = child.find('{%s}t' % A_NS)
                            if t_el is not None and t_el.text:
                                rPr = child.find('{%s}rPr' % A_NS)
                                run_info = {"text": t_el.text}
                                if rPr is not None:
                                    baseline = rPr.get("baseline")
                                    if baseline:
                                        val = int(baseline)
                                        run_info["superscript"] = val > 0
                                        run_info["subscript"] = val < 0
                                para_info["runs"].append(run_info)
                                ordered_elements.append({"type": "text", "idx": len(para_info["runs"]) - 1})
                        elif local == 'm' or local == 'f' or (local == 's' and child.tag.endswith('}m')):
                            # OMML element (a14:m wraps m:oMath)
                            ordered_elements.append({"type": "omml", "idx": -1})
                        elif local == 'endParaRPr':
                            pass

            if formula_text:
                formula_display = "$%s$" % formula_text
                run_info_f = {"text": formula_display}
                # Only add formula run if not already in runs (dedup)
                has_formula_in_runs = any(
                    r.get("text", "").startswith("$") for r in para_info["runs"]
                )
                if not has_formula_in_runs:
                    para_info["runs"].append(run_info_f)
                    ordered_elements.append({"type": "text", "idx": len(para_info["runs"]) - 1})

            if para_info["runs"]:
                para_info["_ordered_elements"] = ordered_elements
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
    fill_info = _get_fill_info(shape, theme_color_map)
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
