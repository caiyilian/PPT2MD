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

    solid = spPr.find(".//{{{}}}solidFill".format(A_NS))
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
