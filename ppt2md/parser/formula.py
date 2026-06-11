"""OMML to LaTeX formula conversion."""

from lxml import etree

OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _ns(tag):
    """Create a namespaced tag."""
    return "{{{}}}{}".format(OMML_NS, tag)


def _get_text(node):
    """Extract all text from an OMML run element."""
    texts = []
    for t in node.iter(_ns("t")):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def convert_omml_to_latex(element):
    """Convert an OMML element to LaTeX string.

    Args:
        element: lxml Element containing OMML math.

    Returns:
        str: LaTeX representation.
    """
    if element is None:
        return ""

    tag = etree.QName(element.tag).localname

    if tag == "oMathPara" or tag == "oMath":
        return _convert_children(element)

    if tag == "r":
        return _get_text(element)

    if tag == "f":
        return _convert_fraction(element)

    if tag == "rad":
        return _convert_radical(element)

    if tag == "sSub":
        return _convert_subscript(element)

    if tag == "sSup":
        return _convert_superscript(element)

    if tag == "d":
        return _convert_delimiter(element)

    if tag == "nary":
        return _convert_nary(element)

    if tag == "func":
        return _convert_func(element)

    if tag == "acc":
        return _convert_acc(element)

    if tag == "bar":
        return _convert_bar(element)

    if tag == "limLow":
        return _convert_lim_low(element)

    if tag == "limUpp":
        return _convert_lim_upp(element)

    if tag == "m":
        return _convert_matrix(element)

    if tag == "eqArr":
        return _convert_eq_arr(element)

    if tag == "sPre":
        return _convert_s_pre(element)

    if tag == "groupChr":
        return _convert_group_chr(element)

    # Default: return children text
    return _convert_children(element)


def _convert_children(element):
    """Convert all children of an element."""
    parts = []
    handled_tags = {"r", "oMath", "oMathPara", "f", "rad", "sSub", "sSup", "d",
                    "nary", "func", "acc", "bar", "limLow", "limUpp", "m",
                    "eqArr", "sPre", "groupChr"}
    for child in element:
        local = etree.QName(child.tag).localname
        if local in handled_tags:
            parts.append(convert_omml_to_latex(child))
        elif local == "t":
            if child.text:
                parts.append(child.text)
    return "".join(parts)


def _convert_fraction(element):
    """Convert m:f to \\frac{num}{den}."""
    num = ""
    den = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "num":
            num = _convert_children(child)
        elif local == "den":
            den = _convert_children(child)
    return "\\frac{{{}}}{{{}}}".format(num, den)


def _convert_radical(element):
    """Convert m:rad to \\sqrt{x} or \\sqrt[n]{x}."""
    deg = ""
    base = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "deg":
            deg = _convert_children(child)
        elif local == "e":
            base = _convert_children(child)
    if deg:
        return "\\sqrt[{}]{{{}}}".format(deg, base)
    return "\\sqrt{{{}}}".format(base)


def _convert_subscript(element):
    """Convert m:sSub to x_{n}."""
    base = ""
    sub = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "e":
            base = _convert_children(child)
        elif local == "sub":
            sub = _convert_children(child)
    return "{}_{{{}}}".format(base, sub)


def _convert_superscript(element):
    """Convert m:sSup to x^{n}."""
    base = ""
    sup = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "e":
            base = _convert_children(child)
        elif local == "sup":
            sup = _convert_children(child)
    return "{}^{{{}}}".format(base, sup)


def _convert_delimiter(element):
    """Convert m:d to \\left( \\right) or custom delimiters."""
    open_d = "("
    close_d = ")"
    content_parts = []

    for child in element:
        local = etree.QName(child.tag).localname
        if local == "begChr":
            open_d = child.get(_ns("val"), "(")
        elif local == "endChr":
            close_d = child.get(_ns("val"), ")")
        elif local == "e":
            content_parts.append(convert_omml_to_latex(child))

    content = ", ".join(content_parts)
    return "\\left{} {} \\right{}".format(open_d, content, close_d)


def _convert_nary(element):
    """Convert m:nary to \\sum, \\int, etc."""
    nary_chr = "∑"
    sub = ""
    sup = ""
    body = ""

    for child in element:
        local = etree.QName(child.tag).localname
        if local == "naryPr":
            chr_elem = child.find(_ns("chr"))
            if chr_elem is not None:
                nary_chr = chr_elem.get(_ns("val"), "∑")
        elif local == "sub":
            sub = _convert_children(child)
        elif local == "sup":
            sup = _convert_children(child)
        elif local == "e":
            body = _convert_children(child)

    chr_map = {
        "∑": "\\sum",
        "∏": "\\prod",
        "∫": "\\int",
        "∬": "\\iint",
        "∭": "\\iiint",
        "∮": "\\oint",
        "⋃": "\\bigcup",
        "⋂": "\\bigcap",
    }
    cmd = chr_map.get(nary_chr, "\\sum")

    result = cmd
    if sub:
        result += "_{{{}}}".format(sub)
    if sup:
        result += "^{{{}}}".format(sup)
    if body:
        result += " {}".format(body)
    return result


def _convert_func(element):
    """Convert m:func to \\sin, \\cos, etc."""
    fname = ""
    body = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "fName":
            fname = _convert_children(child)
        elif local == "e":
            body = _convert_children(child)
    # Clean up fname - remove backslash if already present
    fname = fname.strip().replace("\\", "")
    return "\\{} {}".format(fname, body)


def _convert_acc(element):
    """Convert m:acc to \\hat{x}, \\dot{x}, etc."""
    acc_chr = "^"
    body = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "accPr":
            chr_elem = child.find(_ns("chr"))
            if chr_elem is not None:
                acc_chr = chr_elem.get(_ns("val"), "^")
        elif local == "e":
            body = _convert_children(child)

    chr_map = {
        "^": "\\hat",
        "̂": "\\hat",
        "̇": "\\dot",
        "̈": "\\ddot",
        "̃": "\\tilde",
        "̄": "\\bar",
        "→": "\\vec",
    }
    cmd = chr_map.get(acc_chr, "\\hat")
    return "{}{{{}}}".format(cmd, body)


def _convert_bar(element):
    """Convert m:bar to \\overline{x} or \\underline{x}."""
    bar_pos = "top"
    body = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "barPr":
            pos_elem = child.find(_ns("pos"))
            if pos_elem is not None:
                bar_pos = pos_elem.get(_ns("val"), "top")
        elif local == "e":
            body = _convert_children(child)

    if bar_pos == "bot":
        return "\\underline{{{}}}".format(body)
    return "\\overline{{{}}}".format(body)


def _convert_lim_low(element):
    """Convert m:limLow to x_{below}."""
    body = ""
    lim = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "e":
            body = _convert_children(child)
        elif local == "lim":
            lim = _convert_children(child)
    return "{}_{{{}}}".format(body, lim)


def _convert_lim_upp(element):
    """Convert m:limUpp to x^{above}."""
    body = ""
    lim = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "e":
            body = _convert_children(child)
        elif local == "lim":
            lim = _convert_children(child)
    return "{}^{{{}}}".format(body, lim)


def _convert_matrix(element):
    """Convert m:m to \\begin{matrix}...\\end{matrix}."""
    rows = []
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "mr":
            cells = []
            for mr_child in child:
                mr_local = etree.QName(mr_child.tag).localname
                if mr_local == "e":
                    cells.append(convert_omml_to_latex(mr_child))
            rows.append(" & ".join(cells))

    body = " \\\\ ".join(rows)
    return "\\begin{{matrix}} {} \\end{{matrix}}".format(body)


def _convert_eq_arr(element):
    """Convert m:eqArr to equation array."""
    equations = []
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "e":
            equations.append(convert_omml_to_latex(child))

    body = " \\\\ ".join(equations)
    return "\\begin{{array}} {{}} {} \\end{{array}}".format(body)


def _convert_s_pre(element):
    """Convert m:sPre to pre-subscript/superscript."""
    sub = ""
    sup = ""
    body = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "sub":
            sub = _convert_children(child)
        elif local == "sup":
            sup = _convert_children(child)
        elif local == "e":
            body = _convert_children(child)

    result = ""
    if sub:
        result += "_{{{}}}".format(sub)
    if sup:
        result += "^{{{}}}".format(sup)
    result += " {}".format(body)
    return result


def _convert_group_chr(element):
    """Convert m:groupChr to \\underbrace or \\overbrace."""
    chr_val = ""
    body = ""
    for child in element:
        local = etree.QName(child.tag).localname
        if local == "groupChrPr":
            chr_elem = child.find(_ns("chr"))
            if chr_elem is not None:
                chr_val = chr_elem.get(_ns("val"), "")
        elif local == "e":
            body = _convert_children(child)

    if chr_val in ("⏟",):
        return "\\underbrace{{{}}}".format(body)
    if chr_val in ("⏞",):
        return "\\overbrace{{{}}}".format(body)
    return "\\underbrace{{{}}}".format(body)


def find_omml_elements(xml_element):
    """Find all OMML math elements in an XML tree.

    Args:
        xml_element: lxml Element to search.

    Returns:
        list of OMML elements found.
    """
    results = []
    for elem in xml_element.iter():
        local = etree.QName(elem.tag).localname
        if local in ("oMath", "oMathPara"):
            results.append(elem)
    return results
