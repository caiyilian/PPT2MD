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

    # Default: return children text
    return _convert_children(element)


def _convert_children(element):
    """Convert all children of an element."""
    parts = []
    for child in element:
        local = etree.QName(child.tag).localname
        if local in ("r", "oMath", "oMathPara"):
            parts.append(convert_omml_to_latex(child))
        elif local in ("f", "rad", "sSub", "sSup", "d"):
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
