"""Tests for Phase 14: OMML Basic Formula Conversion."""

from lxml import etree

from ppt2md.parser.formula import convert_omml_to_latex, find_omml_elements

OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _make_elem(tag, text=None):
    """Helper to create an OMML element."""
    el = etree.SubElement(etree.Element("root"), "{{{}}}{}".format(OMML_NS, tag))
    if text:
        t = etree.SubElement(el, "{{{}}}t".format(OMML_NS))
        t.text = text
    return el


def test_convert_text_run():
    r = _make_elem("r")
    t = etree.SubElement(r, "{{{}}}t".format(OMML_NS))
    t.text = "x"
    assert convert_omml_to_latex(r) == "x"


def test_convert_fraction():
    f = _make_elem("f")
    num = etree.SubElement(f, "{{{}}}num".format(OMML_NS))
    nr = etree.SubElement(num, "{{{}}}r".format(OMML_NS))
    nrt = etree.SubElement(nr, "{{{}}}t".format(OMML_NS))
    nrt.text = "a"
    den = etree.SubElement(f, "{{{}}}den".format(OMML_NS))
    dr = etree.SubElement(den, "{{{}}}r".format(OMML_NS))
    drt = etree.SubElement(dr, "{{{}}}t".format(OMML_NS))
    drt.text = "b"
    assert convert_omml_to_latex(f) == "\\frac{a}{b}"


def test_convert_sqrt():
    rad = _make_elem("rad")
    e = etree.SubElement(rad, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "x"
    assert convert_omml_to_latex(rad) == "\\sqrt{x}"


def test_convert_subscript():
    ss = _make_elem("sSub")
    e = etree.SubElement(ss, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "x"
    sub = etree.SubElement(ss, "{{{}}}sub".format(OMML_NS))
    sr = etree.SubElement(sub, "{{{}}}r".format(OMML_NS))
    srt = etree.SubElement(sr, "{{{}}}t".format(OMML_NS))
    srt.text = "n"
    assert convert_omml_to_latex(ss) == "x_{n}"


def test_convert_superscript():
    ss = _make_elem("sSup")
    e = etree.SubElement(ss, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "x"
    sup = etree.SubElement(ss, "{{{}}}sup".format(OMML_NS))
    spr = etree.SubElement(sup, "{{{}}}r".format(OMML_NS))
    sprt = etree.SubElement(spr, "{{{}}}t".format(OMML_NS))
    sprt.text = "2"
    assert convert_omml_to_latex(ss) == "x^{2}"


def test_convert_delimiter():
    d = _make_elem("d")
    e = etree.SubElement(d, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "x"
    result = convert_omml_to_latex(d)
    assert "\\left(" in result
    assert "\\right)" in result
    assert "x" in result


def test_find_omml_elements():
    xml = '<root><m:oMath xmlns:m="{}"><m:r><m:t>x</m:t></m:r></m:oMath></root>'.format(OMML_NS)
    root = etree.fromstring(xml)
    found = find_omml_elements(root)
    assert len(found) == 1
