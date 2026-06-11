"""Tests for Phase 15: OMML Comprehensive Formula Conversion."""

from lxml import etree

from ppt2md.parser.formula import convert_omml_to_latex

OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _make_elem(tag, text=None):
    el = etree.SubElement(etree.Element("root"), "{{{}}}{}".format(OMML_NS, tag))
    if text:
        t = etree.SubElement(el, "{{{}}}t".format(OMML_NS))
        t.text = text
    return el


def test_convert_nary_sum():
    nary = _make_elem("nary")
    pr = etree.SubElement(nary, "{{{}}}naryPr".format(OMML_NS))
    chr_el = etree.SubElement(pr, "{{{}}}chr".format(OMML_NS))
    chr_el.set("{%s}val" % OMML_NS, "∑")
    sub = etree.SubElement(nary, "{{{}}}sub".format(OMML_NS))
    sr = etree.SubElement(sub, "{{{}}}r".format(OMML_NS))
    srt = etree.SubElement(sr, "{{{}}}t".format(OMML_NS))
    srt.text = "i=0"
    sup = etree.SubElement(nary, "{{{}}}sup".format(OMML_NS))
    spr = etree.SubElement(sup, "{{{}}}r".format(OMML_NS))
    sprt = etree.SubElement(spr, "{{{}}}t".format(OMML_NS))
    sprt.text = "n"
    e = etree.SubElement(nary, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "i"
    result = convert_omml_to_latex(nary)
    assert "\\sum" in result
    assert "i=0" in result
    assert "n" in result


def test_convert_func():
    func = _make_elem("func")
    fname = etree.SubElement(func, "{{{}}}fName".format(OMML_NS))
    fr = etree.SubElement(fname, "{{{}}}r".format(OMML_NS))
    frt = etree.SubElement(fr, "{{{}}}t".format(OMML_NS))
    frt.text = "sin"
    e = etree.SubElement(func, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "x"
    result = convert_omml_to_latex(func)
    assert "\\sin" in result
    assert "x" in result


def test_convert_bar_overline():
    bar = _make_elem("bar")
    pr = etree.SubElement(bar, "{{{}}}barPr".format(OMML_NS))
    pos = etree.SubElement(pr, "{{{}}}pos".format(OMML_NS))
    pos.set("{%s}val" % OMML_NS, "top")
    e = etree.SubElement(bar, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "x"
    result = convert_omml_to_latex(bar)
    assert "\\overline{x}" == result


def test_convert_bar_underline():
    bar = _make_elem("bar")
    pr = etree.SubElement(bar, "{{{}}}barPr".format(OMML_NS))
    pos = etree.SubElement(pr, "{{{}}}pos".format(OMML_NS))
    pos.set("{%s}val" % OMML_NS, "bot")
    e = etree.SubElement(bar, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "x"
    result = convert_omml_to_latex(bar)
    assert "\\underline{x}" == result


def test_convert_lim_low():
    ll = _make_elem("limLow")
    e = etree.SubElement(ll, "{{{}}}e".format(OMML_NS))
    er = etree.SubElement(e, "{{{}}}r".format(OMML_NS))
    ert = etree.SubElement(er, "{{{}}}t".format(OMML_NS))
    ert.text = "lim"
    lim = etree.SubElement(ll, "{{{}}}lim".format(OMML_NS))
    lr = etree.SubElement(lim, "{{{}}}r".format(OMML_NS))
    lrt = etree.SubElement(lr, "{{{}}}t".format(OMML_NS))
    lrt.text = "x→0"
    result = convert_omml_to_latex(ll)
    assert "lim" in result
    assert "x→0" in result


def test_convert_matrix():
    m = _make_elem("m")
    mr1 = etree.SubElement(m, "{{{}}}mr".format(OMML_NS))
    e1 = etree.SubElement(mr1, "{{{}}}e".format(OMML_NS))
    r1 = etree.SubElement(e1, "{{{}}}r".format(OMML_NS))
    t1 = etree.SubElement(r1, "{{{}}}t".format(OMML_NS))
    t1.text = "1"
    e2 = etree.SubElement(mr1, "{{{}}}e".format(OMML_NS))
    r2 = etree.SubElement(e2, "{{{}}}r".format(OMML_NS))
    t2 = etree.SubElement(r2, "{{{}}}t".format(OMML_NS))
    t2.text = "2"
    result = convert_omml_to_latex(m)
    assert "\\begin{matrix}" in result
    assert "1 & 2" in result
    assert "\\end{matrix}" in result
