"""Tests for PPTX to HTML conversion."""

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR_TYPE
from pptx.util import Inches, Pt

from ppt2md.converter.html import convert_pptx_to_html, main as ppt2html_main


def _create_html_fixture_pptx(tmp_path):
    image_path = tmp_path / "sample.png"
    Image.new("RGB", (80, 60), "navy").save(image_path)

    prs = Presentation()
    prs.slide_width = Inches(8)
    prs.slide_height = Inches(4.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(245, 246, 250)

    box = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(0.4),
        Inches(0.35),
        Inches(2.5),
        Inches(0.8),
    )
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(248, 203, 173)
    box.line.color.rgb = RGBColor(197, 90, 17)
    box.line.width = Pt(1.5)
    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = ""
    run = paragraph.add_run()
    run.text = "HTML Export"
    run.font.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(32, 55, 100)

    ellipse = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        Inches(3.25),
        Inches(0.35),
        Inches(0.9),
        Inches(0.9),
    )
    ellipse.fill.background()
    ellipse.line.color.rgb = RGBColor(112, 48, 160)
    ellipse.line.width = Pt(2)

    slide.shapes.add_connector(
        MSO_CONNECTOR_TYPE.STRAIGHT,
        Inches(2.95),
        Inches(0.75),
        Inches(3.25),
        Inches(0.75),
    )

    slide.shapes.add_picture(
        str(image_path),
        Inches(0.5),
        Inches(1.6),
        Inches(1.4),
        Inches(1.0),
    )

    table_shape = slide.shapes.add_table(
        2,
        2,
        Inches(2.3),
        Inches(1.55),
        Inches(2.6),
        Inches(1.25),
    )
    table = table_shape.table
    table.cell(0, 0).text = "Metric"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Accuracy"
    table.cell(1, 1).text = "98%"

    formula_box = slide.shapes.add_textbox(
        Inches(5.2),
        Inches(1.55),
        Inches(2.1),
        Inches(0.6),
    )
    formula_box.text_frame.text = "$E=mc^2$"

    pptx_path = tmp_path / "fixture.pptx"
    prs.save(pptx_path)
    return pptx_path


def test_convert_pptx_to_html_embeds_core_content(tmp_path):
    pptx_path = _create_html_fixture_pptx(tmp_path)
    output_path = tmp_path / "fixture.html"

    result = convert_pptx_to_html(pptx_path, output_path)

    html = output_path.read_text(encoding="utf-8")
    assert result == output_path
    assert "HTML Export" in html
    assert "Accuracy" in html
    assert "$E=mc^2$" in html
    assert "data:image/png;base64," in html
    assert "https://cdn.jsdelivr.net/npm/katex" in html
    assert "class=\"ppt-shape\"" in html
    assert "class=\"ppt-line\"" in html
    assert "left:" in html and "top:" in html


def test_ppt2html_cli_writes_requested_output(tmp_path):
    pptx_path = _create_html_fixture_pptx(tmp_path)
    output_path = tmp_path / "cli-output.html"

    exit_code = ppt2html_main([str(pptx_path), "-o", str(output_path)])

    assert exit_code == 0
    assert output_path.exists()
    assert "cli-output" not in output_path.read_text(encoding="utf-8")
