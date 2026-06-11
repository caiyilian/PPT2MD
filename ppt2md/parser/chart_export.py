"""Chart image export using Windows COM (optional)."""

import os
import sys
import tempfile


COM_AVAILABLE = False
if sys.platform == "win32":
    try:
        import win32com.client
        COM_AVAILABLE = True
    except ImportError:
        pass


def export_chart_as_image(chart, output_path):
    """Export a chart as PNG image using Windows COM.

    Args:
        chart: python-pptx Chart object.
        output_path: Path to save the PNG file.

    Returns:
        bool: True if successful, False otherwise.
    """
    if not COM_AVAILABLE:
        return False

    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True

        # Create a temporary presentation with just the chart
        temp_pptx = os.path.join(tempfile.gettempdir(), "temp_chart.pptx")
        chart.part.package.save(temp_pptx)

        presentation = ppt_app.Presentations.Open(temp_pptx)
        slide = presentation.Slides(1)
        shape = slide.Shapes(1)

        shape.Export(output_path, 1)  # 1 = ppPictureFormatPNG

        presentation.Close()
        ppt_app.Quit()

        # Cleanup temp file
        try:
            os.remove(temp_pptx)
        except OSError:
            pass

        return os.path.exists(output_path)
    except Exception:
        try:
            ppt_app.Quit()
        except Exception:
            pass
        return False


def export_charts_from_slide(slide, output_dir, slide_index):
    """Export all charts from a slide as PNG images.

    Args:
        slide: python-pptx Slide object.
        output_dir: Directory to save chart images.
        slide_index: 1-based slide number for naming.

    Returns:
        list of dict with chart info and export status.
    """
    results = []
    chart_count = 0

    for shape in slide.shapes:
        if not shape.has_chart:
            continue

        chart_count += 1
        filename = "slide_{:02d}_chart_{:02d}.png".format(slide_index, chart_count)
        filepath = os.path.join(output_dir, filename)

        success = export_chart_as_image(shape.chart, filepath)
        results.append({
            "shape_index": slide.shapes.index(shape),
            "filename": filename,
            "path": filepath,
            "exported": success,
        })

    return results
