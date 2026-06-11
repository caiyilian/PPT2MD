"""Chart data extraction from PPTX slides."""


def extract_chart_as_markdown(chart):
    """Convert a python-pptx Chart to Markdown table.

    Args:
        chart: python-pptx Chart object.

    Returns:
        str: Markdown table with chart data.
    """
    if not chart.has_data:
        return "*Chart has no data*"

    plot = chart.plots[0]
    series_list = plot.series

    # Build header row from category names
    categories = list(plot.categories) if hasattr(plot, "categories") else []

    # Build table
    lines = []

    # Header: Series name | Cat1 | Cat2 | ...
    header_parts = ["Series"]
    for cat in categories:
        header_parts.append(str(cat))
    lines.append("| " + " | ".join(header_parts) + " |")
    lines.append("| " + " | ".join(["---"] * len(header_parts)) + " |")

    # Data rows
    for series in series_list:
        row_parts = [series.format.replace("{0}", "Series")]
        values = series.values
        for val in values:
            row_parts.append(str(val))
        lines.append("| " + " | ".join(row_parts) + " |")

    return "\n".join(lines)


def extract_charts_from_slide(slide):
    """Extract all charts from a slide as Markdown.

    Args:
        slide: python-pptx Slide object.

    Returns:
        list of dict with shape_index and markdown chart string.
    """
    charts = []
    for i, shape in enumerate(slide.shapes):
        if not shape.has_chart:
            continue
        md = extract_chart_as_markdown(shape.chart)
        charts.append({
            "shape_index": i,
            "markdown": md,
            "chart_type": str(shape.chart.chart_type),
        })
    return charts
