"""Table extraction from PPTX slides."""


def extract_table_as_markdown(table):
    """Convert a python-pptx Table to Markdown table syntax.

    Args:
        table: python-pptx Table object.

    Returns:
        str: Markdown table string.
    """
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            text = cell.text.strip()
            # Handle multi-paragraph cells
            paras = [p.text.strip() for p in cell.text_frame.paragraphs if p.text.strip()]
            text = "<br>".join(paras) if len(paras) > 1 else (paras[0] if paras else "")
            cells.append(text)
        rows.append(cells)

    if not rows:
        return ""

    # Build Markdown table
    num_cols = len(rows[0])
    header = rows[0]
    separator = ["---"] * num_cols

    lines = []
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join(separator) + " |")
    for row in rows[1:]:
        # Pad row if needed
        while len(row) < num_cols:
            row.append("")
        lines.append("| " + " | ".join(row[:num_cols]) + " |")

    return "\n".join(lines)


def extract_tables_from_slide(slide):
    """Extract all tables from a slide as Markdown.

    Args:
        slide: python-pptx Slide object.

    Returns:
        list of dict with shape_index and markdown table string.
    """
    tables = []
    for i, shape in enumerate(slide.shapes):
        if not shape.has_table:
            continue
        md = extract_table_as_markdown(shape.table)
        tables.append({
            "shape_index": i,
            "markdown": md,
        })
    return tables
