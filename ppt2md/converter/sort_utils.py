"""Shape position sorting for reading order."""


def sort_shapes_by_position(shapes, top_threshold=None):
    """Sort shapes by visual reading order (top-to-bottom, left-to-right).

    Args:
        shapes: list of python-pptx Shape objects.
        top_threshold: EMU threshold for treating shapes as same row.
            Defaults to half the slide height / 10.

    Returns:
        list of shapes sorted in reading order.
    """
    if not shapes:
        return []

    if top_threshold is None:
        # Default: shapes within ~5% of slide height are same row
        max_top = max(s.top or 0 for s in shapes)
        top_threshold = max(max_top // 20, 36000)  # At least 0.1cm

    def sort_key(shape):
        top = shape.top if shape.top is not None else 999999999
        left = shape.left if shape.left is not None else 999999999
        # Quantize top to threshold for row grouping
        row = top // top_threshold
        return (row, left)

    return sorted(shapes, key=sort_key)
