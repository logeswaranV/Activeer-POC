import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType, NullableBool  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]

from Components.markdown_tools import add_markdown_to_text_frame


# Default table border color if theme doesn't override it.
TABLE_BORDER = Color.from_argb(255, 38, 50, 125)  # navy line (default)

TABLE_THEMES = {
    "vertical_accent": {
        "header_align": "center_center",
        "label_align": "center_center",
        "body_align": "top_left",
        "header_font_bold": True,
        "label_font_bold": True,
        "label_colors": ["#223073", "#F45826", "#28BEB0"],
        "border_color": "#26327D",
        "label_border_color": "#26327D",
        "row_line_color": "#26327D",
        "header_bg": "#FFFFFF",
        "header_text_color": "#000000",
        "label_text_color": "#FFFFFF",
        "header_font_size": 14,
        "body_font_size": 16,
    }
    ,
    "vertical_accent_2": {
        "header_align": "center_center",
        "label_align": "center_center",
        "body_align": "top_left",
        "header_font_bold": True,
        "label_font_bold": True,
        "label_colors": ["#223073", "#F45826", "#28BEB0"],
        "border_color": "#26327D",
        "label_border_color": "#26327D",
        "row_line_color": "#26327D",
        "header_bg": "#E5E5E5",
        "header_text_color": "#000000",
        "label_text_color": "#FFFFFF",
        "header_font_size": 14,
        "body_font_size": 16,
    },
    "header_accent": {
        "header_alignments": ["center_left", "center_center"],
        "body_alignments": ["center_left", "center_center"],
        "header_font_bold": True,
        "label_font_bold": False,
        "label_colors": [],
        "border_color": "#C7C9CC",
        "row_line_color": "#C7C9CC",
        "header_bg": "#1E2F6E",
        "header_text_color": "#FFFFFF",
        "body_text_color": "#000000",
        "zebra_colors": ["#E6E6E6", "#F4F4F4"],
        "header_font_size": 12,
        "body_font_size": 11,
    },
}


def _color_from_hex(value: str | None, fallback: Color) -> Color:
    # Convert a #RRGGBB hex string to an Aspose color.
    if not value:
        return fallback
    text = value.strip().lstrip("#")
    if len(text) != 6:
        return fallback
    try:
        r = int(text[0:2], 16)
        g = int(text[2:4], 16)
        b = int(text[4:6], 16)
    except ValueError:
        return fallback
    return Color.from_argb(255, r, g, b)


def _set_cell_border(cell: slides.Cell, color: Color, width: float = 1.2) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    # Apply a border to all cell edges.
    for border in (
        cell.cell_format.border_top,
        cell.cell_format.border_bottom,
        cell.cell_format.border_left,
        cell.cell_format.border_right,
    ):
        border.fill_format.fill_type = FillType.SOLID
        border.fill_format.solid_fill_color.color = color
        border.width = width


def _set_cell_fill(cell: slides.Cell, fill: Color | None) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    # Fill the cell background.
    if fill is None:
        cell.cell_format.fill_format.fill_type = FillType.NO_FILL
    else:
        cell.cell_format.fill_format.fill_type = FillType.SOLID
        cell.cell_format.fill_format.solid_fill_color.color = fill


def _set_cell_text(
    cell: slides.Cell,
    text: str,
    size: int,
    bold: bool,
    color: Color,
    align: str = "center",
) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    # Set text content and alignment for a single cell.
    tf = cell.text_frame
    tf.text = text
    tf.text_frame_format.wrap_text = NullableBool.TRUE
    tf.text_frame_format.margin_left = 0
    tf.text_frame_format.margin_right = 0
    tf.text_frame_format.margin_top = 0
    tf.text_frame_format.margin_bottom = 0
    if align == "top_left":
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.TOP
        tf.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.LEFT
    elif align == "center_left":
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER
        tf.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.LEFT
    elif align == "center_right":
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER
        tf.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.RIGHT
    elif align == "center_center":
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER
        tf.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.CENTER
    else:
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER
        tf.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.CENTER
    for portion in tf.paragraphs[0].portions:
        portion.portion_format.font_height = size
        portion.portion_format.font_bold = NullableBool.TRUE if bold else NullableBool.FALSE
        portion.portion_format.fill_format.fill_type = FillType.SOLID
        portion.portion_format.fill_format.solid_fill_color.color = color


def add_table(
    slide: slides.ISlide,
    x: float,
    y: float,
    width: float,
    height: float,
    headers: list[str],
    rows: list[list[str]],
    table_type: str | None = None,
    theme: str | None = None,
    label_width: float = 76,
    header_height: float = 54,
    body_font_size: int = 16,
    heading_font_size: int = 16,
) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    # Render a table with theme-aware styling and alignment.
    theme_key = theme or "vertical_accent"
    theme_spec = TABLE_THEMES.get(theme_key, TABLE_THEMES["vertical_accent"])
    border_color = _color_from_hex(theme_spec.get("border_color"), TABLE_BORDER)
    label_border_color = _color_from_hex(theme_spec.get("label_border_color"), border_color)
    row_line_color = _color_from_hex(theme_spec.get("row_line_color"), border_color)
    header_bg = _color_from_hex(theme_spec.get("header_bg"), Color.white)
    header_text_color = _color_from_hex(theme_spec.get("header_text_color"), Color.black)
    label_text_color = _color_from_hex(theme_spec.get("label_text_color"), Color.white)
    body_text_color = _color_from_hex(theme_spec.get("body_text_color"), Color.black)
    label_colors = theme_spec.get("label_colors") or []
    zebra_colors = theme_spec.get("zebra_colors") or []
    header_font_size = int(theme_spec.get("header_font_size", heading_font_size))
    body_font_size = int(theme_spec.get("body_font_size", body_font_size))
    header_alignments = theme_spec.get("header_alignments") or []
    body_alignments = theme_spec.get("body_alignments") or []
    max_row_len = max((len(row) for row in rows), default=0)
    column_count = max(len(headers), max_row_len)
    if column_count == 0:
        return

    header_rows = 1 if headers else 0
    total_rows = header_rows + len(rows)

    if table_type == "data_listing":
        if column_count == 1:
            col_widths = [width]
        elif column_count == 2:
            col_widths = [width * 0.72, width * 0.28]
        else:
            col_widths = [width / column_count] * column_count
    else:
        if column_count == 1:
            col_widths = [width]
        else:
            data_width = max(10, width - label_width)
            remaining = max(1, column_count - 1)
            col_widths = [label_width] + [data_width / remaining] * remaining

    if total_rows > 0:
        data_row_height = (height - (header_height if header_rows else 0)) / max(1, len(rows))
    else:
        data_row_height = height
    row_heights = [header_height] if header_rows else []
    row_heights.extend([data_row_height] * len(rows))

    table = slide.shapes.add_table(x, y, col_widths, row_heights)

    for r in range(total_rows):
        for c in range(column_count):
            cell = table.rows[r][c]
            _set_cell_border(cell, border_color, 1.2)
            _set_cell_fill(cell, None)

    if headers:
        for c in range(column_count):
            header_cell = table.rows[0][c]
            text = headers[c] if c < len(headers) else ""
            _set_cell_fill(header_cell, header_bg)
            _set_cell_text(
                header_cell,
                text,
                header_font_size,
                theme_spec.get("header_font_bold", True),
                header_text_color,
                align=(
                    header_alignments[c]
                    if c < len(header_alignments)
                    else theme_spec.get("header_align", "center_center")
                ),
            )
            header_cell.text_frame.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER
            for paragraph in header_cell.text_frame.paragraphs:
                if c < len(header_alignments):
                    if header_alignments[c] == "center_left":
                        paragraph.paragraph_format.alignment = slides.TextAlignment.LEFT
                    elif header_alignments[c] == "center_right":
                        paragraph.paragraph_format.alignment = slides.TextAlignment.RIGHT
                    else:
                        paragraph.paragraph_format.alignment = slides.TextAlignment.CENTER
                else:
                    paragraph.paragraph_format.alignment = slides.TextAlignment.CENTER
                paragraph.paragraph_format.space_before = 0
                paragraph.paragraph_format.space_after = 0
            header_cell.text_frame.text_frame_format.margin_left = 0
            header_cell.text_frame.text_frame_format.margin_right = 0
            header_cell.text_frame.text_frame_format.margin_top = 0
            header_cell.text_frame.text_frame_format.margin_bottom = 0

    for idx, row in enumerate(rows):
        r = idx + header_rows
        for c in range(column_count):
            cell_text = row[c] if c < len(row) else ""
            cell = table.rows[r][c]

            cell.text_frame.text_frame_format.margin_left = 0
            cell.text_frame.text_frame_format.margin_right = 0
            cell.text_frame.text_frame_format.margin_top = 0
            cell.text_frame.text_frame_format.margin_bottom = 0
            cell.text_frame.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER

            if table_type == "data_listing":
                if zebra_colors:
                    zebra_index = idx % len(zebra_colors)
                    _set_cell_fill(cell, _color_from_hex(zebra_colors[zebra_index], Color.white))
                _set_cell_text(
                    cell,
                    cell_text,
                    body_font_size,
                    False,
                    body_text_color,
                    align=(
                        body_alignments[c]
                        if c < len(body_alignments)
                        else "center_left"
                    ),
                )
                cell.text_frame.text_frame_format.anchoring_type = slides.TextAnchorType.TOP
                cell.text_frame.text_frame_format.wrap_text = NullableBool.TRUE
                cell.text_frame.text_frame_format.margin_left = 6
                cell.text_frame.text_frame_format.margin_right = 6
                cell.text_frame.text_frame_format.margin_top = 0
                cell.text_frame.text_frame_format.margin_bottom = 0
                for paragraph in cell.text_frame.paragraphs:
                    if c < len(body_alignments):
                        if body_alignments[c] == "center_left":
                            paragraph.paragraph_format.alignment = slides.TextAlignment.LEFT
                        elif body_alignments[c] == "center_right":
                            paragraph.paragraph_format.alignment = slides.TextAlignment.RIGHT
                        else:
                            paragraph.paragraph_format.alignment = slides.TextAlignment.CENTER
                    else:
                        paragraph.paragraph_format.alignment = slides.TextAlignment.LEFT
                    paragraph.paragraph_format.space_before = 0
                    paragraph.paragraph_format.space_after = 0
                    if hasattr(paragraph.paragraph_format, "line_spacing"):
                        paragraph.paragraph_format.line_spacing = body_font_size
                continue

            if c == 0:
                if label_colors:
                    color_index = idx % len(label_colors)
                    label_color = _color_from_hex(label_colors[color_index], Color.from_argb(255, 213, 76, 28))
                else:
                    label_color = Color.from_argb(255, 213, 76, 28)
                _set_cell_fill(cell, label_color)
                _set_cell_text(
                    cell,
                    cell_text,
                    body_font_size,
                    theme_spec.get("label_font_bold", True),
                    label_text_color,
                    align=theme_spec.get("label_align", "center_center"),
                )
                cell.text_frame.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.paragraph_format.alignment = slides.TextAlignment.CENTER
                    paragraph.paragraph_format.space_before = 0
                    paragraph.paragraph_format.space_after = 0
                cell.text_frame.paragraphs[0].paragraph_format.space_before = 0
                cell.text_frame.paragraphs[0].paragraph_format.space_after = 0
            else:
                add_markdown_to_text_frame(
                    cell.text_frame,
                    cell_text,
                    body_size=body_font_size,
                    heading_size=header_font_size,
                    bullet_gap=0,
                    paragraph_gap=0,
                    bullet_symbol="\u2022",
                    align="left",
                    anchor="top" if theme_spec.get("body_align") == "top_left" else "center",
                )
                cell.text_frame.text_frame_format.anchoring_type = slides.TextAnchorType.TOP
                cell.text_frame.text_frame_format.wrap_text = NullableBool.TRUE
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.paragraph_format.alignment = slides.TextAlignment.LEFT
                    paragraph.paragraph_format.space_before = 0
