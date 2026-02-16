import re
from io import StringIO

import pandas as pd
import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType  # pyright: ignore[reportAttributeAccessIssue]

from Components.text_tools import render_html_into_shape


def _split_row(row: str) -> list[str]:
    # handle escaped pipes \|
    row = row.strip().strip("|")
    parts = re.split(r"(?<!\\)\|", row)
    return [cell.replace("\\|", "|").strip() for cell in parts]


def _parse_markdown_table(md: str) -> list[list[str]]:
    lines = [line for line in md.splitlines() if line.strip()]
    if len(lines) < 2:
        return []
    header = _split_row(lines[0])
    data_lines = lines[2:] if set(lines[1].replace("|", "").strip("-").strip()) == set() else lines[1:]
    rows = [_split_row(line) for line in data_lines]
    return [header] + rows


def _render_table_core(
    slide_object,
    component: dict,
    x: float | None,
    y: float | None,
    width: float | None,
    height: float | None,
    header_bg: Color,
    header_text: Color,
    border_color: Color,
    stripe_even: Color | None,
    stripe_odd: Color | None,
    header_bold: bool = True,
    body_bold: bool = False,
    font_size: int = 11,
) -> None:
    content = component.get("content", "")
    rows: list[list[str]] = []
    if isinstance(content, list):
        rows = content
    elif isinstance(content, str) and content.strip():
        rows = _parse_markdown_table(content)
    if not rows:
        return

    num_rows = len(rows)
    num_cols = max(len(r) for r in rows)

    if width is None or height is None or x is None or y is None:
        width = slide_object.chart_width
        height = slide_object.get_chart_height()
        x, y = slide_object.get_next_chart_position(height)

    col_widths = [width / num_cols] * num_cols
    row_height = max(24.0, height / max(3, num_rows))
    table = slide_object.aspose_object.shapes.add_table(
        x,
        y,
        col_widths,
        [row_height] * num_rows,
    )

    # styling
    for r, row in enumerate(rows):
        for c in range(num_cols):
            cell = table.rows[r][c]
            fmt = cell.cell_format
            
            # 1. Background Coloring (Zebra Striping)
            fmt.fill_format.fill_type = FillType.SOLID
            if r == 0:
                fmt.fill_format.solid_fill_color.color = header_bg
            else:
                if stripe_even is None and stripe_odd is None:
                    fmt.fill_format.solid_fill_color.color = Color.white
                else:
                    fmt.fill_format.solid_fill_color.color = stripe_even if r % 2 == 0 else (stripe_odd or Color.white)

            # 2. Borders
            for border in (fmt.border_top, fmt.border_bottom, fmt.border_left, fmt.border_right):
                border.fill_format.fill_type = FillType.SOLID
                border.fill_format.solid_fill_color.color = border_color
                border.width = 1.0

            text_frame = cell.text_frame
            text_frame.paragraphs.clear()

            paragraph = slides.Paragraph()
            
            # 3. Alignment: Left for names, Center for counts
            if c == 0:
                paragraph.paragraph_format.alignment = slides.TextAlignment.LEFT
            else:
                paragraph.paragraph_format.alignment = slides.TextAlignment.CENTER
            
            portion = slides.Portion()
            portion.text = row[c] if c < len(row) else ""
            portion.portion_format.font_height = font_size  # Slightly smaller to match UI
            
            # 4. Text Color and Boldness
            portion.portion_format.fill_format.fill_type = FillType.SOLID
            if r == 0:
                portion.portion_format.font_bold = slides.NullableBool.TRUE if header_bold else slides.NullableBool.FALSE
                portion.portion_format.fill_format.solid_fill_color.color = header_text
            else:
                portion.portion_format.font_bold = slides.NullableBool.TRUE if body_bold else slides.NullableBool.FALSE
                portion.portion_format.fill_format.solid_fill_color.color = Color.black
                
            paragraph.portions.add(portion)
            text_frame.paragraphs.add(paragraph)

    slide_object.last_bottom_y = max(slide_object.last_bottom_y, y + height)


def render_table(
    slide_object,
    component: dict,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    header_bg: Color | None = None,
    header_text: Color | None = None,
    border_color: Color | None = None,
) -> None:
    """Render a table component; prefers HTML parsing, falls back to HTML render."""
    content = component.get("content", "")
    if not isinstance(content, str) or not content.strip():
        return

    if width is None or height is None or x is None or y is None:
        width = slide_object.chart_width
        height = slide_object.get_chart_height()
        x, y = slide_object.get_next_chart_position(height)

    # Try to parse HTML into rows/cols; fallback to HTML-in-textframe if parsing fails
    rows: list[list[str]] | None = None
    try:
        dfs = pd.read_html(StringIO(content))
        if dfs:
            df = dfs[0]
            rows = [list(df.columns)]
            rows.extend(df.astype(str).fillna("").values.tolist())
    except Exception:
        rows = None

    if rows:
        header_bg = header_bg or Color.from_argb(255, 240, 244, 252)
        header_text = header_text or Color.from_argb(255, 16, 32, 94)
        border_color = border_color or Color.from_argb(255, 200, 200, 200)
        stripe_even = Color.from_argb(255, 250, 251, 253)
        stripe_odd = Color.white
        _render_table_core(
            slide_object,
            component | {"content": rows},
            x,
            y,
            width,
            height,
            header_bg,
            header_text,
            border_color,
            stripe_even,
            stripe_odd,
            header_bold=True,
            body_bold=False,
            font_size=11,
        )
        return

    # Fallback: render the raw HTML into the shape
    shape = slide_object.aspose_object.shapes.add_auto_shape(
        slides.ShapeType.RECTANGLE,
        x,
        y,
        width,
        height,
    )
    shape.fill_format.fill_type = FillType.NO_FILL
    shape.line_format.fill_format.fill_type = FillType.NO_FILL
    render_html_into_shape(shape, content)


def render_meeting_info_table(
    slide_object,
    component: dict,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> None:
    """Legacy meeting-info table styling: navy header, light zebra, white borders."""
    header_bg = Color.navy
    header_text = Color.white
    border_color = Color.white
    stripe_even = Color.from_argb(255, 222, 226, 235)
    stripe_odd = Color.white
    _render_table_core(
        slide_object,
        component,
        x,
        y,
        width,
        height,
        header_bg,
        header_text,
        border_color,
        stripe_even,
        stripe_odd,
        header_bold=True,
        body_bold=False,
        font_size=11,
    )
