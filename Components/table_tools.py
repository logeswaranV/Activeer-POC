import re
import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType  # pyright: ignore[reportAttributeAccessIssue]


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
    content = component.get("content", "")
    if not isinstance(content, str) or not content.strip():
        return
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

    # defaults
    header_bg = header_bg or Color.navy
    header_text = header_text or Color.white
    border_color = border_color or Color.from_argb(255, 120, 120, 120)

    # styling
    # styling
    for r, row in enumerate(rows):
        for c in range(num_cols):
            cell = table.rows[r][c]
            fmt = cell.cell_format
            
            # 1. Background Coloring (Zebra Striping)
            fmt.fill_format.fill_type = FillType.SOLID
            if r == 0:
                fmt.fill_format.solid_fill_color.color = header_bg
            elif r % 2 == 0:
                # Even rows (Light Gray)
                fmt.fill_format.solid_fill_color.color = Color.from_argb(255, 222, 226, 235)
            else:
                # Odd rows (White)
                fmt.fill_format.solid_fill_color.color = Color.white

            # 2. Borders (White borders create the "grid" look against the gray)
            for border in (fmt.border_top, fmt.border_bottom, fmt.border_left, fmt.border_right):
                border.fill_format.fill_type = FillType.SOLID
                border.fill_format.solid_fill_color.color = Color.white
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
            portion.portion_format.font_height = 11  # Slightly smaller to match UI
            
            # 4. Text Color and Boldness
            portion.portion_format.fill_format.fill_type = FillType.SOLID
            if r == 0:
                portion.portion_format.font_bold = slides.NullableBool.TRUE
                portion.portion_format.fill_format.solid_fill_color.color = header_text
            else:
                portion.portion_format.font_bold = slides.NullableBool.FALSE
                portion.portion_format.fill_format.solid_fill_color.color = Color.black
                
            paragraph.portions.add(portion)
            text_frame.paragraphs.add(paragraph)
