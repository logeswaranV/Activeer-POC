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


def _color_from_style(value: str, fallback: Color) -> Color:
    if not isinstance(value, str):
        return fallback
    hex_value = value.strip()
    if hex_value.startswith("#"):
        hex_value = hex_value[1:]
    try:
        if len(hex_value) == 6:
            r = int(hex_value[0:2], 16)
            g = int(hex_value[2:4], 16)
            b = int(hex_value[4:6], 16)
            return Color.from_argb(255, r, g, b)
    except ValueError:
        pass
    return fallback


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
    cell_bg: list[Color] | None = None,
    cell_text_color: list[Color] | None = None,
) -> None:
    content = component.get("content", "")
    rows: list[list[str]] = []
    if isinstance(content, list):
        rows = content
    elif isinstance(content, str) and content.strip():
        rows = _parse_markdown_table(content)
    if not rows:
        return

    styles = component.get("styles", {}) if isinstance(component.get("styles"), dict) else {}
    font_size = styles.get("fontSize", font_size)

    num_rows = len(rows)
    num_cols = max(len(r) for r in rows)

    if width is None or height is None or x is None or y is None:
        width = slide_object.chart_width
        height = slide_object.get_chart_height()
        x, y = slide_object.get_next_chart_position(height)
    max_height_style = styles.get("maxHeight")
    if isinstance(max_height_style, (int, float)) and max_height_style > 0:
        height = min(height, float(max_height_style))

    # Column widths: equal split unless caller provides ratios via `column_widths`
    # or styles.ratio inside the component.
    custom_widths = component.get("column_widths") or styles.get("ratio")

    if isinstance(custom_widths, list) and custom_widths and all(
        isinstance(v, (int, float)) and v > 0 for v in custom_widths
    ):
        ratios = custom_widths[:num_cols]
        if len(ratios) < num_cols:
            ratios += [1.0] * (num_cols - len(ratios))
        total = sum(ratios) or num_cols
        col_widths = [max(40.0, width * r / total) for r in ratios]
    else:
        col_widths = [width / num_cols] * num_cols

    def _lines_from_value(value: object) -> list[str]:
        if isinstance(value, list):
            flattened = "\n".join(str(v) for v in value)
        else:
            flattened = str(value)
        normalized = (
            flattened.replace("<br />", "\n")
            .replace("<br/>", "\n")
            .replace("<br>", "\n")
        )
        return [line for line in normalized.splitlines() if line.strip()] or [""]

    estimated_heights: list[float] = []
    for row in rows:
        line_count = max(len(_lines_from_value(cell)) for cell in row) if row else 1
        estimated = max(16.0, line_count * font_size * 1.15 + 4)
        estimated_heights.append(estimated)

    total_estimated = sum(estimated_heights)
    if total_estimated <= height or total_estimated == 0:
        row_heights = estimated_heights
    else:
        scale = height / total_estimated
        row_heights = [max(14.0, h * scale) for h in estimated_heights]

    table = slide_object.aspose_object.shapes.add_table(
        x,
        y,
        col_widths,
        row_heights,
    )

    # styling
    for r, row in enumerate(rows):
        for c in range(num_cols):
            cell = table.rows[r][c]
            fmt = cell.cell_format
            tf = cell.text_frame
            tf.text_frame_format.wrap_text = slides.NullableBool.TRUE
            autofit = styles.get("autofit", "shape")
            if autofit == "normal":
                tf.text_frame_format.autofit_type = slides.TextAutofitType.NORMAL
            else:
                tf.text_frame_format.autofit_type = slides.TextAutofitType.SHAPE  # shrink-to-fit inside cell
            tf.text_frame_format.margin_left = 2
            tf.text_frame_format.margin_right = 2
            tf.text_frame_format.margin_top = 2
            tf.text_frame_format.margin_bottom = 2
            # Default: anchor center vertically (normalize style value).
            anchor_value = str(styles.get("anchor", "center") or "").lower()
            if anchor_value == "top":
                tf.text_frame_format.anchoring_type = slides.TextAnchorType.TOP
            elif anchor_value == "bottom":
                tf.text_frame_format.anchoring_type = slides.TextAnchorType.BOTTOM
            else:
                tf.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER
            
            # 1. Background Coloring (Zebra Striping)
            fmt.fill_format.fill_type = FillType.SOLID
            if r == 0:
                fmt.fill_format.solid_fill_color.color = header_bg
            else:
                if cell_bg:
                    cell_index = ((r - 1) * num_cols + c) % len(cell_bg)
                    fmt.fill_format.solid_fill_color.color = cell_bg[cell_index]
                elif stripe_even is None and stripe_odd is None:
                    fmt.fill_format.solid_fill_color.color = Color.white
                else:
                    fmt.fill_format.solid_fill_color.color = stripe_even if r % 2 == 0 else (stripe_odd or Color.white)

            # 2. Borders
            for border in (fmt.border_top, fmt.border_bottom, fmt.border_left, fmt.border_right):
                border.fill_format.fill_type = FillType.SOLID
                border.fill_format.solid_fill_color.color = border_color
                border.width = 1.0

            tf.paragraphs.clear()

            cell_value = row[c] if c < len(row) else ""
            # Normalize lists and <br> to lines
            if isinstance(cell_value, list):
                lines = [str(item) for item in cell_value]
            else:
                text_val = str(cell_value).replace("<br />", "\n").replace("<br/>", "\n").replace("<br>", "\n")
                lines = text_val.splitlines() or [""]

            for line in lines:
                para = slides.Paragraph()
                # Alignment: first column left, others center
                # Default: align left; override with styles.align
                align = styles.get("align", "left")
                if align == "right":
                    para.paragraph_format.alignment = slides.TextAlignment.RIGHT
                elif align == "center":
                    para.paragraph_format.alignment = slides.TextAlignment.CENTER
                else:
                    para.paragraph_format.alignment = slides.TextAlignment.LEFT

                # Bullet if markdown-style list item
                stripped = line.lstrip()
                if stripped.startswith("- "):
                    para.paragraph_format.bullet.type = slides.BulletType.SYMBOL
                    para.paragraph_format.bullet.char = "\u2022"
                    line = stripped[2:].lstrip()

                parts = re.split(r"(\*\*.*?\*\*)", line)
                base_bold = header_bold if r == 0 else body_bold
                base_color = header_text if r == 0 else Color.black
                if r > 0 and cell_text_color:
                    idx = ((r - 1) * num_cols + c) % len(cell_text_color)
                    base_color = cell_text_color[idx]

                for part in parts:
                    if not part:
                        continue
                    is_bold = part.startswith("**") and part.endswith("**") and len(part) >= 4
                    text_part = part[2:-2] if is_bold else part
                    portion = slides.Portion(text_part)
                    pf = portion.portion_format
                    pf.font_height = font_size
                    pf.font_bold = slides.NullableBool.TRUE if (base_bold or is_bold) else slides.NullableBool.FALSE
                    pf.fill_format.fill_type = FillType.SOLID
                    pf.fill_format.solid_fill_color.color = base_color
                    para.portions.add(portion)

                tf.paragraphs.add(para)

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

    styles = component.get("styles", {}) if isinstance(component.get("styles"), dict) else {}
    font_size = styles.get("font_size") or styles.get("fontSize") or 11

    if width is None or height is None or x is None or y is None:
        width = slide_object.chart_width
        height = slide_object.get_chart_height()
        x, y = slide_object.get_next_chart_position(height)
    max_height_style = styles.get("maxHeight")
    if isinstance(max_height_style, (int, float)) and max_height_style > 0:
        height = min(height, float(max_height_style))

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

    # If HTML parsing fails, try markdown parsing.
    if rows is None:
        md_rows = _parse_markdown_table(content)
        rows = md_rows if md_rows else None

    if rows:
        header_bg_color = _color_from_style(
            styles.get("header_bg"), header_bg or Color.from_argb(255, 240, 244, 252)
        )
        header_text_color = _color_from_style(
            styles.get("header_text"), header_text or Color.from_argb(255, 16, 32, 94)
        )
        border_color = _color_from_style(
            styles.get("border_color"), border_color or Color.from_argb(255, 200, 200, 200)
        )
        cell_bg_values = styles.get("cell_bg")
        cell_bg_list: list[Color] | None = None
        if isinstance(cell_bg_values, list):
            parsed = [
                _color_from_style(val, Color.white) for val in cell_bg_values if isinstance(val, str)
            ]
            if parsed:
                cell_bg_list = parsed
        cell_text_color_values = styles.get("cell_text_color")
        cell_text_color_list: list[Color] | None = None
        if isinstance(cell_text_color_values, list):
            parsed_text = [
                _color_from_style(val, Color.black) for val in cell_text_color_values if isinstance(val, str)
            ]
            if parsed_text:
                cell_text_color_list = parsed_text
        stripe_even = Color.from_argb(255, 250, 251, 253)
        stripe_odd = Color.white
        _render_table_core(
            slide_object,
            component | {"content": rows},
            x,
            y,
            width,
            height,
            header_bg_color,
            header_text_color,
            border_color,
            stripe_even,
            stripe_odd,
            header_bold=True,
            body_bold=False,
            font_size=font_size,
            cell_bg=cell_bg_list,
            cell_text_color=cell_text_color_list,
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
    """Render meeting-info table from markdown only (no HTML parsing)."""
    content = component.get("content", "")
    if not isinstance(content, str) or not content.strip():
        return

    if width is None or height is None or x is None or y is None:
        width = slide_object.chart_width
        height = slide_object.get_chart_height()
        x, y = slide_object.get_next_chart_position(height)
        
    styles = component.get("styles", {}) if isinstance(component.get("styles"), dict) else {}

    header_bg = Color.from_argb(255, 33, 45, 106)
    header_text = Color.white
    border_color = Color.white
    stripe_even = Color.from_argb(255, 232, 232, 235)
    stripe_odd = Color.from_argb(255, 204, 205, 212)

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
