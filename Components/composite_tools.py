import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]

from Components.table_tools import add_table
from Components.markdown_tools import add_markdown_text

# Divider color between left and right columns.
DIVIDER_COLOR = Color.from_argb(255, 214, 72, 18)


def add_two_column_components(slide_object, body) -> None:
    """Render composite slides from a nested body layout."""
    slide = slide_object.aspose_object
    margin = 24
    content_width = slide_object.slide_width - margin * 2
    content_height = slide_object.slide_height - slide_object.chart_start_y - margin
    left_width = content_width * 0.46
    right_width = content_width - left_width
    left_x = margin
    right_x = left_x + left_width
    top_y = slide_object.chart_start_y

    def is_component(value) -> bool:
        # Components are dict payloads that map to a renderer.
        return isinstance(value, dict)

    def normalize_layout(payload):
        # Normalize body into left/right/center stacks:
        # - dict => center
        # - [dict] => left only
        # - [dict, [dict, dict]] => left full + right split
        # - [[dict, dict], dict] => left split + right full
        if isinstance(payload, dict):
            return {"center": [payload], "left": [], "right": []}
        if not isinstance(payload, list) or not payload:
            return {"center": [], "left": [], "right": []}

        if len(payload) == 1:
            single = payload[0]
            if isinstance(single, list):
                return {"center": [], "left": single, "right": []}
            return {"center": [], "left": [single], "right": []}

        if len(payload) == 2:
            left_raw, right_raw = payload
            left_stack = left_raw if isinstance(left_raw, list) else [left_raw]
            right_stack = right_raw if isinstance(right_raw, list) else [right_raw]
            left_stack = [item for item in left_stack if is_component(item)]
            right_stack = [item for item in right_stack if is_component(item)]
            return {"center": [], "left": left_stack, "right": right_stack}

        left_stack = [item for item in payload if is_component(item)]
        return {"center": [], "left": left_stack, "right": []}

    layout = normalize_layout(body)
    center_stack = layout["center"]
    left_stack = layout["left"]
    right_stack = layout["right"]

    def render_stack(stack: list[dict], x: float, width: float, height: float) -> None:
        # Stack items vertically within the allotted column area.
        if not stack:
            return
        section_height = height / len(stack)
        current_y = top_y
        for item in stack:
            item_height = section_height
            content_type = item.get("content_type") or item.get("type")
            if content_type == "table":
                # Tables are rendered via the table tool.
                add_table(
                    slide,
                    x,
                    current_y,
                    width,
                    item_height,
                    item.get("headers", []),
                    item.get("rows", []),
                    table_type=item.get("table_type"),
                    theme=item.get("theme"),
                    body_font_size=16,
                    heading_font_size=16,
                )
            elif content_type == "markdown":
                # Markdown content is rendered into a text panel.
                panel = slide.shapes.add_auto_shape(
                    slides.ShapeType.RECTANGLE,
                    x,
                    current_y,
                    width,
                    item_height,
                )
                panel.fill_format.fill_type = FillType.NO_FILL
                panel.line_format.fill_format.fill_type = FillType.NO_FILL
                panel.text_frame.text_frame_format.margin_left = 12
                panel.text_frame.text_frame_format.margin_right = 12
                panel.text_frame.text_frame_format.margin_top = 10
                panel.text_frame.text_frame_format.margin_bottom = 10
                add_markdown_text(
                    panel,
                    item.get("content", ""),
                    body_size=17,
                    heading_size=17,
                    bullet_gap=4,
                    paragraph_gap=8,
                    bullet_symbol="\u25A1",
                    align="left",
                    anchor="top",
                )
            current_y += item_height

    # Centered single item.
    if center_stack:
        render_stack(center_stack, left_x, content_width, content_height)
        return

    # Left-only or right-only stacks fill the full width.
    if left_stack and not right_stack:
        render_stack(left_stack, left_x, content_width, content_height)
        return

    if right_stack and not left_stack:
        render_stack(right_stack, left_x, content_width, content_height)
        return

    # Standard two-column layout.
    render_stack(left_stack, left_x, left_width, content_height)
    render_stack(right_stack, right_x, right_width, content_height)

    # Divider between left and right.
    divider = slide.shapes.add_auto_shape(
        slides.ShapeType.LINE,
        right_x,
        top_y,
        0,
        content_height,
    )
    divider.line_format.fill_format.fill_type = FillType.SOLID
    divider.line_format.fill_format.solid_fill_color.color = DIVIDER_COLOR
    divider.line_format.width = 1.0
