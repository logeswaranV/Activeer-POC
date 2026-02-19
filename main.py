import json
import math
from pathlib import Path

import aspose.slides as slides
from aspose.pydrawing import Color
from aspose.slides import FillType
from Components.utils import (
    add_title,
    add_title_only,
    _remove_default_placeholders,
)
from Components.chart_tools import add_graph
from Components.map_tools import render_map_image
from Components.text_tools import render_html_into_shape, render_meeting_info_markdown, render_list_into_shape
from Components.table_tools import render_table, render_meeting_info_table

CARD_PADDING = 12
INCH_TO_PT = 72
SHAPE_MAX_HEIGHT_IN = 7
SHAPE_MAX_HEIGHT = SHAPE_MAX_HEIGHT_IN * INCH_TO_PT
CARD_MAX_HEIGHT_IN = 5.2
CARD_MAX_HEIGHT = CARD_MAX_HEIGHT_IN * INCH_TO_PT
INPUT_JSON_PATH = Path('Input.json')


def load_deck(path: Path = INPUT_JSON_PATH) -> dict:
    """Load the deck definition from Input.json."""

    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

    return data.get("deck", {})


class SlideObject:
    """State holder for a slide while we build it."""

    def __init__(
        self,
        aspose_object: slides.ISlide,
        slide_width: float,
        slide_height: float,
        chart_columns: int = 3,
        column_gap: float = 60,
        row_gap: float = 50,
        total_charts: int = 0,
        height_cap: float = CARD_MAX_HEIGHT,
    ):  # pyright: ignore[reportAttributeAccessIssue]
        self.aspose_object = aspose_object
        self.last_right_x = 0
        self.last_bottom_y = 0
        self.slide_width = slide_width
        self.slide_height = slide_height
        self.left_margin = 20
        self.chart_columns = max(1, chart_columns)
        self.column_gap = max(25, column_gap)
        self.row_gap = max(20, row_gap)
        self.total_charts = max(1, total_charts)
        self.max_rows = max(1, math.ceil(self.total_charts / self.chart_columns))
        self.chart_start_y = 120
        self.current_column = 0
        self.current_row = 0
        self.height_cap = height_cap
        self.chart_width = (
            (self.slide_width - self.left_margin * 2)
            - self.column_gap * (self.chart_columns - 1)
        ) / self.chart_columns

    def get_next_chart_position(self, chart_height: float) -> tuple[float, float]:
        if self.current_column >= self.chart_columns:
            self.current_column = 0
            self.current_row += 1
        x = self.left_margin + (self.chart_width + self.column_gap) * self.current_column
        y = self.chart_start_y + self.current_row * (chart_height + self.row_gap)
        self.current_column += 1
        self.last_bottom_y = y + chart_height
        return x, y

    def get_chart_height(self) -> float:
        """Return a per-row height that keeps all charts within the slide."""

        rows = max(1, self.max_rows)
        available_height = (
            self.slide_height
            - self.chart_start_y
            - self.row_gap * (rows - 1)
            - CARD_PADDING * 1  # reduced bottom margin
        )
        per_row = available_height / rows if available_height > 0 else 120
        per_row = min(per_row, self.height_cap)
        return max(120, per_row)

def create_slide(presentation: slides.Presentation, deck_payload: dict) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    """Build slides from the parsed deck JSON definition."""

    presentation.slide_size.set_size(
        slides.SlideSizeType.WIDESCREEN, slides.SlideSizeScaleType.MAXIMIZE
    )
    layout_slide = presentation.layout_slides[0]
    slide_width = presentation.slide_size.size.width
    slide_height = presentation.slide_size.size.height

    slide_data = sorted(deck_payload.get("slides", []), key=lambda slide: slide.get("order", 0))
    if not slide_data:
        return
    for slide_payload in slide_data:
        slide_type = slide_payload.get("slide_type")
        slide = presentation.slides.add_empty_slide(layout_slide)
        _remove_default_placeholders(slide)
        if slide_type == "title_only":
            slide_object = SlideObject(
                slide,
                slide_width,
                slide_height,
                chart_columns=1,
                column_gap=0,
                row_gap=0,
                total_charts=1,
                height_cap=SHAPE_MAX_HEIGHT,
            )
            add_title_only(slide_object, slide_payload.get("title", ""))
            continue

        components = slide_payload.get("body") or []
        chart_only = _all_charts(components)

        if chart_only:
            column_count = max(1, len(components))
            slide_object = SlideObject(
                slide,
                slide_width,
                slide_height,
                chart_columns=column_count,
                column_gap=35,
                row_gap=35,
                total_charts=max(1, len(components)),
                height_cap=CARD_MAX_HEIGHT,
            )
            slide_title = slide_payload.get("title", "")
            if slide_title:
                add_title(slide_object, slide_title)
            _add_layout_guides(slide_object, column_count)
            for component in components:
                add_graph(
                    slide_object,
                    component,
                    component.get("name", slide_title or "Chart"),
                )
        else:
            _render_manual_layout(
                presentation,
                slide,
                components,
                slide_width,
                slide_height,
                slide_payload.get("title", ""),
                slide_payload.get("column_widths"),
            )


def _add_layout_guides(slide_object: SlideObject, columns: int) -> None:
    """Draw plain background guides that divide the available width into columns."""

    slide = slide_object.aspose_object
    chart_height = slide_object.get_chart_height()
    total_gap = (columns - 1) * slide_object.column_gap
    col_width = (
        (slide_object.slide_width - slide_object.left_margin * 2 - total_gap)
        / max(1, columns)
    )

    y = slide_object.chart_start_y
    for idx in range(columns):
        x = slide_object.left_margin + idx * (col_width + slide_object.column_gap)
        guide = slide.shapes.add_auto_shape(
            slides.ShapeType.RECTANGLE,
            x,
            y,
            col_width,
            chart_height,
        )
        guide.fill_format.fill_type = FillType.NO_FILL
        guide.line_format.fill_format.fill_type = FillType.NO_FILL


def _all_charts(components: list) -> bool:
    if not components:
        return False
    return all(isinstance(c, dict) and c.get("component") == "chart" for c in components)


def _render_manual_layout(
    presentation: slides.Presentation,
    slide: slides.ISlide,
    components: list,
    slide_width: float,
    slide_height: float,
    title: str,
    column_widths: list | None = None,
) -> None:
    slide_object = SlideObject(
        slide,
        slide_width,
        slide_height,
        chart_columns=len(components),
        column_gap=35,
        row_gap=35,
        total_charts=len(components),
        height_cap=SHAPE_MAX_HEIGHT,
    )
    if title:
        add_title(slide_object, title)

    chart_height = slide_object.get_chart_height()
    total_gap = slide_object.column_gap * (len(components) - 1)
    # Equal split unless ratios are provided via column_widths.
    if isinstance(column_widths, list) and column_widths and all(isinstance(v, (int, float)) and v > 0 for v in column_widths):
        ratios = column_widths[: len(components)]
        if len(ratios) < len(components):
            ratios += [1.0] * (len(components) - len(ratios))
        total = sum(ratios) or len(components)
        widths = [
            (slide_object.slide_width - slide_object.left_margin * 2 - total_gap)
            * r
            / total
            for r in ratios
        ]
    else:
        widths = [
            (slide_object.slide_width - slide_object.left_margin * 2 - total_gap)
            / max(1, len(components))
        ] * len(components)
    base_y = slide_object.chart_start_y

    for idx, component in enumerate(components):
        col_width = widths[idx]
        x = slide_object.left_margin + sum(widths[:idx]) + slide_object.column_gap * idx
        _render_component_in_slot(slide_object, component, x, base_y, col_width, chart_height)

    slide_object.last_bottom_y = base_y + chart_height


def _render_component_in_slot(slide_object: SlideObject, component, x: float, y: float, width: float, height: float) -> None:
    if isinstance(component, list):
        items = [c for c in component if c]
        if not items:
            return
        gap = 12
        available_height = height - gap * (len(items) - 1)
        if len(items) == 2:
            heights = [available_height * 0.6, available_height * 0.4]
        else:
            per = available_height / len(items)
            heights = [per] * len(items)

        current_y = y
        for item, h in zip(items, heights):
            _render_component_in_slot(slide_object, item, x, current_y, width, h)
            current_y += h + gap
        return

    if not isinstance(component, dict):
        component = {"component": "text", "content": str(component)}

    comp_type = component.get("component")
    if comp_type == "chart":
        # Fallback to chart renderer using current slot; add_graph uses internal positioning, so temporarily override chart width/position.
        original_left = slide_object.left_margin
        original_chart_width = slide_object.chart_width
        original_chart_start_y = slide_object.chart_start_y
        slide_object.left_margin = x
        slide_object.chart_width = width
        slide_object.chart_start_y = y
        slide_object.current_column = 0
        slide_object.current_row = 0
        add_graph(slide_object, component, component.get("name", "Chart"))
        slide_object.left_margin = original_left
        slide_object.chart_width = original_chart_width
        slide_object.chart_start_y = original_chart_start_y
    elif comp_type == "map":
        map_bytes = render_map_image(component.get("content", []) or [], width=int(width), height=int(height))
        image = slide_object.aspose_object.presentation.images.add_image(map_bytes)
        frame = slide_object.aspose_object.shapes.add_picture_frame(
            slides.ShapeType.RECTANGLE,
            x,
            y,
            width,
            height,
            image,
        )
        frame.line_format.fill_format.fill_type = FillType.NO_FILL
    elif comp_type == "table":
        render_table(slide_object, component, x, y, width, height)
    elif comp_type == "meeting_info_table":
        render_meeting_info_table(slide_object, component, x, y, width, height)
    elif comp_type == "list":
        shape = slide_object.aspose_object.shapes.add_auto_shape(
            slides.ShapeType.RECTANGLE,
            x,
            y,
            width,
            height,
        )
        shape.fill_format.fill_type = FillType.NO_FILL
        shape.line_format.fill_format.fill_type = FillType.NO_FILL
        render_list_into_shape(
            shape,
            None,
            component.get("content", ""),
            component.get("styles"),
        )
    elif comp_type == "meeting_info_text":
        shape = slide_object.aspose_object.shapes.add_auto_shape(
            slides.ShapeType.RECTANGLE,
            x,
            y,
            width,
            height,
        )
        shape.fill_format.fill_type = FillType.NO_FILL
        shape.line_format.fill_format.fill_type = FillType.NO_FILL
        render_meeting_info_markdown(shape, component.get("content", ""))
    else:
        shape = slide_object.aspose_object.shapes.add_auto_shape(
            slides.ShapeType.RECTANGLE,
            x,
            y,
            width,
            height,
        )
        shape.fill_format.fill_type = FillType.NO_FILL
        shape.line_format.fill_format.fill_type = FillType.NO_FILL
        render_html_into_shape(shape, component.get("content", ""))

# Instantiate a Presentation object that represents a presentation file
deck_definition = load_deck()
with slides.Presentation() as presentation:  # pyright: ignore[reportAttributeAccessIssue]
    create_slide(presentation, deck_definition)
    presentation.save("NewPresentation.pptx", slides.export.SaveFormat.PPTX)
