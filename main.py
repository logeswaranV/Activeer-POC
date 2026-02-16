import json
import math
from pathlib import Path

import aspose.slides as slides
from Components.utils import add_title, _remove_default_placeholders
from Components.chart_tools import add_graph
from Components.composite_tools import add_two_column_components

# Layout constants for chart card sizing.
CARD_PADDING = 12
INCH_TO_PT = 72
CARD_MAX_HEIGHT_IN = 5.2
CARD_MAX_HEIGHT = CARD_MAX_HEIGHT_IN * INCH_TO_PT

# Default deck definition file.
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
        # Compute chart card width based on slide width and column spacing.
        self.chart_width = (
            (self.slide_width - self.left_margin * 2)
            - self.column_gap * (self.chart_columns - 1)
        ) / self.chart_columns

    def get_next_chart_position(self, chart_height: float) -> tuple[float, float]:
        # Move to next row after the last column is filled.
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
        # Keep cards within the available slide height.
        available_height = (
            self.slide_height
            - self.chart_start_y
            - self.row_gap * (rows - 1)
            - CARD_PADDING * 2
        )
        per_row = available_height / rows if available_height > 0 else 120
        per_row = min(per_row, CARD_MAX_HEIGHT)
        return max(120, per_row)

def create_slide(presentation: slides.Presentation, deck_payload: dict) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    """Build slides from the parsed deck JSON definition."""

    # Configure widescreen slide size.
    presentation.slide_size.set_size(
        slides.SlideSizeType.WIDESCREEN, slides.SlideSizeScaleType.MAXIMIZE
    )
    layout_slide = presentation.layout_slides[0]
    slide_width = presentation.slide_size.size.width
    slide_height = presentation.slide_size.size.height

    # Slides are already ordered in Input.json.
    slide_data = deck_payload.get("slides", [])
    if not slide_data:
        return
    for slide_payload in slide_data:
        # Create a blank slide and strip default placeholders.
        slide = presentation.slides.add_empty_slide(layout_slide)
        _remove_default_placeholders(slide)
        components = slide_payload.get("body") or []
        chart_components = []
        layout_components = []
        if isinstance(components, list):
            chart_components = [
                component for component in components if isinstance(component, dict) and component.get("component") == "chart"
            ]

        slide_object = SlideObject(
            slide,
            slide_width,
            slide_height,
            chart_columns=3,
            column_gap=35,
            row_gap=35,
            total_charts=len(chart_components),
        )

        slide_title = slide_payload.get("title", "")
        if slide_title:
            add_title(slide_object, slide_title)

        composite_candidate = components
        is_composite = isinstance(composite_candidate, dict)
        if isinstance(composite_candidate, list):
            is_composite = any(
                isinstance(item, dict) and (item.get("content_type") or item.get("type")) in {"table", "markdown"}
                for item in composite_candidate
            ) or any(isinstance(item, list) for item in composite_candidate)

        # Composite slides are rendered from the nested body layout.
        if is_composite:
            add_two_column_components(slide_object, composite_candidate)
        else:
            # Fallback: render chart components.
            for component in chart_components:
                add_graph(
                    slide_object,
                    component,
                    component.get("name", slide_title or "Chart"),
                )

            for component in layout_components:
                pass


# Instantiate a Presentation object that represents a presentation file.
deck_definition = load_deck()
with slides.Presentation() as presentation:  # pyright: ignore[reportAttributeAccessIssue]
    create_slide(presentation, deck_definition)
    presentation.save("NewPresentation.pptx", slides.export.SaveFormat.PPTX)
