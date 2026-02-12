import aspose.pydrawing as draw  # pyright: ignore[reportMissingModuleSource]
import math
import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
import aspose.slides.charts as charts  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType, NullableBool  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides.util import SlideUtil  # pyright: ignore[reportMissingModuleSource]


INCH_TO_PT = 72
CARD_MAX_HEIGHT_IN = 5.2
CARD_MAX_HEIGHT = CARD_MAX_HEIGHT_IN * INCH_TO_PT
CHART_TYPE_MAP = {
    "horizontal_bar_chart": charts.ChartType.CLUSTERED_BAR,
    "donut_chart": charts.ChartType.DOUGHNUT,
}
DONUT_COLORS = [
    Color.from_argb(255, 237, 102, 67),  # orange-red
    Color.from_argb(255, 16, 32, 94),    # deep navy
    Color.from_argb(255, 27, 187, 166),  # aqua teal
]

BAR_DIMENSIONS_IN = (3.3, 4.0)
DONUT_DIMENSIONS_IN = (3.0, 3.51)

DONUT_LEGEND_BOTTOM_OFFSET = 36
AXIS_LABEL_FONT_HEIGHT = 10

AGGREGATION_LIBRARY = {
    "chart1": {
        "chartType": "horizontal_bar_chart",
        "name": "Years Post-Residency Treating Patients With BC",
        "bucket_label": "Years",
        "count_label": "Physicians",
        "aggregations": {
            ">25": 20,
            "21 to 25": 0,
            "16 to 20": 30,
            "11 to 15": 30,
            "7 to 10": 10,
            "4 to 6": 0,
            "0 to 3": 10,
        },
    },
    "chart2": {
        "chartType": "horizontal_bar_chart",
        "name": "Patient Location: ",
        "bucket_label": "Physicians",
        "count_label": "Patients",
        "aggregations": {
            ">50": 20,
            "31 to 50": 10,
            "21 to 30": 30,
           "11 to 20": 20,
            "6 to 10": 40,
            "1 to 5": 0,
        },
    },
    "chart3": {
        "chartType": "donut_chart",
        "name": "How comfortable are you managing ocular toxicities?​",
        "aggregations": {"Treat the patient": 20, "Refer to an optomtertist": 40, "Refer to opthomigist": 40},
    },
}


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
            - CARD_PADDING * 2
        )
        per_row = available_height / rows if available_height > 0 else 120
        per_row = min(per_row, CARD_MAX_HEIGHT)
        return max(120, per_row)


def get_aggregation_data(chart_name: str) -> dict:
    """Return aggregation payload by chart name."""

    return AGGREGATION_LIBRARY.get(chart_name, {})


def _find_existing_title_shape(slide: slides.ISlide) -> slides.IShape | None:  # pyright: ignore[reportAttributeAccessIssue]
    for placeholder_type in (
        slides.PlaceholderType.TITLE,
        slides.PlaceholderType.CENTERED_TITLE,
    ):
        shapes = SlideUtil.find_shapes_by_placeholder_type(slide, placeholder_type)
        if shapes:
            return shapes[0]

    return None


def _emphasize_title_text(title_frame: slides.TextFrame) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    if not title_frame.paragraphs:
        return

    paragraph = title_frame.paragraphs[0]
    paragraph.paragraph_format.default_portion_format.font_bold = NullableBool.TRUE
    for portion in paragraph.portions:
        portion.portion_format.font_bold = NullableBool.TRUE


def _remove_default_placeholders(slide: slides.ISlide) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    default_types = {
        slides.PlaceholderType.BODY,
        slides.PlaceholderType.SUBTITLE,
        slides.PlaceholderType.CENTERED_TITLE,
    }
    for shape in list(slide.shapes):
        placeholder = shape.placeholder
        if placeholder and placeholder.type in default_types:
            slide.shapes.remove(shape)


def add_title(slide_object: SlideObject, text: str) -> None:
    """Ensure the slide has a bold title shape with no fill."""

    slide = slide_object.aspose_object
    title_shape = _find_existing_title_shape(slide)

    if title_shape is None:
        title_shape = slide.shapes.add_auto_shape(
            slides.ShapeType.RECTANGLE, 40, 30, 640, 60  # pyright: ignore[reportAttributeAccessIssue]
        )

    title_shape.fill_format.fill_type = FillType.NO_FILL
    title_frame = title_shape.text_frame
    title_frame.text = text
    paragraph = title_frame.paragraphs[0]
    paragraph.paragraph_format.alignment = slides.TextAlignment.LEFT  # pyright: ignore[reportAttributeAccessIssue]
    for portion in paragraph.portions:
        portion.portion_format.font_height = 28
        portion.portion_format.font_bold = NullableBool.TRUE
        portion.portion_format.fill_format.fill_type = FillType.SOLID
        portion.portion_format.fill_format.solid_fill_color.color = Color.navy
    title_shape.line_format.fill_format.fill_type = FillType.NO_FILL
    _emphasize_title_text(title_frame)

    title_bottom_y = title_shape.y + title_shape.height
    slide_object.last_bottom_y = max(slide_object.last_bottom_y, title_bottom_y)
    slide_object.chart_start_y = slide_object.last_bottom_y + 20


CARD_PADDING = 12

def _apply_card_shadow(card: slides.IShape) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    effect = card.effect_format
    effect.enable_outer_shadow_effect()
    shadow = effect.outer_shadow_effect
    shadow.blur_radius = 4
    shadow.distance = 3
    shadow.direction = 45
    shadow.shadow_color.color = Color.from_argb(102, 0, 0, 0)


def _add_card_background(
    slide: slides.ISlide, x: float, y: float, width: float, height: float
) -> slides.IShape:  # pyright: ignore[reportAttributeAccessIssue]
    card = slide.shapes.add_auto_shape(
        slides.ShapeType.ROUND_CORNER_RECTANGLE,
        x - CARD_PADDING,
        y - CARD_PADDING,
        width + CARD_PADDING * 2,
        height + CARD_PADDING * 2,
    )
    card.name = f"ChartCard_{int(x)}_{int(y)}"
    card.fill_format.fill_type = FillType.SOLID
    # Background 1 darker 5% is roughly #F2F2F2
    card.fill_format.solid_fill_color.color = Color.from_argb(255, 242, 242, 242)
    card.line_format.fill_format.fill_type = FillType.SOLID
    card.line_format.fill_format.solid_fill_color.color = Color.from_argb(255, 204, 204, 204)
    card.line_format.width = 1.2
    _apply_card_shadow(card)
    return card


GRAPH_SCALE = 0.95
GRAPH_TOP_MARGIN = 8
DESIRED_CHART_WIDTH_IN = 3.3
DESIRED_CHART_HEIGHT_IN = 4.0


def add_graph(slide_object: SlideObject, aggregation_payload: dict, chart_type: str) -> None:
    """Add a real Aspose chart populated from the aggregation payload inside a card."""

    if not aggregation_payload:
        return

    # 1. Setup Dimensions and Card
    card_height = slide_object.get_chart_height()
    x, y = slide_object.get_next_chart_position(card_height)
    card = _add_card_background(
        slide_object.aspose_object,
        x,
        y,
        slide_object.chart_width,
        card_height,
    )

    # 2. Map Chart Types
    chart_type_key = aggregation_payload.get("chartType", "horizontal_bar_chart")
    chart_enum = CHART_TYPE_MAP.get(chart_type_key, charts.ChartType.CLUSTERED_BAR)
    
    dimensions_in = DONUT_DIMENSIONS_IN if chart_type_key == "donut_chart" else BAR_DIMENSIONS_IN
    max_graph_width = max(0, slide_object.chart_width - CARD_PADDING)
    max_graph_height = max(0, card_height - CARD_PADDING)
    
    graph_width = min(dimensions_in[0] * INCH_TO_PT, max_graph_width)
    graph_height = min(dimensions_in[1] * INCH_TO_PT, max_graph_height)
    graph_x = x + max(0, (slide_object.chart_width - graph_width) / 2)
    base_y = y + max(0, (card_height - graph_height) / 2)
    graph_y = min(base_y + GRAPH_TOP_MARGIN, y + card_height - graph_height)

    # 3. Create Chart Shape
    chart_shape = slide_object.aspose_object.shapes.add_chart(
        chart_enum,
        graph_x,
        graph_y,
        graph_width,
        graph_height,
    )

    # 4. Title & Legend Formatting
    chart_title_text = aggregation_payload.get("name", chart_type).title()
    title_frame = card.text_frame
    title_frame.text = chart_title_text
    title_frame.text_frame_format.anchoring_type = slides.TextAnchorType.TOP
    title_frame.text_frame_format.margin_top = 5
    title_frame.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.CENTER
    
    for portion in title_frame.paragraphs[0].portions:
        portion.portion_format.font_height = 18
        portion.portion_format.font_bold = slides.NullableBool.TRUE
        portion.portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion.portion_format.fill_format.solid_fill_color.color = Color.black

    chart_shape.has_title = False
    chart_shape.has_legend = True
    chart_shape.legend.position = charts.LegendPositionType.BOTTOM
    chart_shape.legend.text_format.portion_format.font_height = 12

    # 5. Axis Formatting
    for axis in (chart_shape.axes.horizontal_axis, chart_shape.axes.vertical_axis):
        if axis:
            axis.major_grid_lines_format.line.fill_format.fill_type = slides.FillType.NO_FILL
            axis.format.line.fill_format.fill_type = slides.FillType.SOLID
            axis.format.line.fill_format.solid_fill_color.color = Color.from_argb(255, 102, 102, 102)
            axis.has_title = False
            axis.text_format.portion_format.font_height = AXIS_LABEL_FONT_HEIGHT
            axis.text_format.portion_format.fill_format.fill_type = slides.FillType.SOLID
            axis.text_format.portion_format.fill_format.solid_fill_color.color = Color.from_argb(255, 51, 51, 51)

    # 6. Data Population
    workbook = chart_shape.chart_data.chart_data_workbook
    chart_shape.chart_data.series.clear()
    chart_shape.chart_data.categories.clear()
    workbook.clear(0)

    series_name = workbook.get_cell(0, 0, 1, aggregation_payload.get("count_label", "Value"))
    series = chart_shape.chart_data.series.add(series_name, chart_shape.type)

    aggregations = aggregation_payload.get("aggregations", {})
    for row_index, (bucket_label, bucket_value) in enumerate(aggregations.items(), start=1):
        chart_shape.chart_data.categories.add(workbook.get_cell(0, row_index, 0, bucket_label))
        value_cell = workbook.get_cell(0, row_index, 1, bucket_value)
        
        if chart_enum == charts.ChartType.DOUGHNUT:
            dp = series.data_points.add_data_point_for_doughnut_series(value_cell)
        else:
            dp = series.data_points.add_data_point_for_bar_series(value_cell)
            
        # Data Label Styling
        dp.label.data_label_format.show_value = True
        tf = dp.label.add_text_frame_for_overriding(str(value_cell.value))
        for portion in tf.paragraphs[0].portions:
            portion.portion_format.font_height = 20
            portion.portion_format.fill_format.fill_type = slides.FillType.SOLID
            if chart_type_key == "donut_chart":
                portion.portion_format.fill_format.solid_fill_color.color = Color.white
                portion.portion_format.font_bold = NullableBool.TRUE
            else:
                portion.portion_format.fill_format.solid_fill_color.color = Color.black

    # 7. Chart-Specific Styling (The Fix)
    if chart_enum == charts.ChartType.DOUGHNUT:
        chart_shape.has_legend = False
        if len(chart_shape.chart_data.series_groups) > 0:
            chart_shape.chart_data.series_groups[0].doughnut_hole_size = 50
        legend_start_x = graph_x + graph_width / 2 - 70
        legend_start_y = y + graph_height + DONUT_LEGEND_BOTTOM_OFFSET
        entry_spacing = 18
        items = list(aggregations.items())
        for idx, (label, _) in enumerate(items):
            swatch = slide_object.aspose_object.shapes.add_auto_shape(
                slides.ShapeType.RECTANGLE,
                legend_start_x,
                legend_start_y + idx * entry_spacing,
                14,
                14,
            )
            swatch.fill_format.fill_type = FillType.SOLID
            swatch.fill_format.solid_fill_color.color = DONUT_COLORS[idx % len(DONUT_COLORS)]
            swatch.line_format.fill_format.fill_type = FillType.NO_FILL
            text_box = slide_object.aspose_object.shapes.add_auto_shape(
                slides.ShapeType.RECTANGLE,
                legend_start_x + 18,
                legend_start_y + idx * entry_spacing - 2,
                110,
                16,
            )
            text_box.fill_format.fill_type = FillType.NO_FILL
            text_box.line_format.fill_format.fill_type = FillType.NO_FILL
            text_frame = text_box.text_frame
            text_frame.text = label
            text_frame.paragraphs[0].paragraph_format.alignment = slides.TextAlignment.LEFT
            text_frame.text_frame_format.wrap_text = NullableBool.FALSE
            for portion in text_frame.paragraphs[0].portions:
                portion.portion_format.font_height = 12
                portion.portion_format.font_bold = NullableBool.TRUE
                portion.portion_format.fill_format.fill_type = FillType.SOLID
                portion.portion_format.fill_format.solid_fill_color.color = Color.black
        for idx, point in enumerate(series.data_points):
            color = DONUT_COLORS[idx % len(DONUT_COLORS)]
            point.format.fill.fill_type = FillType.SOLID
            point.format.fill.solid_fill_color.color = color
            point.format.line.fill_format.fill_type = FillType.SOLID
            point.format.line.fill_format.solid_fill_color.color = Color.white
            point.format.line.width = 3.0
    else:
        series.format.fill.fill_type = slides.FillType.SOLID
        series.format.fill.solid_fill_color.color = Color.navy

def create_slide(presentation: slides.Presentation) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    """Create a slide with multiple graphs using the aggregation library."""
    presentation.slide_size.set_size(
        slides.SlideSizeType.WIDESCREEN, slides.SlideSizeScaleType.MAXIMIZE
    )
    layout_slide = presentation.layout_slides[0]
    slide = presentation.slides.add_empty_slide(layout_slide)
    slide_width = presentation.slide_size.size.width
    slide_height = presentation.slide_size.size.height
    _remove_default_placeholders(slide)

    chart_ids = ["chart1", "chart2", "chart3"]
    slide_object = SlideObject(
        slide,
        slide_width,
        slide_height,
        chart_columns=3,
        column_gap=35,
        row_gap=35,
        total_charts=len(chart_ids),
    )

    add_title(slide_object, "Attendee Demographics")
    for chart_id in chart_ids:
        add_graph(slide_object, get_aggregation_data(chart_id), chart_id)


# Instantiate a Presentation object that represents a presentation file
with slides.Presentation() as presentation:  # pyright: ignore[reportAttributeAccessIssue]
    create_slide(presentation)
    presentation.save("NewPresentation.pptx", slides.export.SaveFormat.PPTX)
