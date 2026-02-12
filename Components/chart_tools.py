import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType, NullableBool  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides.util import SlideUtil  # pyright: ignore[reportMissingModuleSource]
import aspose.slides.charts as charts # pyright: ignore[reportMissingModuleSource]

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from main import SlideObject

CARD_PADDING = 12
GRAPH_SCALE = 0.95
GRAPH_TOP_MARGIN = 8
DESIRED_CHART_WIDTH_IN = 3.3
DESIRED_CHART_HEIGHT_IN = 4.0
DONUT_COLORS = [
    Color.from_argb(255, 237, 102, 67),  # orange-red
    Color.from_argb(255, 16, 32, 94),    # deep navy
    Color.from_argb(255, 27, 187, 166),  # aqua teal
]
INCH_TO_PT = 72
CARD_MAX_HEIGHT_IN = 5.2
CARD_MAX_HEIGHT = CARD_MAX_HEIGHT_IN * INCH_TO_PT
CHART_TYPE_MAP = {
    "horizontal_bar_chart": charts.ChartType.CLUSTERED_BAR,
    "donut_chart": charts.ChartType.DOUGHNUT,
}
BAR_DIMENSIONS_IN = (3.3, 4.0)
DONUT_DIMENSIONS_IN = (3.0, 3.51)

DONUT_LEGEND_BOTTOM_OFFSET = 36
AXIS_LABEL_FONT_HEIGHT = 10

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

def add_graph(slide_object: "SlideObject", aggregation_payload: dict, fallback_name: str) -> None:
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
    chart_title_text = aggregation_payload.get("name", fallback_name).title()
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
