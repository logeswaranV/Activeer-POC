import matplotlib.pyplot as plt
import aspose.slides as slides
from aspose.pydrawing import Color
from aspose.slides import FillType
from io import BytesIO
from matplotlib.patches import Rectangle
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from main import SlideObject


CARD_PADDING = 12
INCH_TO_PT = 72

DESIRED_CHART_WIDTH_IN = 3.3
DESIRED_CHART_HEIGHT_IN = 4.0

DONUT_COLORS = ["#27C1B5", "#10205E", "#F15A24"]
BAR_CHART_COLOR = "#10205E"

WIDTH_SCALE = 1.3
HEIGHT_SCALE = 0.88
DONUT_SCALE = 1.55
DONUT_WIDTH_SCALE = 1.3
DONUT_HEIGHT_SCALE = 1.35
def _apply_card_shadow(card: slides.IShape) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    effect = card.effect_format
    effect.enable_outer_shadow_effect()
    shadow = effect.outer_shadow_effect
    shadow.blur_radius = 4
    shadow.distance = 3
    shadow.direction = 45
    shadow.shadow_color.color = Color.from_argb(102, 0, 0, 0)

def _add_card_background(
    slide: slides.ISlide,
    x: float,
    y: float,
    width: float,
    height: float,
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
    card.fill_format.solid_fill_color.color = Color.from_argb(255, 242, 242, 242)
    card.line_format.fill_format.fill_type = FillType.SOLID
    card.line_format.fill_format.solid_fill_color.color = Color.from_argb(
        255, 204, 204, 204
    )
    card.line_format.width = 1.2
    _apply_card_shadow(card)
    return card

def add_graph(
    slide_object: "SlideObject",
    aggregation_payload: dict,
    fallback_name: str,
) -> None:
    """Add a Matplotlib-rendered chart inside an Aspose card."""
    if not aggregation_payload:
        return
    card_height = slide_object.get_chart_height()
    x, y = slide_object.get_next_chart_position(card_height)
    card = _add_card_background(
        slide_object.aspose_object,
        x,
        y,
        slide_object.chart_width,
        card_height,
    )
    graph_width = slide_object.chart_width + CARD_PADDING * 2
    graph_height = max(0, card_height - CARD_PADDING * 2)
    graph_x = x - CARD_PADDING
    graph_y = y + (card_height - graph_height) / 2
    chart_title_text = aggregation_payload.get("name", fallback_name).title()
    title_frame = card.text_frame
    title_frame.text = chart_title_text
    title_frame.text_frame_format.anchoring_type = slides.TextAnchorType.TOP
    title_frame.text_frame_format.margin_top = 5
    title_frame.paragraphs[0].paragraph_format.alignment = (
        slides.TextAlignment.CENTER
    )
    for portion in title_frame.paragraphs[0].portions:
        portion.portion_format.font_height = 18
        portion.portion_format.font_bold = slides.NullableBool.TRUE
        portion.portion_format.fill_format.fill_type = slides.FillType.SOLID
        portion.portion_format.fill_format.solid_fill_color.color = Color.black
    if aggregation_payload.get("chartType") == "donut_chart":
        graph_width_in = graph_width / INCH_TO_PT
        graph_height_in = graph_height / INCH_TO_PT

        width_in = (
            graph_width_in * WIDTH_SCALE * DONUT_WIDTH_SCALE
        )
        height_in = (
            graph_height_in * HEIGHT_SCALE * DONUT_HEIGHT_SCALE
        )
    else:
        width_in = (
            min(DESIRED_CHART_WIDTH_IN, graph_width / INCH_TO_PT)
            * WIDTH_SCALE
        )
        height_in = (
            min(DESIRED_CHART_HEIGHT_IN, graph_height / INCH_TO_PT)
            * HEIGHT_SCALE
        )
    final_w_scale = WIDTH_SCALE
    final_h_scale = HEIGHT_SCALE
    chart_bytes = _render_chart_image(
        aggregation_payload,
        width_in,
        height_in,
    )
    final_w = graph_width
    final_h = min(graph_height, graph_height * final_h_scale)
    shift_left_offset = 0
    centered_y = graph_y + (graph_height - final_h) / 2
    image = slide_object.aspose_object.presentation.images.add_image(chart_bytes)
    frame = slide_object.aspose_object.shapes.add_picture_frame(
        slides.ShapeType.RECTANGLE,
        graph_x + shift_left_offset,
        centered_y,
        final_w,
        final_h,
        image,
    )
    frame.line_format.fill_format.fill_type = FillType.NO_FILL
    frame.line_format.width = 0

def _render_chart_image(
    payload: dict,
    width_in: float,
    height_in: float,
) -> BytesIO:
    aggregations = payload.get("aggregations", {})
    labels = list(aggregations.keys())
    values = list(aggregations.values())
    chart_type = payload.get("chartType", "horizontal_bar_chart")
    fig, ax = plt.subplots(figsize=(width_in, height_in), dpi=150)
    fig.patch.set_alpha(0)
    if chart_type == "donut_chart":
        ax.axis("off")
        ax.set_frame_on(False)
        ax.set_facecolor("none")
        ax.patch.set_alpha(0)
        for spine in ax.spines.values():
            spine.set_visible(False)
        fig.patch.set_visible(False)
        colors = [
            DONUT_COLORS[i % len(DONUT_COLORS)]
            for i in range(len(values))
        ]
        wedges, _, autotexts = ax.pie(
            values,
            labels=None,
            startangle=90,
            colors=colors,
            autopct="%d%%",
            pctdistance=0.7,
            textprops={
                "color": "white",
                "fontweight": "bold",
                "fontsize": 12,
            },
            wedgeprops=dict(
                width=0.55,
                edgecolor="white",
                linewidth=2,
            ),
        )
        centre_circle = plt.Circle((0, 0), 0.25, fc="white")
        ax.add_artist(centre_circle)
        ax.set_aspect("equal")
        legend_handles = [
            Rectangle(
                (0, 0),
                1,
                1,
                facecolor=colors[i],
                edgecolor="none",
            )
            for i in range(len(labels))
        ]
        legend = ax.legend(
            legend_handles,
            labels,
            loc="lower center",
            bbox_to_anchor=(0.5, -0.32),
            ncol=1,
            frameon=False,
            handletextpad=0.6,
            handlelength=1.1,
            labelspacing=1.0,
        )
        for text in legend.get_texts():
            text.set_fontweight("light")
            text.set_color("#666666")
            text.set_fontsize(12)

        n_count = sum(values)

        fig.text(
            0.90,
            0.08,
            f"n={n_count}",
            ha="right",
            va="bottom",
            fontsize=10,
            color="#666666",
            fontweight="light",
        )

        fig.subplots_adjust(bottom=0.25, top=0.85)

    else:
        bar_height = 0.4

        ax.barh(
            labels,
            values,
            color=BAR_CHART_COLOR,
            height=bar_height,
        )

        ax.invert_yaxis()
        ax.spines["right"].set_visible(False)
        ax.spines["top"].set_visible(False)
        x_label_style = {
            "fontweight": 300,
            "fontsize": 12,
            "fontfamily": "sans-serif",
            "color": "#444444",
        }
        y_label_style = {
            "fontweight": 400,
            "fontsize": 12,
            "fontfamily": "sans-serif",
            "color": "#000000",
        }

        ax.set_xlabel(
            payload.get("count_label", "Value"),
            **x_label_style,
        )

        ax.set_ylabel(
            payload.get("bucket_label", "Category"),
            **y_label_style,
        )

        ax.tick_params(axis="both", labelsize=11)

        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_fontweight(300)
            label.set_color("#666666")

        ax.xaxis.set_ticks_position("bottom")
        ax.tick_params(axis="x", which="both", length=0)

        max_value = max(values) if values else 0

        for idx, val in enumerate(values):
            ax.text(
                val + max_value * 0.02,
                idx,
                str(val),
                va="center",
                fontweight="bold",
                color="black",
                fontsize=10,
            )

        ax.margins(y=0.1)
        ax.set_xlim(0, max_value * 1.1 if max_value > 0 else 1)

        fig.subplots_adjust(
            left=0.30,
            right=0.92,
            top=0.90,
            bottom=0.15,
        )
    buf = BytesIO()
    fig.savefig(buf, format="png", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf
