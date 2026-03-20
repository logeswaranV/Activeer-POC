import matplotlib.pyplot as plt
from io import BytesIO
from matplotlib.patches import Rectangle
from typing import TYPE_CHECKING

from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

from Components.utils import pt, rgb, set_no_fill, set_solid_fill, set_no_line, to_hex

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


def _add_card_background(
    slide,
    x: float,
    y: float,
    width: float,
    height: float,
) -> None:
    """Draw a rounded-rectangle card background with a shadow."""
    from pptx.oxml.ns import qn
    from lxml import etree

    # Draw the main card
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Pt(x - CARD_PADDING),
        Pt(y - CARD_PADDING),
        Pt(width + CARD_PADDING * 2),
        Pt(height + CARD_PADDING * 2),
    )
    card.name = f"ChartCard_{int(x)}_{int(y)}"

    # Solid light-grey fill
    set_solid_fill(card, RGBColor(248, 248, 248))

    # Thin border
    card.line.color.rgb = RGBColor(215, 215, 215)
    card.line.width = Pt(0.75)

    # ADD SHADOW via XML
    # p:sp -> p:spPr -> a:effectLst -> a:outerShdw
    spPr = card._element.find(qn("p:spPr"))
    if spPr is None:
        spPr = etree.SubElement(card._element, qn("p:spPr"))
    effectLst = etree.SubElement(spPr, qn("a:effectLst"))
    outerShdw = etree.SubElement(effectLst, qn("a:outerShdw"))
    outerShdw.set("blurRad", "50800") # ~4pt
    outerShdw.set("dist", "25400")    # ~2pt
    outerShdw.set("dir", "2700000")   # 45 degrees
    outerShdw.set("algn", "tl")
    outerShdw.set("rotWithShape", "0")
    
    srgbClr = etree.SubElement(outerShdw, qn("a:srgbClr"))
    srgbClr.set("val", "000000")
    alpha = etree.SubElement(srgbClr, qn("a:alpha"))
    alpha.set("val", "15000") # 15% transparency

    return card


def add_graph(
    slide_object: "SlideObject",
    aggregation_payload: dict,
    fallback_name: str,
) -> None:
    """Add a Matplotlib-rendered chart inside a card on the slide."""
    if not aggregation_payload:
        return

    slide = slide_object.slide
    card_height = slide_object.get_chart_height()
    x, y = slide_object.get_next_chart_position(card_height)

    card = _add_card_background(slide, x, y, slide_object.chart_width, card_height)

    graph_width = slide_object.chart_width + CARD_PADDING * 2
    graph_height = max(0, card_height - CARD_PADDING * 2)
    graph_x = x - CARD_PADDING
    graph_y = y + (card_height - graph_height) / 2

    # ── Card title ────────────────────────────────────────────────────────────
    # Do NOT use .title() as it breaks acronyms like 'HER2' or 'BC'
    chart_title_text = aggregation_payload.get("name", fallback_name)
    
    # Use a separate text box for the title to avoid overlapping with the image
    title_height = 50
    title_shape = slide.shapes.add_textbox(
        Pt(x),
        Pt(y),
        Pt(slide_object.chart_width),
        Pt(title_height),
    )
    tf = title_shape.text_frame
    tf.word_wrap = True
    title_lines = [line.strip() for line in chart_title_text.split('\n')]
    tf.text = title_lines[0]

    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    for run in para.runs:
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = RGBColor(64, 64, 64)

    if len(title_lines) > 1:
        for extra_line in title_lines[1:]:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = extra_line
            run.font.size = Pt(9)
            run.font.bold = False
            run.font.color.rgb = RGBColor(80, 80, 80)

    # ── Size figure ───────────────────────────────────────────────────────────
    is_donut = aggregation_payload.get("chartType") == "donut_chart"
    if is_donut:
        graph_width_in = graph_width / INCH_TO_PT
        graph_height_in = (graph_height - title_height) / INCH_TO_PT
        width_in = graph_width_in * WIDTH_SCALE * DONUT_WIDTH_SCALE
        height_in = graph_height_in * HEIGHT_SCALE * DONUT_HEIGHT_SCALE
    else:
        width_in = min(DESIRED_CHART_WIDTH_IN, graph_width / INCH_TO_PT) * WIDTH_SCALE
        height_in = min(DESIRED_CHART_HEIGHT_IN, (graph_height - title_height) / INCH_TO_PT) * HEIGHT_SCALE

    chart_bytes = _render_chart_image(aggregation_payload, width_in, height_in)

    # ── Insert picture frame ──────────────────────────────────────────────────
    final_w = graph_width
    final_h = graph_height - title_height - CARD_PADDING
    picture_y = y + title_height + CARD_PADDING / 2

    # ── Double check n-value for the card ─────────────────────────────────────
    n_count = sum(values := list(aggregation_payload.get("aggregations", {}).values()))
    n_text = f"n={int(n_count)}"
    n_shape = slide.shapes.add_textbox(
        Pt(x + slide_object.chart_width - 40),
        Pt(y + card_height - 25),
        Pt(60),
        Pt(20),
    )
    ntf = n_shape.text_frame
    ntf.text = n_text
    npara = ntf.paragraphs[0]
    npara.alignment = PP_ALIGN.RIGHT
    for nrun in npara.runs:
        nrun.font.size = Pt(8)
        nrun.font.color.rgb = RGBColor(120, 120, 120)

    # ── Insert picture frame ──────────────────────────────────────────────────
    slide.shapes.add_picture(
        ensure_stream(chart_bytes),
        Pt(graph_x),
        Pt(picture_y),
        width=Pt(final_w),
        height=Pt(final_h),
    )


def ensure_stream(data):
    if isinstance(data, BytesIO):
        data.seek(0)
        return data
    return BytesIO(data)

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
        colors = [DONUT_COLORS[i % len(DONUT_COLORS)] for i in range(len(values))]
        wedges, _, autotexts = ax.pie(
            values,
            labels=None,
            startangle=90,
            colors=colors,
            autopct="%d%%",
            pctdistance=0.7,
            textprops={"color": "white", "fontweight": "bold", "fontsize": 9},
            wedgeprops=dict(width=0.6, edgecolor="white", linewidth=1.0),
        )
        ax.set_aspect("equal")
        legend_handles = [
            Rectangle((0, 0), 1, 1, facecolor=colors[i], edgecolor="none")
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
            text.set_fontweight(400)
            text.set_color("#444444")
            text.set_fontsize(9)

        fig.subplots_adjust(bottom=0.28, top=0.88)

    else:
        bar_height = 0.32
        ax.barh(labels, values, color=BAR_CHART_COLOR, height=bar_height)
        ax.invert_yaxis()
        ax.spines["right"].set_visible(False)
        # ax.spines["left"].set_visible(False)
        ax.spines["top"].set_visible(False)
        ax.spines["bottom"].set_color("#CCCCCC")
        ax.spines["left"].set_color("#CCCCCC")
        x_label_style = {
            "fontweight": 400, "fontsize": 10,
            "fontfamily": "sans-serif", "color": "#666666",
        }
        y_label_style = {
            "fontweight": 400, "fontsize": 10,
            "fontfamily": "sans-serif", "color": "#000000",
        }
        ax.set_xlabel(payload.get("count_label", "Value"), **x_label_style)
        ax.set_ylabel(payload.get("bucket_label", "Category"), **y_label_style)
        ax.tick_params(axis="both", labelsize=11)
        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_fontweight(300)
            label.set_color("#666666")
        ax.xaxis.set_ticks_position("bottom")
        ax.tick_params(axis="x", which="both", length=0)
        max_value = max(values) if values else 0
        for idx, val in enumerate(values):
            ax.text(
                val + max_value * 0.02, idx, str(val),
                va="center", fontweight="bold", color="black", fontsize=10,
            )
        ax.margins(y=0.1)
        ax.set_xlim(0, max_value * 1.1 if max_value > 0 else 1)
        fig.subplots_adjust(left=0.30, right=0.92, top=0.90, bottom=0.15)

    ax.set_facecolor("none")
    buf = BytesIO()
    fig.savefig(buf, format="png", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf
