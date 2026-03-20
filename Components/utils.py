from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE
import pptx.shapes.autoshape

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from main import SlideObject


# ── Coordinate helper ──────────────────────────────────────────────────────────
def pt(value: float) -> Emu:
    """Convert points to EMU (python-pptx's native unit)."""
    return Emu(int(value * 12700))


# ── Color helper ───────────────────────────────────────────────────────────────
def rgb(r: int, g: int, b: int) -> RGBColor:
    return RGBColor(r, g, b)


def to_hex(color: RGBColor) -> str:
    """Convert RGBColor to a 6-character hex string (e.g. 'FF0000')."""
    return "%02X%02X%02X" % (color[0], color[1], color[2])


# ── Fill helpers ───────────────────────────────────────────────────────────────
def set_no_fill(shape) -> None:
    shape.fill.background()


def set_solid_fill(shape, color: RGBColor) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def set_no_line(shape) -> None:
    shape.line.fill.background()
    # For some shapes, we need to explicitly set width to 0
    try:
        shape.line.width = 0
    except:
        pass


# ── Title helpers ──────────────────────────────────────────────────────────────

def _find_existing_title_shape(slide):
    """Return the title placeholder on *slide*, or None."""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            return ph
    return None


def _emphasize_title_text(tf) -> None:
    if not tf.paragraphs:
        return
    for run in tf.paragraphs[0].runs:
        run.font.bold = True


def add_title(slide_object: "SlideObject", text: str) -> None:
    """Ensure the slide has a bold title shape with no fill."""

    slide = slide_object.slide
    title_shape = _find_existing_title_shape(slide)

    if title_shape is None:
        from pptx.util import Emu
        from pptx.enum.shapes import MSO_SHAPE
        title_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Pt(40), Pt(30), Pt(640), Pt(60),
        )

    set_no_fill(title_shape)
    tf = title_shape.text_frame
    tf.word_wrap = True
    tf.text = text

    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    for run in para.runs:
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = rgb(33, 45, 106)
    set_no_line(title_shape)

    title_bottom_y = title_shape.top + title_shape.height
    title_bottom_pt = title_bottom_y / 12700
    slide_object.last_bottom_y = max(slide_object.last_bottom_y, title_bottom_pt)
    slide_object.chart_start_y = slide_object.last_bottom_y + 20


def add_section_divider(slide_object: "SlideObject", text: str) -> None:
    """Render a large, centered title for section divider slides."""

    slide = slide_object.slide
    width = slide_object.slide_width - 80
    height = 160
    x = 40
    y = (slide_object.slide_height - height) / 3

    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Pt(x), Pt(y), Pt(width), Pt(height),
    )
    set_no_fill(shape)
    set_no_line(shape)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.text = text
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    for run in para.runs:
        run.font.size = Pt(40)
        run.font.bold = True
        run.font.color.rgb = rgb(33, 45, 106)

    slide_object.last_bottom_y = y + height
    slide_object.chart_start_y = slide_object.last_bottom_y + 20


def _remove_default_placeholders(slide) -> None:
    """Remove all placeholders except the title (idx 0) from a fresh slide."""
    sp_tree = slide.shapes._spTree
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx > 0:
            sp = ph._element
            sp_tree.remove(sp)
