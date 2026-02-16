import re

import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType  # pyright: ignore[reportAttributeAccessIssue]


def render_html_into_shape(shape: slides.IShape, html: str) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    tf = shape.text_frame
    tf.text = ""
    tf.text_frame_format.wrap_text = slides.NullableBool.TRUE
    tf.text_frame_format.margin_left = 8
    tf.text_frame_format.margin_right = 8
    tf.text_frame_format.margin_top = 6
    tf.text_frame_format.margin_bottom = 6
    tf.paragraphs.clear()
    tf.paragraphs.add_from_html(html)

    # normalize font styling
    for para in tf.paragraphs:
        para.paragraph_format.alignment = slides.TextAlignment.LEFT
        for portion in para.portions:
            portion.portion_format.font_height = portion.portion_format.font_height or 16
            portion.portion_format.font_bold = portion.portion_format.font_bold
            portion.portion_format.fill_format.fill_type = FillType.SOLID
            portion.portion_format.fill_format.solid_fill_color.color = Color.from_argb(255, 45, 45, 45)


def render_meeting_info_markdown(shape: slides.IShape, markdown: str) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    """Render meeting-info markdown as bullets with custom sizing."""

    lines = [ln.strip() for ln in markdown.splitlines() if ln.strip()]
    # if we cannot access slide, fall back to rendering inside the shape
    slide = getattr(shape, "slide", None)
    if not slide:
        shape.fill_format.fill_type = FillType.NO_FILL
        shape.line_format.fill_format.fill_type = FillType.NO_FILL
        tf = shape.text_frame
        tf.text = ""
        tf.paragraphs.clear()
        tf.text_frame_format.wrap_text = slides.NullableBool.TRUE
        tf.text_frame_format.margin_left = 10
        tf.text_frame_format.margin_right = 10
        tf.text_frame_format.margin_top = 10
        tf.text_frame_format.margin_bottom = 10
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER

        for line in lines:
            if line.startswith("- "):
                line = line[2:].strip()
            para = slides.Paragraph()
            tf.paragraphs.add(para)
            para.paragraph_format.alignment = slides.TextAlignment.LEFT
            para.portions.clear()

            parts = re.split(r"(\*\*.*?\*\*)", line)
            for part in parts:
                if not part:
                    continue
                is_bold = part.startswith("**") and part.endswith("**") and len(part) >= 4
                text = part[2:-2] if is_bold else part
                portion = slides.Portion(text)
                pf = portion.portion_format
                pf.font_height = 12 if is_bold else 18
                pf.font_bold = slides.NullableBool.TRUE if is_bold else slides.NullableBool.FALSE
                pf.fill_format.fill_type = FillType.SOLID
                pf.fill_format.solid_fill_color.color = Color.black
                para.portions.add(portion)
        return

    # transparent base shape
    shape.fill_format.fill_type = FillType.NO_FILL
    shape.line_format.fill_format.fill_type = FillType.NO_FILL

    # layout per item as separate rectangles
    gap = 8
    x = shape.x
    y = shape.y
    width = shape.width
    height = shape.height
    n = max(1, len(lines))
    per_h = max(30.0, (height - gap * (n - 1)) / n)
    per_h = min(per_h, 0.75 * 72)  # cap at 0.75 inches

    for line in lines:
        rect = slide.shapes.add_auto_shape(
            slides.ShapeType.RECTANGLE,
            x,
            y,
            width,
            per_h,
        )
        rect.fill_format.fill_type = FillType.SOLID
        rect.fill_format.solid_fill_color.color = Color.from_argb(64, 201, 203, 224)  # c9cbe0 alpha=25
        rect.line_format.fill_format.fill_type = FillType.NO_FILL

        tf = rect.text_frame
        tf.text = ""
        tf.paragraphs.clear()
        tf.text_frame_format.wrap_text = slides.NullableBool.TRUE
        tf.text_frame_format.margin_left = 10
        tf.text_frame_format.margin_right = 10
        tf.text_frame_format.margin_top = 8
        tf.text_frame_format.margin_bottom = 8
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER

        if line.startswith("- "):
            line = line[2:].strip()

        para = slides.Paragraph()
        tf.paragraphs.add(para)
        para.paragraph_format.alignment = slides.TextAlignment.LEFT
        para.portions.clear()

        parts = re.split(r"(\*\*.*?\*\*)", line)
        for part in parts:
            if not part:
                continue
            is_bold = part.startswith("**") and part.endswith("**") and len(part) >= 4
            text = part[2:-2] if is_bold else part
            portion = slides.Portion(text)
            pf = portion.portion_format
            pf.font_height = 12 if is_bold else 18
            pf.font_bold = slides.NullableBool.TRUE if is_bold else slides.NullableBool.FALSE
            pf.fill_format.fill_type = FillType.SOLID
            pf.fill_format.solid_fill_color.color = Color.black
            para.portions.add(portion)

        y += per_h + gap

def render_list_into_shape(shape: slides.IShape, items, level: int = 0) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    """Render a (possibly nested) Python list as bullet points inside the shape."""
    tf = shape.text_frame
    tf.text = ""
    tf.paragraphs.clear()
    tf.text_frame_format.wrap_text = slides.NullableBool.TRUE
    tf.text_frame_format.margin_left = 10
    tf.text_frame_format.margin_right = 10
    tf.text_frame_format.margin_top = 8
    tf.text_frame_format.margin_bottom = 8
    tf.text_frame_format.anchoring_type = slides.TextAnchorType.TOP

    def _add_item(item, lvl: int) -> None:
        if isinstance(item, list):
            for sub in item:
                _add_item(sub, min(lvl + 1, 4))
            return
        text = str(item)
        para = slides.Paragraph()
        para.paragraph_format.alignment = slides.TextAlignment.LEFT
        para.paragraph_format.depth = lvl
        para.paragraph_format.bullet.type = slides.BulletType.SYMBOL
        para.paragraph_format.bullet.char = "\u2022"
        para.paragraph_format.bullet.color.color = Color.black
        para.portions.clear()
        portion = slides.Portion(text)
        pf = portion.portion_format
        pf.font_height = 16 if lvl == 0 else 14
        pf.font_bold = slides.NullableBool.FALSE
        pf.fill_format.fill_type = FillType.SOLID
        pf.fill_format.solid_fill_color.color = Color.black
        para.portions.add(portion)
        tf.paragraphs.add(para)

    if isinstance(items, list):
        for it in items:
            _add_item(it, level)
    else:
        _add_item(items, level)
