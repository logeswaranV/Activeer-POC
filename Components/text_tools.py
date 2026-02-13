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


def render_text_content(slide_object, component: dict) -> None:
    content = component.get("content", "")
    if not isinstance(content, str) or not content.strip():
        return

    width = slide_object.chart_width
    height = slide_object.get_chart_height()
    x, y = slide_object.get_next_chart_position(height)

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
