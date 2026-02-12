import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType, NullableBool  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides.util import SlideUtil  # pyright: ignore[reportMissingModuleSource]

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from main import SlideObject


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


def add_title(slide_object: "SlideObject", text: str) -> None:
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