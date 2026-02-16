import re
import aspose.slides as slides  # pyright: ignore[reportMissingModuleSource]
from aspose.pydrawing import Color  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]
from aspose.slides import FillType, NullableBool  # pyright: ignore[reportAttributeAccessIssue, reportMissingModuleSource]


def _add_portion(
    paragraph: slides.Paragraph,
    text: str,
    bold: bool,
    italic: bool,
    size: int,
) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    # Adds a formatted text portion to a paragraph.
    if not text:
        return
    portion = slides.Portion(text)
    paragraph.portions.add(portion)
    portion.portion_format.font_height = size
    portion.portion_format.font_bold = NullableBool.TRUE if bold else NullableBool.FALSE
    portion.portion_format.font_italic = NullableBool.TRUE if italic else NullableBool.FALSE
    portion.portion_format.fill_format.fill_type = FillType.SOLID
    portion.portion_format.fill_format.solid_fill_color.color = Color.black


def _add_inline_segments(paragraph: slides.Paragraph, text: str, size: int) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    # Split text by bold/italic markdown and render segments.
    parts = re.split(r"(\*\*.+?\*\*|\*.+?\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            _add_portion(paragraph, part[2:-2], True, False, size)
        elif part.startswith("*") and part.endswith("*"):
            _add_portion(paragraph, part[1:-1], False, True, size)
        else:
            _add_portion(paragraph, part, False, False, size)


def add_markdown_to_text_frame(
    tf: slides.TextFrame,
    markdown: str,
    body_size: int = 12,
    heading_size: int = 14,
    bullet_gap: int = 4,
    paragraph_gap: int = 8,
    bullet_symbol: str = "\u2022",
    align: str = "left",
    anchor: str = "top",
    line_spacing: float | None = None,
) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    # Render a limited markdown subset into a text frame.
    if hasattr(tf.paragraphs, "clear"):
        tf.paragraphs.clear()
    tf.text = ""
    if anchor == "top":
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.TOP
    elif anchor == "center":
        tf.text_frame_format.anchoring_type = slides.TextAnchorType.CENTER

    # Normalize input lines and trim leading/trailing blanks.
    lines = [ln.lstrip() for ln in markdown.splitlines()]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()

    first_para = None
    try:
        first_para = tf.paragraphs[0]
    except Exception:
        first_para = None
    if first_para is not None:
        if hasattr(first_para.portions, "clear"):
            first_para.portions.clear()
        else:
            try:
                while first_para.portions.count > 0:
                    first_para.portions.remove_at(0)
            except Exception:
                pass

    for index, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line:
            continue

        if index == 0 and first_para is not None:
            paragraph = first_para
        else:
            paragraph = slides.Paragraph()
            tf.paragraphs.add(paragraph)
        paragraph.paragraph_format.bullet.type = slides.BulletType.NONE
        paragraph.paragraph_format.space_before = 0
        if align == "center":
            paragraph.paragraph_format.alignment = slides.TextAlignment.CENTER
        else:
            paragraph.paragraph_format.alignment = slides.TextAlignment.LEFT
        if line_spacing is not None and hasattr(paragraph.paragraph_format, "line_spacing"):
            paragraph.paragraph_format.line_spacing = line_spacing

        # Simple heading.
        if line.startswith("### "):
            paragraph.paragraph_format.space_after = paragraph_gap
            _add_inline_segments(paragraph, line[4:].strip(), heading_size)
            continue

        # Bullet list item.
        if line.startswith("- "):
            paragraph.paragraph_format.space_after = bullet_gap
            if bullet_symbol:
                _add_portion(paragraph, f"{bullet_symbol} ", False, False, body_size)
            _add_inline_segments(paragraph, line[2:].strip(), body_size)
            continue

        paragraph.paragraph_format.space_after = paragraph_gap
        _add_inline_segments(paragraph, line, body_size)


def add_markdown_text(
    shape: slides.IShape,
    markdown: str,
    body_size: int = 12,
    heading_size: int = 14,
    bullet_gap: int = 4,
    paragraph_gap: int = 8,
    bullet_symbol: str = "\u2022",
    align: str = "left",
    anchor: str = "top",
    line_spacing: float | None = None,
) -> None:  # pyright: ignore[reportAttributeAccessIssue]
    # Convenience wrapper to render markdown into a shape.
    add_markdown_to_text_frame(
        shape.text_frame,
        markdown,
        body_size=body_size,
        heading_size=heading_size,
        bullet_gap=bullet_gap,
        paragraph_gap=paragraph_gap,
        bullet_symbol=bullet_symbol,
        align=align,
        anchor=anchor,
        line_spacing=line_spacing,
    )
