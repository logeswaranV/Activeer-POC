import re
from html.parser import HTMLParser

from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

from Components.utils import pt, rgb, set_no_fill, set_solid_fill, set_no_line


# ── Minimal HTML → run parser ─────────────────────────────────────────────────

class _Span:
    """A flat span of text with styling attributes."""
    __slots__ = ("text", "bold", "italic", "font_size", "color")

    def __init__(self, text, bold=False, italic=False, font_size=None, color=None):
        self.text = text
        self.bold = bold
        self.italic = italic
        self.font_size = font_size  # int in pt, or None
        self.color = color          # RGBColor or None


class _HTMLToSpans(HTMLParser):
    """Convert a subset of HTML to a list-of-list structure: [[Span,...], ...]
    where each inner list is a paragraph (split on <p>/<br>/<li>/<h*>).
    """

    def __init__(self):
        super().__init__()
        self._bold = 0
        self._italic = 0
        self._font_size = None
        self._color = None
        self.paragraphs: list[list[_Span]] = [[]]

    # helpers
    def _cur_para(self):
        return self.paragraphs[-1]

    def _new_para(self):
        if self.paragraphs[-1]:  # don't add empty trailing paras redundantly
            self.paragraphs.append([])

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        attr_dict = dict(attrs)
        if tag in ("b", "strong"):
            self._bold += 1
        elif tag in ("i", "em"):
            self._italic += 1
        elif tag in ("p", "li", "h1", "h2", "h3", "h4", "h5", "h6"):
            self._new_para()
        elif tag == "br":
            self._new_para()
        elif tag == "span":
            style = attr_dict.get("style", "")
            fs_match = re.search(r"font-size\s*:\s*([\d.]+)pt", style)
            if fs_match:
                self._font_size = int(float(fs_match.group(1)))
            color_match = re.search(r"color\s*:\s*#([0-9a-fA-F]{6})", style)
            if color_match:
                h = color_match.group(1)
                self._color = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("b", "strong"):
            self._bold = max(0, self._bold - 1)
        elif tag in ("i", "em"):
            self._italic = max(0, self._italic - 1)
        elif tag in ("p", "li", "h1", "h2", "h3", "h4", "h5", "h6"):
            self._new_para()
        elif tag == "span":
            self._font_size = None
            self._color = None

    def handle_data(self, data):
        text = data.replace("\xa0", " ")
        if not text:
            return
        span = _Span(
            text=text,
            bold=self._bold > 0,
            italic=self._italic > 0,
            font_size=self._font_size,
            color=self._color,
        )
        self._cur_para().append(span)


def _html_to_paragraphs(html: str) -> list[list[_Span]]:
    parser = _HTMLToSpans()
    parser.feed(html)
    return [p for p in parser.paragraphs if p]


# ── Public helpers ────────────────────────────────────────────────────────────

def render_html_into_shape(shape, html: str) -> None:
    """Render a subset of HTML text into the shape's text frame."""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(8)
    tf.margin_right = Pt(8)
    tf.margin_top = Pt(6)
    tf.margin_bottom = Pt(6)

    paragraphs = _html_to_paragraphs(html)
    if not paragraphs:
        tf.text = html  # fallback: dump raw text
        return

    # Clear existing paragraphs (keep at least one)
    for i in range(len(tf.paragraphs) - 1, 0, -1):
        p = tf.paragraphs[i]._p
        p.getparent().remove(p)

    first = True
    for spans in paragraphs:
        if first:
            para = tf.paragraphs[0]
            # clear existing runs
            for r in para.runs:
                r._r.getparent().remove(r._r)
            first = False
        else:
            para = tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        for span in spans:
            run = para.add_run()
            run.text = span.text
            run.font.bold = span.bold
            run.font.italic = span.italic
            if span.font_size:
                run.font.size = Pt(span.font_size)
            if span.color:
                run.font.color.rgb = span.color


def render_meeting_info_markdown(shape, markdown: str) -> None:
    """Render meeting-info markdown as bullets with custom sizing."""

    lines = [ln.strip() for ln in markdown.splitlines() if ln.strip()]
    slide = getattr(shape, "slide", None)

    if not slide:
        # fallback: render inside the shape itself
        set_no_fill(shape)
        set_no_line(shape)
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Pt(10)
        tf.margin_right = Pt(10)
        tf.margin_top = Pt(10)
        tf.margin_bottom = Pt(10)

        # clear
        for i in range(len(tf.paragraphs) - 1, 0, -1):
            p = tf.paragraphs[i]._p
            p.getparent().remove(p)

        first = True
        for line in lines:
            if line.startswith("- "):
                line = line[2:].strip()
            if first:
                para = tf.paragraphs[0]
                for r in para.runs:
                    r._r.getparent().remove(r._r)
                first = False
            else:
                para = tf.add_paragraph()
            para.alignment = PP_ALIGN.LEFT

            parts = re.split(r"(\*\*.*?\*\*)", line)
            for part in parts:
                if not part:
                    continue
                is_bold = part.startswith("**") and part.endswith("**") and len(part) >= 4
                text = part[2:-2] if is_bold else part
                run = para.add_run()
                run.text = text
                run.font.size = Pt(12) if is_bold else Pt(18)
                run.font.bold = is_bold
                run.font.color.rgb = RGBColor(0, 0, 0)
        return

    # Layout per item as separate rectangles on the slide
    set_no_fill(shape)
    set_no_line(shape)

    gap = 8
    x = shape.left / 12700
    y = shape.top / 12700
    width = shape.width / 12700
    height = shape.height / 12700
    n = max(1, len(lines))
    per_h = max(30.0, (height - gap * (n - 1)) / n)
    per_h = min(per_h, 0.75 * 72)

    from pptx.enum.shapes import MSO_SHAPE
    for line in lines:
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Pt(x), Pt(y), Pt(width), Pt(per_h),
        )
        set_solid_fill(rect, RGBColor(201, 203, 224))
        rect.fill.fore_color.theme_color  # side-effect: makes fill solid
        # set 25% transparency via XML
        try:
            fgClr = rect.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
            alpha_elem = etree.SubElement(fgClr, qn("a:alpha"))
            alpha_elem.set("val", "25000")  # 25% = 25000 out of 100000
        except Exception:
            pass

        set_no_line(rect)

        tf = rect.text_frame
        tf.word_wrap = True
        tf.margin_left = Pt(10)
        tf.margin_right = Pt(10)
        tf.margin_top = Pt(8)
        tf.margin_bottom = Pt(8)

        # clear default content
        for i in range(len(tf.paragraphs) - 1, 0, -1):
            p = tf.paragraphs[i]._p
            p.getparent().remove(p)

        cur_line = line[2:].strip() if line.startswith("- ") else line

        para = tf.paragraphs[0]
        for r in para.runs:
            r._r.getparent().remove(r._r)
        para.alignment = PP_ALIGN.LEFT

        parts = re.split(r"(\*\*.*?\*\*)", cur_line)
        for part in parts:
            if not part:
                continue
            is_bold = part.startswith("**") and part.endswith("**") and len(part) >= 4
            text = part[2:-2] if is_bold else part
            run = para.add_run()
            run.text = text
            run.font.size = Pt(12) if is_bold else Pt(18)
            run.font.bold = is_bold
            run.font.color.rgb = RGBColor(0, 0, 0)

        y += per_h + gap


def render_list_into_shape(shape, items, level: int = 0) -> None:
    """Render a (possibly nested) Python list as bullet points inside the shape."""
    from pptx.oxml.ns import qn
    from lxml import etree

    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(10)
    tf.margin_right = Pt(10)
    tf.margin_top = Pt(8)
    tf.margin_bottom = Pt(8)

    # Clear
    for i in range(len(tf.paragraphs) - 1, 0, -1):
        p = tf.paragraphs[i]._p
        p.getparent().remove(p)
    # Remove runs from first paragraph
    for r in tf.paragraphs[0].runs:
        r._r.getparent().remove(r._r)

    first_para_used = False

    def _add_item(item, lvl: int) -> None:
        nonlocal first_para_used
        if isinstance(item, list):
            for sub in item:
                _add_item(sub, min(lvl + 1, 4))
            return
        text = str(item)

        if not first_para_used:
            para = tf.paragraphs[0]
            first_para_used = True
        else:
            para = tf.add_paragraph()

        para.alignment = PP_ALIGN.LEFT
        para.level = lvl

        # Set bullet via XML (<a:buChar char="•">)
        pPr = para._pPr
        if pPr is None:
            pPr = para._p.get_or_add_pPr()
        # Remove existing buNone/buChar/buAutoNum
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buClrTx"):
            elem = pPr.find(qn(tag))
            if elem is not None:
                pPr.remove(elem)
        buChar = etree.SubElement(pPr, qn("a:buChar"))
        buChar.set("char", "■")

        run = para.add_run()
        run.text = text
        run.font.size = Pt(14)
        run.font.bold = False
        run.font.color.rgb = RGBColor(64, 64, 64)

    if isinstance(items, list):
        for it in items:
            _add_item(it, level)
    else:
        _add_item(items, level)
