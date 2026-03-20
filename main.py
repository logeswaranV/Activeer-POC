import json
import math
import copy
from pathlib import Path
from io import BytesIO

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

from Components.utils import (
    add_title,
    add_section_divider,
    _remove_default_placeholders,
    pt,
    set_no_fill,
    set_no_line,
    set_solid_fill,
)
from Components.chart_tools import add_graph
from Components.map_tools import render_map_image
from Components.text_tools import render_html_into_shape, render_meeting_info_markdown, render_list_into_shape
from Components.table_tools import render_table, render_meeting_info_table

CARD_PADDING = 12
INCH_TO_PT = 72
SHAPE_MAX_HEIGHT_IN = 7
SHAPE_MAX_HEIGHT = SHAPE_MAX_HEIGHT_IN * INCH_TO_PT
CARD_MAX_HEIGHT_IN = 5.2
CARD_MAX_HEIGHT = CARD_MAX_HEIGHT_IN * INCH_TO_PT
INPUT_JSON_PATH = Path('Input.json')
SAMPLE_TEMPLATE = Path('/Users/SIVAKAMI/Documents/Decera_project/Activeer-POC/SAMPLE.pptx')


def load_deck(path: Path = INPUT_JSON_PATH) -> dict:
    """Load the deck definition from Input.json."""
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}
    return data.get("deck", {})


class SlideObject:
    """State holder for a slide while we build it."""

    def __init__(
        self,
        slide,                  # pptx.slide.Slide
        slide_width: float,     # in points
        slide_height: float,    # in points
        chart_columns: int = 3,
        column_gap: float = 60,
        row_gap: float = 50,
        total_charts: int = 0,
        height_cap: float = CARD_MAX_HEIGHT,
    ):
        self.slide = slide
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
        self.height_cap = height_cap
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
            - CARD_PADDING * 1
        )
        per_row = available_height / rows if available_height > 0 else 120
        per_row = min(per_row, self.height_cap)
        return max(120, per_row)


# ── Slide dimension helper (returns pt from EMU) ───────────────────────────────

def _prs_slide_size_pt(prs: Presentation) -> tuple[float, float]:
    """Return (width_pt, height_pt) of the presentation slide size."""
    return prs.slide_width / 12700, prs.slide_height / 12700


# ── Template cloning via lxml ─────────────────────────────────────────────────

def _clone_slide_from_template(output_prs: Presentation, template_slide) -> object:
    """
    Clone a template slide into output_prs, copying both shape XML AND all
    slide relationships (images, media). Without relationships, PowerPoint
    Online rejects the file due to dangling rId references.
    """
def _clone_slide_from_template(output_prs: Presentation, template_slide) -> object:
    """Clone a slide from the template into the output presentation."""
    try:
        # 1. Add a slide using the MATCHING layout from the template
        layout = _get_layout_matching_template_slide(output_prs, template_slide)
        new_slide = output_prs.slides.add_slide(layout)
        
        new_sld = new_slide._element
        tmpl_sld = template_slide._element
        
        # 2. CLEAR the shapes that the layout injected (avoid duplicates)
        new_spTree = new_slide.shapes._spTree
        for child in list(new_spTree):
            if child.tag != qn("p:nvGrpSpPr"):
                new_spTree.remove(child)

        # 3. Copy shapes from template spTree
        tmpl_spTree = template_slide.shapes._spTree
        for child in tmpl_spTree:
            if child.tag != qn("p:nvGrpSpPr"):
                new_spTree.append(copy.deepcopy(child))

        # 4. Copy background override if it exists
        tmpl_bg = tmpl_sld.find(qn("p:bg"))
        if tmpl_bg is not None:
            existing_bg = new_sld.find(qn("p:bg"))
            if existing_bg is not None:
                new_sld.remove(existing_bg)
            new_sld.insert(0, copy.deepcopy(tmpl_bg))

        # 5. Map Relationships (Images, Tables, etc.)
        rId_map = {}
        for rel in template_slide.part.rels.values():
            # Skip architectural relationships
            if any(skip in rel.reltype for skip in ("slideLayout", "notesSlide", "slide")):
                continue
            try:
                if rel.is_external:
                    new_rId = new_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                else:
                    new_rId = new_slide.part.relate_to(rel.target_part, rel.reltype)
                rId_map[rel.rId] = new_rId
            except Exception:
                pass

        # 6. Update rId references in the XML in-place
        # Using .iter() is more robust across different element implementations
        for elem in new_sld.iter():
            for attr in [qn('r:id'), qn('r:embed'), qn('r:link')]:
                if attr in elem.attrib:
                    val = elem.attrib[attr]
                    if val in rId_map:
                        elem.attrib[attr] = rId_map[val]

        # 7. Finalize element sync
        new_sld.set("showMasterSp", "1")
        new_slide.part._element = new_sld
        
        # Invalidate internal shape cache so it reconstructs from the new XML
        if hasattr(new_slide, "_shapes"):
            delattr(new_slide, "_shapes")
            
        return new_slide

    except Exception as e:
        print(f"Error during slide cloning: {e}")
        return output_prs.slides.add_slide(output_prs.slide_layouts[0])

    return new_slide


def _get_layout_matching_template_slide(output_prs: Presentation, template_slide):
    """Find a layout in output_prs that matches the template slide's layout name."""
    target_name = template_slide.slide_layout.name
    for layout in output_prs.slide_layouts:
        if layout.name == target_name:
            return layout
    # Fallback: use first available layout
    return output_prs.slide_layouts[0]


def create_empty_output_presentation(template_prs: Presentation) -> Presentation:
    """
    Open the template and remove every existing slide, keeping masters/layouts.
    Slides must be removed by also dropping their OPC relationship — simply
    deleting from _sldIdLst leaves the XML parts in the ZIP, producing
    'Duplicate name' warnings and a corrupt output file.
    """
    output = Presentation(SAMPLE_TEMPLATE)

    sldIdLst = output.slides._sldIdLst
    # Collect all (rId, sldId) pairs first, then remove
    ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    entries = [
        (sld_id.get(f"{{{ns}}}id"), sld_id)
        for sld_id in list(sldIdLst)
    ]
    for rId, sld_id in entries:
        if rId:
            try:
                output.part.drop_rel(rId)   # removes the XML part from the package
            except Exception:
                pass
        sldIdLst.remove(sld_id)            # removes the <p:sldId> list entry

    return output


# ── Layout helpers ─────────────────────────────────────────────────────────────

def get_layout_by_name(prs: Presentation, layout_name: str):
    for layout in prs.slide_layouts:
        if layout.name == layout_name:
            return layout
    # Fallback to first layout if name not found
    print(f"⚠️  Layout '{layout_name}' not found, using default layout")
    return prs.slide_layouts[0]


def find_report_title_template(template_prs: Presentation, slide_name: str):
    matches = []
    for i, slide in enumerate(template_prs.slides):
        name = slide.slide_layout.name if slide.slide_layout else "Unknown"
        if name == slide_name:
            matches.append((i, slide))
    
    if matches:
        # Sort by shape count, same as before
        matches.sort(key=lambda s: len(s[1].shapes), reverse=True)
        idx, slide = matches[0]
        print(f"Found template '{slide_name}' at slide index {idx}")
        return slide
        
    raise ValueError(f"Template slide with layout '{slide_name}' not found")


# ── Shape helpers ──────────────────────────────────────────────────────────────

def get_shape_by_name(slide, name: str):
    """Find a shape by its name."""
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def remove_shape_by_name(slide, shape_name: str):
    sp_tree = slide.shapes._spTree
    for shape in list(slide.shapes):
        if shape.name == shape_name:
            sp = shape._element
            sp_tree.remove(sp)
            return


def rename_selection_panes(slide, slide_layout_name: str) -> None:
    """Rename shapes by matching their current names to semantic identifiers."""
    rename_map = {}

    if slide_layout_name == "Report Title":
        rename_map = {
            "Rectangle 7": "report_title_main",
            "Text Placeholder 13": "report_title_subtitle",
            "Rectangle 8": "report_title_speaker",
            "TextBox 2": "report_title_section",
            "Text Placeholder 74754": "report_title_bio",
        }
    elif slide_layout_name == "Meeting Information":
        rename_map = {
            "Group 9": "MI_Map",
            "Table 10": "MI_Attendee_Table",
            "TextBox 39": "MI_Attendees_Value",
            "TextBox 33": "MI_Moderator_Value",
            "TextBox 25": "MI_DateTime_Value",
            "TextBox 17": "MI_Location_Value",
            "TextBox 7": "MI_Event_Value",
        }

    if not rename_map:
        return

    for shape in slide.shapes:
        old_name = shape.name
        if old_name in rename_map:
            shape.name = rename_map[old_name]
            print(f"Renamed '{old_name}' → '{shape.name}'")


# ── Content replacement helpers ────────────────────────────────────────────────

def replace_content_by_selection_pane(slide, body: list) -> None:
    """Replace content on a slide using Selection Pane names."""
    for item in body:
        pane_name = item.get("selection_pane")
        component = item.get("component")
        content = item.get("content", "")

        if not pane_name:
            continue

        shape = get_shape_by_name(slide, pane_name)
        if not shape:
            print(f"⚠️ Shape '{pane_name}' not found in slide")
            continue

        if component == "text":
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            # Clear existing text
            for i in range(len(tf.paragraphs) - 1, 0, -1):
                p = tf.paragraphs[i]._p
                p.getparent().remove(p)
            for r in tf.paragraphs[0].runs:
                r._r.getparent().remove(r._r)
            run = tf.paragraphs[0].add_run()
            run.text = content


def replace_meeting_info_content(slide, body: list, prs: Presentation) -> None:
    slide_width_pt, slide_height_pt = _prs_slide_size_pt(prs)

    for item in body:
        pane = item.get("selection_pane")
        component = item.get("component")
        content = item.get("content")

        if not pane:
            continue

        shape = get_shape_by_name(slide, pane)

        if component == "meeting_info_table":
            if shape:
                x = shape.left / 12700
                y = shape.top / 12700
                w = shape.width / 12700
                h = shape.height / 12700
                remove_shape_by_name(slide, pane)
                render_meeting_info_table(
                    slide_object=SlideObject(slide, slide_width_pt, slide_height_pt),
                    component=item,
                    x=x, y=y, width=w, height=h,
                )
            continue

        if component == "map":
            if shape:
                x = shape.left / 12700
                y = shape.top / 12700
                w = shape.width / 12700
                h = shape.height / 12700
                remove_shape_by_name(slide, pane)
                map_bytes = render_map_image(content, int(w), int(h))
                slide.shapes.add_picture(
                    ensure_stream(map_bytes),
                    Pt(x),
                    Pt(y),
                    width=Pt(w),
                    height=Pt(h),
                )
            continue

        if component == "text" and shape and shape.has_text_frame:
            tf = shape.text_frame
            from pptx.enum.text import MSO_VERTICAL_ANCHOR
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

            orig_size = Pt(14)
            orig_color = RGBColor(0, 0, 0)
            orig_bold = False
            if tf.paragraphs and tf.paragraphs[0].runs:
                r = tf.paragraphs[0].runs[0]
                if r.font.size: orig_size = r.font.size
                if r.font.bold is not None: orig_bold = r.font.bold
                try:
                    if r.font.color and r.font.color.rgb:
                        orig_color = r.font.color.rgb
                except:
                    pass

            tf.text = content or ""
            if tf.paragraphs and tf.paragraphs[0].runs:
                run = tf.paragraphs[0].runs[0]
                run.font.size = orig_size
                run.font.color.rgb = orig_color
                run.font.bold = orig_bold
            
            print(f"Applied robust text to '{pane}' on slide {slide}: {content[:30]}...")


# ── Layout guide helper ────────────────────────────────────────────────────────

def _add_layout_guides(slide_object: SlideObject, columns: int) -> None:
    """Draw plain background guides that divide the available width into columns."""
    slide = slide_object.slide
    chart_height = slide_object.get_chart_height()
    total_gap = (columns - 1) * slide_object.column_gap
    col_width = (
        (slide_object.slide_width - slide_object.left_margin * 2 - total_gap)
        / max(1, columns)
    )
    y = slide_object.chart_start_y
    for idx in range(columns):
        x = slide_object.left_margin + idx * (col_width + slide_object.column_gap)
        guide = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Pt(x), Pt(y), Pt(col_width), Pt(chart_height),
        )
        set_no_fill(guide)
        set_no_line(guide)


def _all_charts(components: list) -> bool:
    if not components:
        return False
    return all(isinstance(c, dict) and c.get("component") == "chart" for c in components)

def ensure_stream(data):
    if isinstance(data, BytesIO):
        data.seek(0)
        return data
    return BytesIO(data)

# ── Component-in-slot renderer ─────────────────────────────────────────────────

def _render_component_in_slot(
    slide_object: SlideObject,
    component,
    x: float,
    y: float,
    width: float,
    height: float,
) -> None:
    if isinstance(component, list):
        items = [c for c in component if c]
        if not items:
            return
        gap = 12
        available_height = height - gap * (len(items) - 1)
        if len(items) == 2:
            heights = [available_height * 0.6, available_height * 0.4]
        else:
            per = available_height / len(items)
            heights = [per] * len(items)
        current_y = y
        for item, h in zip(items, heights):
            _render_component_in_slot(slide_object, item, x, current_y, width, h)
            current_y += h + gap
        return

    if not isinstance(component, dict):
        component = {"component": "text", "content": str(component)}

    slide = slide_object.slide
    comp_type = component.get("component")

    if comp_type == "chart":
        original_left = slide_object.left_margin
        original_chart_width = slide_object.chart_width
        original_chart_start_y = slide_object.chart_start_y
        slide_object.left_margin = x
        slide_object.chart_width = width
        slide_object.chart_start_y = y
        slide_object.current_column = 0
        slide_object.current_row = 0
        add_graph(slide_object, component, component.get("name", "Chart"))
        slide_object.left_margin = original_left
        slide_object.chart_width = original_chart_width
        slide_object.chart_start_y = original_chart_start_y

    elif comp_type == "map":
        map_bytes = render_map_image(
            component.get("content", []) or [], width=int(width), height=int(height)
        )
        slide.shapes.add_picture(
                    ensure_stream(map_bytes),
                    Pt(x),
                    Pt(y),
                    width=Pt(width),
                    height=Pt(height),
                )

    elif comp_type == "table":
        render_table(slide_object, component, x, y, width, height)

    elif comp_type == "meeting_info_table":
        render_meeting_info_table(slide_object, component, x, y, width, height)

    elif comp_type == "list":
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Pt(x), Pt(y), Pt(width), Pt(height),
        )
        set_no_fill(shape)
        set_no_line(shape)
        content = component.get("content", [])
        if isinstance(content, str):
            # Split by lines and remove bullet indicators if present
            lines = [ln.strip().lstrip("- ").strip() for ln in content.splitlines() if ln.strip()]
            render_list_into_shape(shape, lines)
        else:
            render_list_into_shape(shape, content)

    elif comp_type == "meeting_info_text":
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Pt(x), Pt(y), Pt(width), Pt(height),
        )
        set_no_fill(shape)
        set_no_line(shape)
        render_meeting_info_markdown(shape, component.get("content", ""))

    else:
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Pt(x), Pt(y), Pt(width), Pt(height),
        )
        set_no_fill(shape)
        set_no_line(shape)
        render_html_into_shape(shape, component.get("content", ""))


# ── Manual layout renderer ─────────────────────────────────────────────────────

def _render_manual_layout(
    prs: Presentation,
    slide,
    components: list,
    slide_width: float,
    slide_height: float,
    title: str,
    column_widths: list | None = None,
) -> None:
    slide_object = SlideObject(
        slide,
        slide_width,
        slide_height,
        chart_columns=len(components),
        column_gap=35,
        row_gap=35,
        total_charts=len(components),
        height_cap=SHAPE_MAX_HEIGHT,
    )
    if title:
        add_title(slide_object, title)

    chart_height = slide_object.get_chart_height()
    total_gap = slide_object.column_gap * (len(components) - 1)

    if isinstance(column_widths, list) and column_widths and all(
        isinstance(v, (int, float)) and v > 0 for v in column_widths
    ):
        ratios = column_widths[: len(components)]
        if len(ratios) < len(components):
            ratios += [1.0] * (len(components) - len(ratios))
        total = sum(ratios) or len(components)
        widths = [
            (slide_object.slide_width - slide_object.left_margin * 2 - total_gap)
            * r / total
            for r in ratios
        ]
    else:
        widths = [
            (slide_object.slide_width - slide_object.left_margin * 2 - total_gap)
            / max(1, len(components))
        ] * len(components)

    base_y = slide_object.chart_start_y
    for idx, component in enumerate(components):
        col_width = widths[idx]
        x = slide_object.left_margin + sum(widths[:idx]) + slide_object.column_gap * idx
        _render_component_in_slot(slide_object, component, x, base_y, col_width, chart_height)

    slide_object.last_bottom_y = base_y + chart_height


# ── Main slide builder ─────────────────────────────────────────────────────────

def create_slide(
    presentation: Presentation,
    deck_payload: dict,
    template_pres: Presentation,
) -> None:
    slide_data = sorted(
        deck_payload["slides"],
        key=lambda s: s.get("order", 0),
    )

    slide_width_pt, slide_height_pt = _prs_slide_size_pt(presentation)

    for slide_payload in slide_data:
        slide_type = slide_payload["slide_type"]
        layout_name = slide_payload["layout"]

        # ── REPORT TITLE (clone from template) ──────────────────────────────
        if slide_type == "report_title":
            report_title_template = find_report_title_template(template_pres, "Report Title")
            slide = _clone_slide_from_template(presentation, report_title_template)
            rename_selection_panes(slide, "Report Title")
            replace_content_by_selection_pane(slide, slide_payload.get("body", []))
            continue

        # ── MEETING INFORMATION (clone from template) ────────────────────────
        if slide_type == "meeting_info":
            mi_template = find_report_title_template(template_pres, "Meeting Information")
            slide = _clone_slide_from_template(presentation, mi_template)
            rename_selection_panes(slide, "Meeting Information")
            replace_meeting_info_content(slide, slide_payload.get("body", []), presentation)
            continue

        # ── TITLE ONLY (clone from template) ────────────────────────
        if slide_type == "section_divider":
            to_template = find_report_title_template(template_pres, "Section Divider A")
            slide = _clone_slide_from_template(presentation, to_template)
            # Find the title placeholder or any text frame to populate
            title = slide_payload.get("title", "")
            populated = False
            to_remove = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    tf = shape.text_frame
                    if not populated:
                        tf.text = title
                        for para in tf.paragraphs:
                            for run in para.runs:
                                run.font.bold = True
                                run.font.size = Pt(40)
                                run.font.color.rgb = RGBColor(33, 45, 106)
                        populated = True
                    else:
                        # Mark for removal to avoid "Click to add text" prompt
                        to_remove.append(shape)
            
            for shape in to_remove:
                sp = shape._element
                sp.getparent().remove(sp)
            if not populated:
                # Fallback if no text frame found in template
                add_section_divider(SlideObject(slide, slide_width_pt, slide_height_pt), title)
            continue

        # ── ALL OTHER SLIDES (empty from layout) ────────────────────────────
        layout = get_layout_by_name(presentation, layout_name)
        slide = presentation.slides.add_slide(layout)
        _remove_default_placeholders(slide)

        components = slide_payload.get("body") or []
        chart_only = _all_charts(components)

        if chart_only:
            slide_object = SlideObject(
                slide,
                slide_width_pt,
                slide_height_pt,
                chart_columns=len(components),
                total_charts=len(components),
                height_cap=CARD_MAX_HEIGHT,
            )
            if slide_payload.get("title"):
                add_title(slide_object, slide_payload["title"])

            for component in components:
                add_graph(slide_object, component, component.get("name"))
        else:
            _render_manual_layout(
                presentation,
                slide,
                components,
                slide_width_pt,
                slide_height_pt,
                slide_payload.get("title", ""),
                slide_payload.get("column_widths"),
            )


# ── Entry point ────────────────────────────────────────────────────────────────

deck_definition = load_deck()

template_pres = Presentation(str(SAMPLE_TEMPLATE))
presentation = create_empty_output_presentation(template_pres)

create_slide(presentation, deck_definition, template_pres)

presentation.save("NewPresentation.pptx")
print("✅  Saved NewPresentation.pptx")