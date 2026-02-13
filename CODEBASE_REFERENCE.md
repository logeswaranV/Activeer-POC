# Activeer-POC Codebase Reference

## Table of Contents
- [Architecture Overview](#architecture-overview)
  - [Execution Flow](#execution-flow)
  - [Layout Strategy & Naming Conventions](#layout-strategy--naming-conventions)
  - [Input Schema & Data Flow](#input-schema--data-flow)
- [Module Breakdown](#module-breakdown)
  - [main.py (entry & orchestrator)](#mainpy-entry--orchestrator)
  - [Components/utils.py](#componentsutilspy)
  - [Components/chart_tools.py](#componentschart_toolspy)
  - [Components/meeting_info_tools.py](#componentsmeeting_info_toolspy)
  - [Input.json sample deck](#inputjson-sample-deck)
- [Configuration & Dependencies](#configuration--dependencies)
- [Build, Run, Test & Deploy](#build-run-test--deploy)
- [Onboarding Checklist](#onboarding-checklist)
- [Troubleshooting Tips](#troubleshooting-tips)
- [Sources of Truth](#sources-of-truth)
- [Open Questions & Next Steps](#open-questions--next-steps)

## Architecture Overview

### Execution Flow
- `main.py` loads `Input.json`, instantiates an Aspose `Presentation`, and delegates slide creation through `create_slide` before saving `NewPresentation.pptx` (see `main.py:17-135`).
- `create_slide` normalizes the slide size to widescreen, removes layout placeholders, initializes a `SlideObject`, and routes `meetingInfo` versus chart components so layout, titles, and visual rendering stay isolated (`main.py:85-135`).
- Chart rendering happens off the Aspose canvas: Matplotlib generates PNG slices from aggregation data and Aspose wraps each PNG in a card-shaped picture frame (`Components/chart_tools.py:62-294`).

### Layout Strategy & Naming Conventions
- `SlideObject` encapsulates per-slide state (current column/row, margins, chart width, and adaptive row height) so the rest of the code merely asks for positions (`main.py:29-83`).
- Card padding, column gaps, and row gaps are defined as constants (e.g., `CARD_PADDING`, `INCH_TO_PT`, `row_gap`) to keep layout tunable without digging into Aspose calls (`main.py:10-14` and `Components/chart_tools.py:13-114`).
- Titles use `add_title` to ensure rectangles remain fill-free and bold, reusing placeholder detection for existing titles to avoid duplicate shapes (`Components/utils.py:12-60`).
- Naming sticks to snake_case files and PascalCase folders per `Coding_Style.md:1-21`, plus docstrings before functions.

### Input Schema & Data Flow
- The deck definition in `Input.json:1-74` contains `metadata` (title, version, theme) plus a `slides` array. Each slide can be `slide_type: "meetingInfo"` or `content` with chart components.
- Each chart entry must be `component: "chart"`, a `chartType` (e.g., `horizontal_bar_chart`, `donut_chart`), optional labels, and an `aggregations` map of buckets to counts (`Input.json:25-70`).
- `load_deck` quietly returns `{}` if the file is missing or malformed, making the generator a no-op rather than crashing (`main.py:17-27`).
- Meeting info slides reuse markdown-like strings; `_parse_columns` strips bullet markers, bold syntax, and inline links before rendering (`Components/meeting_info_tools.py:8-106`).

## Module Breakdown

### main.py (entry & orchestrator)
- Responsibilities: parse `Input.json`, instantiate Aspose `Presentation`, manage layout state, distinguish slide types, and trigger chart or meeting info rendering before saving to `NewPresentation.pptx` (`main.py:17-135`).
- Key internals:
  - `load_deck` handles JSON loading/fallbacks (`main.py:17-27`).
  - `SlideObject` stores layout-related constants and computes positions/heights so chart rendering remains oblivious to Aspose grid details (`main.py:29-83`).
  - `create_slide` sets slide size, removes placeholders, and dispatches to component helpers (`main.py:85-135`).
- Code snippet (critical layout helper):
```python
    def get_next_chart_position(self, chart_height: float) -> tuple[float, float]:
        if self.current_column >= self.chart_columns:
            self.current_column = 0
            self.current_row += 1
        x = self.left_margin + (self.chart_width + self.column_gap) * self.current_column
        y = self.chart_start_y + self.current_row * (chart_height + self.row_gap)
        self.current_column += 1
        self.last_bottom_y = y + chart_height
        return x, y
```
This snippet (`main.py:61-69`) shows how SlideObject staggers cards in rows and columns so Matplotlib cards never overlap.

### Components/utils.py
- Purpose: reusable helpers for Aspose slide shaping and placeholder cleanup (`Components/utils.py:1-71`).
- `_find_existing_title_shape` searches placeholders so the title text can be reused rather than recreating shapes, preserving formatting anchors (`Components/utils.py:12-21`).
- `add_title` ensures the text is bold, navy, left-aligned, and that the slide state updates `last_bottom_y` so charts start below titles (`Components/utils.py:34-60`).
- `_remove_default_placeholders` clears body/subtitle placeholders so slides start from a blank canvas (`Components/utils.py:62-71`).

### Components/chart_tools.py
- Converts aggregation payloads into Aspose-backed cards by generating Matplotlib figures and wrapping them as PNG images inside Aspose picture frames (`Components/chart_tools.py:62-294`).
- `_add_card_background` draws rounded rectangles, applies solid fills/borders, and adds a drop shadow for depth (`Components/chart_tools.py:36-60`).
- `_render_chart_image` branches on `chartType`; donut charts build legend handles and stylized pies, while horizontal bar charts apply consistent typography, labels, and values with dynamic x-axis limits (`Components/chart_tools.py:138-289`).
- Width/height scaling constants (`WIDTH_SCALE`, `HEIGHT_SCALE`, `DONUT_*`) keep the PNG size proportional to the card space and allow tweaking the rendered DPI without changing Aspose code (`Components/chart_tools.py:13-136`).

### Input.json sample deck
 - Shows metadata plus a single chart slide whose `body` array contains chart components (`Input.json:1-74`).
 - Chart slides contain `component: "chart"`, optional label overrides, and aggregation maps that become Matplotlib data (`Input.json:25-70`).

## Configuration & Dependencies

| Concern | Location | Notes |
| --- | --- | --- |
| Deck definition | `Input.json:1-74` | Update slide metadata, body arrays, component fields, and aggregation maps to influence the output. |
| Layout knobs | `main.py:10-83`, `Components/chart_tools.py:13-136` | Constants like `CARD_PADDING`, `chart_columns`, `column_gap`, `row_gap`, and `DONUT_*` control spacing and DPI scaling. |
| Dependencies | `requirements.txt:1-2` | `aspose-slides` + `matplotlib`; Aspose must be licensed/available at runtime, and Matplotlib renders charts. |
| Coding style reminders | `Coding_Style.md:1-21` | Use PascalCase for folders, snake_case for files, docstrings before functions, and avoid nested logic. |

## Build, Run, Test & Deploy
- Ensure a Python interpreter (3.11+ recommended to match Aspose compatibility) and pip are available.
- Create/activate a virtual environment (the repo already has `.venv/`, but recreate if needed) and install dependencies:
  ```sh
  python -m venv .venv
  .\\.venv\\Scripts\\activate
  pip install -r requirements.txt
  ```
- Run the generator:
  ```sh
  python main.py
  ```
  This reads `Input.json`, generates slides via Aspose, and writes `NewPresentation.pptx` next to the script (`main.py:133-136`).
- There are no automated tests or CI scripts yet, so manual verification (opening `NewPresentation.pptx`) is required after each change.
- Deploying currently means handing over the generated PPTX; there is no packaging script beyond Aspose's save call.

## Onboarding Checklist
1. Clone the repo and inspect `Coding_Style.md` for naming rules before coding.
2. Set up/activate a virtual environment and install requirements (`requirements.txt`).
3. Confirm Aspose Slides licensing/installation to avoid runtime errors (`requirements.txt`).
4. Open `Input.json` to understand existing slides; update or expand the `deck.slides` array as needed.
5. Run `python main.py`, then open `NewPresentation.pptx` in PowerPoint to verify layout.
6. Track layout metadata (e.g., `SlideObject.chart_columns`, `column_gap`, `row_gap`) when adding new chart types/components.
7. Update decks, components, and documentation here whenever a new pattern or configuration option is introduced.

## Troubleshooting Tips
- Aspose initialization errors usually mean a missing/expired license or incorrect runtime; ensure `ASPOSE_LICENSE` environment variables (if any) are set and Aspose is installed (`requirements.txt`).
- If `NewPresentation.pptx` shows overlapping cards, inspect `SlideObject` constants and `chart_width` calculation to ensure gaps/margins match the number of charts (`main.py:29-83`).
- Missing slides or charts often stem from empty `body` arrays or components missing the `component: "chart"` flag, so validate `Input.json` carefully (`Input.json:25-70`).
- When meeting info text renders with markdown artifacts, chips in `_clean_markdown` strip bold/links, but additional characters (like `Â`) are hardcoded replacements (`Components/meeting_info_tools.py:8-28`). Extend `_clean_markdown` if new encoding quirks appear.
- Font/color inconsistencies in Matplotlib charts can be adjusted via `x_label_style`, `y_label_style`, and `BAR_CHART_COLOR` within `_render_chart_image` (`Components/chart_tools.py:236-288`).

## Sources of Truth
| Area | File(s) | Why it matters |
| --- | --- | --- |
| Entry & layout orchestration | `main.py:17-135` | Controls JSON loading, slide creation loops, and the SlideObject layout toolkit. |
| Title + cleanup utilities | `Components/utils.py:12-71` | Centralizes placeholder massage and ensures new slides align visually. |
| Chart rendering | `Components/chart_tools.py:62-294` | Documents how aggregation data becomes Aspose cards via Matplotlib imaging. |
| Meeting info rendering | `Components/meeting_info_tools.py:8-106` | Explains markdown cleanup and multi-column layout for textual slides. |
| Deck configuration | `Input.json:1-74` | Shows the exact JSON shape the system consumes today. |

## Open Questions & Next Steps
1. The repo lacks automated tests or CI; adding smoke tests (e.g., verifying `load_deck` parsing or comparing slide counts) would improve confidence. (Unknown coverage.)
2. Aspose licensing/setup instructions are absent—document how to obtain/point to the license file before runtime errors occur. (Need confirmation from the team.)
3. No logging/metrics exist; consider adding debug logs around `add_graph` and `render_meeting_info` to trace failing slide payloads when inputs grow. (Future step.)
4. The generator currently only supports `chart` components and `meetingInfo` slides. Additional component types (tables, images) would need a new extensibility strategy. (Need product requirements.)
5. `NewPresentation.pptx` is overwritten with each run; if multiple versions are needed, add a timestamp/parameterized path. (Design opportunity.)
6. `.venv/` and `__pycache__/` are ignored per `.gitignore` but were not examined because they are runtime artifacts—no additional source files live there.
