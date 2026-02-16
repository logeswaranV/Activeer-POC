# Activeer-POC

## Flow Overview
This repo generates a PowerPoint deck from `Input.json` using Aspose.Slides.

1. `main.py` loads `Input.json` with `load_deck()`, sorts slides by `order`, and creates each slide.
2. For each slide, it routes the content by body type:
   - `body` items with `type: "table"` or `type: "markdown"` use the composite renderer.
   - Otherwise, it renders chart components and the legacy two-column layout.
3. The output is saved to `NewPresentation.pptx`.

## Composite Two-Column (Table + Markdown)
The composite slide you added in `Input.json` is rendered by these pieces:

- `main.py` calls `add_two_column_components()` when `body` contains `table`/`markdown` items.
- `Components/composite_tools.py` splits components into left/right stacks and renders them.
- Table rendering is handled by `Components/table_tools.py` via `add_table()`.
  - It also uses `Components/markdown_tools.py` (`add_markdown_to_text_frame`) to format the table body cell text.
- Markdown rendering for the right column uses `Components/markdown_tools.py` (`add_markdown_text`).

### Input.json shape for this slide
The composite slide expects:
- `body`: a list of objects with:
  - `type: "table"` or `type: "markdown"`
  - `position: "left"` or `position: "right"`
  - `headers`/`rows` for tables, or `content` for markdown

## Composite Slide Fields (What + Why)
These are the fields used by the composite slide flow and what they do:

- `order`: Sort key so slides render in a stable sequence.
- `slide_type`: Informational; not currently used by the renderer.
- `title`: Slide title text rendered at the top.
- `body`: List of renderable items for left/right columns.

For each component in `body`:

- `type`: `"table"` or `"markdown"` to choose the rendering function.
- `position`: `"left"` or `"right"` to choose the column stack.
- `name`: Optional label for the component (not used by the renderer today).

Table component fields:

- `headers`: Array of header strings (1+ columns). Each header maps to a column.
- `rows`: Array of rows, each containing 1+ cell values. Column count is dynamic and based on the longest row.

Markdown component fields:

- `content`: String of markdown to render in the right column text box.
