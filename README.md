# html_rich_text_creator

A small in-browser rich text editor focused on creating question papers and TOC-driven documents.

## Quick Start

- Open `index.html` in a modern browser (Chrome, Edge, Safari). No build step required — files are static.
- The editor UI contains a left **Table of Contents** sidebar and the main editor area with a toolbar.

## Features

- Templates: load predefined templates from the toolbar `Templates` dropdown. Available templates include:
	- `question-paper` — standard question paper
	- `question-paper-toc` — TOC layout with 4-column options
	- `question-paper-toc-2col-options` — TOC with 2-column options
	- `question-paper-toc-1col-options` — TOC with single-column options
	- `question-paper-toc-2col-questions-1col-options` — 2-column question layout with 1-column options
	- `question-paper-toc-3col-questions-1col-options` — 3-column question layout with 1-column options
	- `answer-key`, `worksheet`, `lesson-plan`, `blank`

- Table of Contents (TOC): automatically generated from `h1`–`h4` headings in the editor.
	- The TOC shows hierarchical numbering (e.g., `1`, `1.1`, `1.1.1`).
	- Images and tables that belong to a heading are shown as previews inside the TOC.
	- Click a TOC entry to scroll the editor to the corresponding heading.
	- Export TOC as JSON using the `JSON` button in the TOC header.

- Resizable Sidebar: drag the narrow gutter between the TOC and the editor to resize the left sidebar (min 220px, max 600px).

- Images:
	- Paste from clipboard or upload via the image buttons in the toolbar.
	- Images are wrapped with a small resize handle — drag to resize; size metadata is preserved for DOCX export.

- Tables:
	- Insert a table from the toolbar.
	- Each table gets a small options button (ellipsis). Click it to open table actions.
	- Actions include: convert table shape (single column/row, N columns), add/delete row, add/delete column, and context-aware Add Above/Below / Add Left/Right when the cursor is inside a cell.
	- Conversions that are not meaningful are hidden (for example "Convert to single column" is hidden when the table already has a single column).

- Export:
	- Export the full document to DOCX using the `Export DOCX` button in the toolbar. Images and tables are converted to Word format where possible.
	- Export TOC structure as JSON from the TOC header.

## Toolbar Shortcuts & Actions

- Bold: Cmd/Ctrl+B
- Italic: Cmd/Ctrl+I
- Underline: Cmd/Ctrl+U
- Undo/Redo: Cmd/Ctrl+Z / Cmd/Ctrl+Y (or Cmd/Ctrl+Shift+Z)
- Insert numbered/bulleted lists: toolbar or Ctrl/Cmd+Shift+7 / Ctrl/Cmd+Shift+8
- Heading dropdown: converts the current block to `p`, `H1`–`H4` and syncs with the caret.

## Table Editing Workflow

1. Place the caret inside any table cell and click the table options (ellipsis) button.
2. Choose actions:
	 - Add row above / Add row below — inserts a new row relative to the selected row.
	 - Delete this row — removes the selected row (disabled if only one row remains).
	 - Add column left / Add column right — inserts a new column relative to the selected column.
	 - Delete this column — removes the selected column (disabled if only one column remains).
	 - Conversion actions (single column / single row / N columns) appear only when applicable.

If no cell is selected, the table menu offers end/start row/column actions instead.

## Development & Debugging

- The code is plain ES module JavaScript and CSS. Key source files:
	- `index.html` — main UI
	- `style.css` — styling
	- `script.js` — functionality (TOC, templates, tables, images, export)
- To check for syntax errors in the editor script (quick):

```bash
node --check script.js
```

## Known Behaviors & Tips

- The TOC numbering is generated on render; if you programmatically modify headings, call the `renderTOC()` function or interact with the editor to refresh the TOC.
- The DOCX export uses the `docx` library loaded via importmap and `FileSaver.js` for downloads.
- For large images referenced by URL, the export may attempt to fetch the resource — ensure CORS allows it.

## Contributing

- Fork and PR if you want to add keyboard shortcuts for table actions, undo-friendly table edits, or persistence of sidebar width.

---
If you'd like, I can also add a short cheatsheet section to the UI or wire keyboard shortcuts for Add/Delete row/column — which would you prefer next?
