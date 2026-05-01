# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Setup and Dependencies

Use **`uv`** for all package and environment management. The project is configured with `uv.lock` and `pyproject.toml`.

```bash
uv pip install -e ".[dev]"    # Install package with dev dependencies
```

## Common Commands

**Run tests:**
```bash
uv run pytest                              # Run all tests
uv run pytest tests/test_converter.py -v  # Run a single test file with verbose output
uv run pytest tests/test_charts.py::test_all_charts_in_single_workbook -v  # Run a specific test
```

**Lint and format:**
```bash
ruff check markitdownite tests             # Check code style
ruff format markitdownite tests            # Auto-format code
```

**CLI usage:**
```bash
uv run markitdownite input.xlsx                          # Convert and output to input.md
uv run markitdownite input.xlsx -o report.md             # Explicit output path
uv run markitdownite input.xlsx --image-dir assets/img   # Custom image directory
```

## Architecture

### Core Design

The tool converts Excel workbooks to Markdown documents by:
1. Reading each worksheet as a pandas DataFrame → Markdown table
2. Rendering each chart to a PNG using matplotlib
3. Extracting embedded raster images from cells
4. Assembling everything into a single Markdown file with image links

### Key Modules

**`converter.py`** — Main conversion logic with three entry points:

- **Chart rendering system** (`render_chart()` and `_render_*()` functions):
  - Supports 15+ chart types: Bar (horizontal/stacked), Line, Pie, Doughnut, Area, Scatter, Bubble, Radar, Stock, Surface
  - Extracts data from openpyxl chart objects using internal `numRef`/`strRef` XML properties
  - Renders via matplotlib with appropriate axes (polar for radar, 3D for surface)
  - Returns `True` only if data was actually plotted (empty charts are skipped)

- **Helper functions** (cell-range resolution):
  - `_read_cell_range()` — Parses Excel range strings like `"'Sheet1'!$A$1:$A$10"`
  - `_read_ref()` — Extracts numeric/string cell values from chart data sources
  - `_numeric()` — Filters lists to keep only numeric values (for plotting)

- **Public API**:
  - `excel_to_markdown(xlsx_path, output_md, image_dir)` — Main entry point
  - `extract_images()` — Extracts embedded raster images from worksheets
  - `df_to_markdown()` — Renders DataFrames as GH-flavored Markdown tables

**`cli.py`** — Minimal Click-based CLI that accepts an Excel path and optional output/image directory flags.

### Testing

**`test_converter.py`** — Integration tests that generate test Excel files in-memory using openpyxl and verify:
- Table data appears in Markdown
- Charts are rendered to PNG
- Images are embedded as Markdown links

**`test_charts.py`** — Parametrized tests covering all supported chart types:
- Each chart factory (`make_bar()`, `make_pie()`, etc.) creates a sample workbook
- Consolidated into a single test workbook with one sheet per chart type
- Verifies `render_chart()` successfully renders each type
- Includes helper `_update_chart_references()` to remap sheet names in chart formulas (openpyxl stores references like `'Data'!$B$1:$B$5`)

### Dependencies

- **pandas** — Read Excel sheets to DataFrames
- **openpyxl** — Load workbooks, access chart objects and embedded images
- **matplotlib** — Render charts to PNG
- **click** — CLI framework
- **tabulate** — DataFrame-to-Markdown conversion (via pandas `.to_markdown()`)

## Design Notes

- Chart rendering uses matplotlib's Agg backend (`matplotlib.use("Agg")`) for headless operation
- Empty charts (no data to plot) are silently skipped; `render_chart()` returns `False` as the signal
- Cell reference parsing is defensive; malformed ranges return empty lists rather than raising exceptions
- Tests generate Excel fixtures programmatically; there are no pre-stored .xlsx files
