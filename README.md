# xldown

[![PyPI](https://img.shields.io/pypi/v/xldown.svg)](https://pypi.org/project/xldown/)
[![Tests](https://github.com/clingen-data-model/xldown/workflows/Tests/badge.svg)](https://github.com/clingen-data-model/xldown/actions)

Convert Excel (`.xlsx`) files to Markdown.

## Install

```bash
uv pip install -e /path/to/xldown
```

## CLI

```bash
xldown input.xlsx              # creates input_output/ folder
xldown input.xlsx -o my_report # creates my_report/ folder
xldown --help
```

Output folder structure:
```
my_report/
├── output.md      # converted markdown with tables and chart links
├── charts/        # rendered chart images (1.png, 2.png, ...)
└── images/        # extracted embedded images (1.png, 2.png, ...)
```

## Python API

```python
from xldown import excel_to_markdown

excel_to_markdown("data.xlsx", "my_report/")
```

Creates `my_report/` with `output.md`, `charts/`, and `images/` subdirectories.

## Dependencies

- pandas
- openpyxl
- matplotlib
- click
- tabulate
- pydantic

## Excel Edge Cases Handled

The converter is designed to gracefully handle common Excel edge cases without failing or losing data:

### Worksheet and Cell-Level

- **Empty worksheets**: Worksheets with no cell content are skipped entirely (no output generated)
- **Prose cells**: Single isolated cells are rendered as plain text paragraphs
- **Row length variance**: Rows may have different numbers of cells; they are padded to the region's width before table construction
- **Merged cells**: Merged cell ranges are filled with the top-left cell's value and formatting applied to all cells in the range
- **Hidden columns**: Columns marked as hidden in the worksheet are detected and labeled with "(hidden)" in the table header
- **Cell formatting**: Rich text with character-level subscript/superscript (e.g., H₂O) is detected and rendered as `<sub>` / `<sup>` HTML tags; cell-level formatting (bold, italic, strikethrough, superscript, subscript, rotation) is applied as Markdown or HTML annotations
- **Cell colors and borders**: Font colors, background colors, and border styles are extracted and documented in an Annotations section below each table (filtering out default black/white)
- **Cell metadata**: Comments and hyperlinks are extracted and documented with cell coordinates below each table

### Data Organization

- **Non-contiguous regions**: Adjacent cells are grouped into connected components (4-connected flood-fill), and isolated cells are treated as prose while multi-cell regions become tables
- **Annotation grouping**: Cells with identical formatting annotations are grouped into connected components; solid rectangles are expressed as ranges (e.g., `A1:C3`), while irregular patterns list individual cells

### Chart Edge Cases

- **Missing or invalid data**: Empty charts, missing sheets, and malformed range references are silently skipped
- **Data length mismatches**: Series with varying lengths are padded with zeros; missing category labels are replaced with numeric indices
- **Missing attributes**: Unset or None chart attributes default to sensible values (e.g., "clustered" for bar grouping)

### Chart Type-Specific Handling

- **Single-series charts** (Pie, Doughnut, Radar): Only the first series is plotted
- **Stacked charts**: Series are stacked correctly, with percent-stacked variants normalized to 100%
- **Minimum requirements** (Stock, Surface): Charts requiring specific data combinations may be skipped if incomplete
- **Coordinate systems**: Charts using special projections (3D, polar) are rendered with appropriate matplotlib settings
