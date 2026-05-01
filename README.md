# markitdownite

Convert Excel (`.xlsx`) files to Markdown.

## Install

```bash
pip install -e /path/to/markitdownite
```

## CLI

```bash
markitdownite input.xlsx              # creates input_output/ folder
markitdownite input.xlsx -o my_report # creates my_report/ folder
markitdownite --help
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
from markitdownite import excel_to_markdown

excel_to_markdown("data.xlsx", "my_report/")
```

Creates `my_report/` with `output.md`, `charts/`, and `images/` subdirectories.

## Dependencies

- pandas
- openpyxl
- tabulate
- click
