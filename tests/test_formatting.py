from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from markitdownite.converter import excel_to_markdown
from markitdownite import paths


def test_all_formatting_types(tmp_path: Path):
    """Verify inline formatting (bold, italic, strikethrough) and annotations (colors)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Formatting"

    # Headers with bold
    ws["A1"] = "Name"
    ws["A1"].font = Font(bold=True)
    ws["B1"] = "Status"
    ws["B1"].font = Font(bold=True)
    ws["C1"] = "Value"
    ws["C1"].font = Font(bold=True)

    # Bold text
    ws["A2"] = "Alice"
    ws["A2"].font = Font(bold=True)
    ws["B2"] = "Active"
    ws["C2"] = 100

    # Italic text
    ws["A3"] = "Bob"
    ws["A3"].font = Font(italic=True)
    ws["B3"] = "Inactive"
    ws["C3"] = 50

    # Strikethrough text
    ws["A4"] = "Charlie"
    ws["A4"].font = Font(strikethrough=True)
    ws["B4"] = "Pending"
    ws["C4"] = 75

    # Bold + Italic
    ws["A5"] = "Diana"
    ws["A5"].font = Font(bold=True, italic=True)
    ws["B5"] = "Active"
    ws["C5"] = 200

    # Red text (annotation)
    ws["A6"] = "Eve"
    ws["A6"].font = Font(color="FFFF0000")
    ws["B6"] = "Error"
    ws["C6"] = 0

    # Yellow background (annotation)
    ws["A7"] = "Frank"
    ws["B7"] = "Warning"
    ws["B7"].fill = PatternFill(
        start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid"
    )
    ws["C7"] = 25

    test_excel = tmp_path / "test.xlsx"
    wb.save(test_excel)

    output_dir = tmp_path / "output"
    excel_to_markdown(test_excel, output_dir)

    content = paths.output_file_path(output_dir).read_text()
    expected = """# Formatting

## Table

| **Name**    | **Status**   |   **Value** |
|:------------|:-------------|------------:|
| **Alice**   | Active       |         100 |
| *Bob*       | Inactive     |          50 |
| ~~Charlie~~ | Pending      |          75 |
| ***Diana*** | Active       |         200 |
| Eve         | Error        |           0 |
| Frank       | Warning      |          25 |


### Annotations

- A6: fg_color=FFFF0000

- B7: bg_color=FFFFFF00


"""
    assert content == expected
