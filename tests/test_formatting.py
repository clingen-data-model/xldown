from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.comments import Comment
from openpyxl.worksheet.hyperlink import Hyperlink

from markitdownite.converter import excel_to_markdown
from markitdownite import paths


def test_prose_before_table(tmp_path: Path):
    """Verify converter handles prose text before a table."""
    wb = Workbook()
    ws = wb.active
    ws.title = "WithProse"

    # Add prose before the table
    ws["A1"] = "All values are in TPM units."

    # Add table starting at row 3
    ws["A3"] = "Gene"
    ws["B3"] = "Expression"
    ws["A4"] = "BRCA1"
    ws["B4"] = 1234
    ws["A5"] = "TP53"
    ws["B5"] = 5678

    test_excel = tmp_path / "test.xlsx"
    wb.save(test_excel)

    output_dir = tmp_path / "output"
    excel_to_markdown(test_excel, output_dir)

    content = paths.output_file_path(output_dir).read_text()
    expected = """# WithProse

## Table

| All values are in TPM units.   | nan        |
|:-------------------------------|:-----------|
| nan                            | nan        |
| Gene                           | Expression |
| BRCA1                          | 1234       |
| TP53                           | 5678       |

"""
    assert content == expected


def test_multiple_tables_same_sheet(tmp_path: Path):
    """Verify converter handles multiple separate tables in one sheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = "MultipleTables"

    # First table (rows 1-3)
    ws["A1"] = "Name"
    ws["B1"] = "Score"
    ws["A2"] = "Alice"
    ws["B2"] = 95
    ws["A3"] = "Bob"
    ws["B3"] = 87

    # Second table (rows 5-7, with gap)
    ws["A5"] = "Product"
    ws["B5"] = "Price"
    ws["A6"] = "Apple"
    ws["B6"] = 1.50
    ws["A7"] = "Banana"
    ws["B7"] = 0.75

    test_excel = tmp_path / "test.xlsx"
    wb.save(test_excel)

    output_dir = tmp_path / "output"
    excel_to_markdown(test_excel, output_dir)

    content = paths.output_file_path(output_dir).read_text()
    expected = """# MultipleTables

## Table

| Name    | Score   |
|:--------|:--------|
| Alice   | 95      |
| Bob     | 87      |
| nan     | nan     |
| Product | Price   |
| Apple   | 1.5     |
| Banana  | 0.75    |

"""
    assert content == expected


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

    # Red text (annotation) with comment
    ws["A6"] = "Eve"
    ws["A6"].font = Font(color="FFFF0000")
    ws["A6"].comment = Comment("This record needs review", "Author")
    ws["B6"] = "Error"
    ws["C6"] = 0

    # Frank with background on entire Status column
    ws["A7"] = "Frank"
    ws["B7"] = "Warning"
    ws["B7"].hyperlink = Hyperlink(ref="B7", target="https://example.com/docs")
    ws["C7"] = 25

    # Triangle pattern (tests non-rectangular cluster handling)
    ws["A8"] = "Grace"
    ws["B8"] = "OK"
    ws["C8"] = 10
    ws["A9"] = "left"
    ws["B9"] = "apex"
    ws["C9"] = "right"

    # Apply yellow background to entire Status column (annotation merging test)
    yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    for row in range(1, 8):
        ws[f"B{row}"].fill = yellow_fill

    # Apply blue background to triangle (annotation gap handling test)
    blue_fill = PatternFill(start_color="FF0000FF", end_color="FF0000FF", fill_type="solid")
    ws["B8"].fill = blue_fill
    ws["A9"].fill = blue_fill
    ws["C9"].fill = blue_fill

    test_excel = tmp_path / "test.xlsx"
    wb.save(test_excel)

    output_dir = tmp_path / "output"
    excel_to_markdown(test_excel, output_dir)

    content = paths.output_file_path(output_dir).read_text()
    expected = """# Formatting

## Table

| **Name**    | **Status**   | **Value**   |
|:------------|:-------------|:------------|
| **Alice**   | Active       | 100         |
| *Bob*       | Inactive     | 50          |
| ~~Charlie~~ | Pending      | 75          |
| ***Diana*** | Active       | 200         |
| Eve         | Error        | 0           |
| Frank       | Warning      | 25          |
| Grace       | OK           | 10          |
| left        | apex         | right       |


### Annotations

- B1:B7: bg_color=FFFFFF00

- A6: fg_color=FFFF0000

- B8: bg_color=FF0000FF

- A9: bg_color=FF0000FF

- C9: bg_color=FF0000FF

- A6: comment: This record needs review

- B7: link: https://example.com/docs


"""
    assert content == expected
