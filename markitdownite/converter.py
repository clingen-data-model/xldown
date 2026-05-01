from pathlib import Path

from markitdownite import paths
from markitdownite.charts import render_chart
import pandas as pd
from openpyxl import load_workbook


def extract_images(workbook, output_dir: Path) -> dict:
    """Extract all embedded raster images from every worksheet.

    Returns a dict mapping sheet_name to list of image indices for that sheet.
    """
    images_dir = paths.images_dir_path(output_dir)
    images_dir.mkdir(parents=True, exist_ok=True)

    img_counter = 0
    sheet_images: dict[str, list[int]] = {}

    for ws in workbook.worksheets:
        if getattr(ws, "_images", []):
            sheet_images[ws.title] = []
        for img in getattr(ws, "_images", []):
            path = paths.image_path(output_dir, img_counter)
            with open(path, "wb") as f:
                f.write(img._data())
            sheet_images[ws.title].append(img_counter)
            img_counter += 1

    return sheet_images


def df_to_markdown(df: pd.DataFrame) -> str:
    """Render a DataFrame as a GitHub-flavoured Markdown table."""
    return df.to_markdown(index=False)


def excel_to_markdown(
    xlsx_path: str | Path,
    output_dir: str | Path,
) -> None:
    """Convert an Excel workbook to a folder with Markdown and assets.

    Args:
        xlsx_path: Path to the input Excel file.
        output_dir: Directory where output will be written. Created if it doesn't exist.

    Creates output_dir with:
        output.md: Markdown file with worksheets as sections. Each section contains:
            - Level-1 heading with worksheet name
            - Level-2 "Table" heading with DataFrame rendered as GH-flavored table
            - Image links for each rendered chart and embedded image
        charts/: Numbered PNG images (0.png, 1.png, ...) for each chart found
        images/: Numbered PNG images (0.png, 1.png, ...) for embedded images

    Charts are rendered via matplotlib. Empty charts are skipped.
    Embedded images are extracted from worksheet cells and included in markdown.
    """
    xlsx_path = Path(xlsx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    charts_dir = paths.charts_dir_path(output_dir)
    charts_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(xlsx_path, data_only=True)

    sheet_images = extract_images(wb, output_dir)

    md_parts: list[str] = []
    chart_counter = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name)

        md_parts.append(f"# {sheet_name}\n")

        md_parts.append("## Table\n")
        md_parts.append(df_to_markdown(df))
        md_parts.append("\n")

        for chart in getattr(ws, "_charts", []):
            chart_path = paths.chart_path(output_dir, chart_counter)
            if render_chart(wb, chart, chart_path):
                md_parts.append(f"![Chart]({chart_path})\n")
            chart_counter += 1

        for img_idx in sheet_images.get(sheet_name, []):
            img_path = paths.image_path(output_dir, img_idx)
            md_parts.append(f"![Image]({img_path})\n")

    paths.output_file_path(output_dir).write_text("\n".join(md_parts))
