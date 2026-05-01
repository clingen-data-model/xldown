from pathlib import Path

import click

from markitdownite.converter import excel_to_markdown


@click.command()
@click.argument("xlsx_path", type=click.Path(exists=True, dir_okay=False))
@click.option(
    "-o",
    "--output",
    "output_dir",
    default=None,
    help="Output folder. Defaults to <input_name>_output in the same directory.",
)
def main(xlsx_path: str, output_dir: str | None) -> None:
    """Convert an Excel file to Markdown with charts and images.

    Creates a folder containing output.md, charts/, and images/ subdirectories.
    """
    input_path = Path(xlsx_path)
    if output_dir is None:
        output_dir = str(input_path.with_stem(input_path.stem + "_output"))

    excel_to_markdown(input_path, output_dir)
    click.echo(f"Written to: {output_dir}")


if __name__ == "__main__":
    main()
