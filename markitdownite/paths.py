from pathlib import Path

# Directory names
CHARTS_DIR = "charts"
IMAGES_DIR = "images"
OUTPUT_FILE = "output.md"


def output_file_path(base_dir: Path) -> Path:
    return base_dir / OUTPUT_FILE


def charts_dir_path(base_dir: Path) -> Path:
    return base_dir / CHARTS_DIR


def images_dir_path(base_dir: Path) -> Path:
    return base_dir / IMAGES_DIR


def chart_path(base_dir: Path, idx: int) -> Path:
    return charts_dir_path(base_dir) / f"{idx}.png"


def image_path(base_dir: Path, idx: int) -> Path:
    return images_dir_path(base_dir) / f"{idx}.png"
