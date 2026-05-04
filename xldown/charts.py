import math
from pathlib import Path

import matplotlib
from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart,
    AreaChart3D,
    BarChart,
    BarChart3D,
    BubbleChart,
    DoughnutChart,
    LineChart,
    LineChart3D,
    PieChart,
    PieChart3D,
    ProjectedPieChart,
    RadarChart,
    ScatterChart,
    Series,
    StockChart,
    SurfaceChart,
    SurfaceChart3D,
)

matplotlib.use("Agg")
import matplotlib.pyplot as plt

ChartType = (
    BarChart
    | BarChart3D
    | LineChart
    | LineChart3D
    | PieChart
    | PieChart3D
    | ProjectedPieChart
    | DoughnutChart
    | AreaChart
    | AreaChart3D
    | ScatterChart
    | BubbleChart
    | RadarChart
    | StockChart
    | SurfaceChart
    | SurfaceChart3D
)

# ---------------------------------------------------------------------------
# Cell-range helpers
# ---------------------------------------------------------------------------


def _read_cell_range(workbook: Workbook, range_str: str) -> list:
    """Resolve a range like \"'Sheet1'!$A$1:$A$10\" to a flat list of non-None values."""
    range_str = range_str.replace("$", "")
    if "!" not in range_str:
        return []
    sheet_part, cell_range = range_str.split("!", 1)
    sheet_name = sheet_part.strip("'")
    try:
        ws = workbook[sheet_name]
    except KeyError:
        return []
    values: list = []
    try:
        cells = ws[cell_range]
        if not isinstance(cells, tuple):
            cells = (cells,)
        for row in cells:
            if not isinstance(row, tuple):
                row = (row,)
            for cell in row:
                if cell.value is not None:
                    values.append(cell.value)
    except Exception:
        pass
    return values


def _read_ref(workbook: Workbook, container) -> list:
    """Read values from a numRef or strRef container."""
    if container is None:
        return []
    for attr in ("numRef", "strRef"):
        ref_obj = getattr(container, attr, None)
        if ref_obj and ref_obj.ref:
            return _read_cell_range(workbook, ref_obj.ref)
    return []


def _numeric(values: list) -> list[float]:
    return [v for v in values if isinstance(v, (int, float))]


def _series_label(workbook: Workbook, series: Series) -> str | None:
    if not series.title:
        return None
    if hasattr(series.title, "v") and series.title.v:
        return str(series.title.v)
    if (
        hasattr(series.title, "strRef")
        and series.title.strRef
        and series.title.strRef.f
    ):
        values = _read_cell_range(workbook, series.title.strRef.f)
        if values:
            return str(values[0])
    return None


# ---------------------------------------------------------------------------
# Per-chart-type renderers
# ---------------------------------------------------------------------------


def _render_bar(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: Extract grouping/orientation → collect series data and categories → pad series to same length
    # → apply percent-stacking if needed → draw each series as bars → return True if any data plotted
    grouping = getattr(chart, "grouping", "clustered") or "clustered"
    horizontal = getattr(chart, "type", "col") == "bar"
    stacked = grouping in ("stacked", "percentStacked")
    pct = grouping == "percentStacked"

    series_list: list[tuple[list, str | None]] = []
    cats: list | None = None

    for series in chart.series:
        y = _numeric(_read_ref(workbook, getattr(series, "val", None)))
        x = _read_ref(workbook, getattr(series, "cat", None))
        if cats is None and x:
            cats = x
        series_list.append((y, _series_label(workbook, series)))

    if not any(y for y, _ in series_list):
        return False

    n = max(len(y) for y, _ in series_list)
    xs = list(range(n))
    tick_labels = [str(c) for c in cats[:n]] if cats else [str(i) for i in xs]

    if pct:
        padded = [y + [0] * (n - len(y)) for y, _ in series_list]
        totals = [sum(s[i] for s in padded) for i in range(n)]
        series_list = [
            ([s[i] / totals[i] * 100 if totals[i] else 0 for i in range(n)], lbl)
            for s, (_, lbl) in zip(padded, series_list)
        ]

    bottoms = [0.0] * n
    for y, label in series_list:
        if not y:
            continue
        vals = y + [0] * (n - len(y))
        if horizontal:
            ax.barh(xs, vals, left=bottoms if stacked else None, label=label)
        else:
            ax.bar(xs, vals, bottom=bottoms if stacked else None, label=label)
        if stacked:
            bottoms = [b + v for b, v in zip(bottoms, vals)]

    if horizontal:
        ax.set_yticks(xs)
        ax.set_yticklabels(tick_labels)
    else:
        ax.set_xticks(xs)
        ax.set_xticklabels(tick_labels, rotation=45, ha="right")

    return True


def _render_line(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: For each series, extract Y values → extract or generate X values → plot line with markers
    plotted = False
    for series in chart.series:
        y = _numeric(_read_ref(workbook, getattr(series, "val", None)))
        x = _read_ref(workbook, getattr(series, "cat", None))
        label = _series_label(workbook, series)
        if not y:
            continue
        xs = x if (x and len(x) == len(y)) else list(range(len(y)))
        ax.plot(xs, y, marker="o", label=label)
        plotted = True
    return plotted


def _render_pie(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: Iterate series until one has data → extract Y values and optional labels → draw pie with percentages
    # Note: Only first series is rendered (pie charts don't support multiple series)
    for series in chart.series:
        y = _numeric(_read_ref(workbook, getattr(series, "val", None)))
        x = _read_ref(workbook, getattr(series, "cat", None))
        if not y:
            continue
        labels = [str(v) for v in x] if x else None
        ax.pie(y, labels=labels, autopct="%1.1f%%")
        # Pie charts can only display one series; return immediately after plotting
        return True
    return False


def _render_doughnut(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: Extract hole size → iterate series until one has data → draw pie with hollow center via wedgeprops
    # Note: Only first series is rendered (doughnut charts don't support multiple series)
    hole = (getattr(chart, "holeSize", 50) or 50) / 100
    for series in chart.series:
        y = _numeric(_read_ref(workbook, getattr(series, "val", None)))
        x = _read_ref(workbook, getattr(series, "cat", None))
        if not y:
            continue
        labels = [str(v) for v in x] if x else None
        ax.pie(y, labels=labels, autopct="%1.1f%%", wedgeprops={"width": 1 - hole})
        # Doughnut charts can only display one series; return immediately after plotting
        return True
    return False


def _render_area(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: Extract grouping → collect series data and categories → pad series to same length
    # → apply percent-stacking if needed → draw stacked areas or overlapping filled areas
    grouping = getattr(chart, "grouping", "standard") or "standard"
    stacked = grouping in ("stacked", "percentStacked")
    pct = grouping == "percentStacked"

    series_list: list[tuple[list, str | None]] = []
    cats: list | None = None

    for series in chart.series:
        y = _numeric(_read_ref(workbook, getattr(series, "val", None)))
        x = _read_ref(workbook, getattr(series, "cat", None))
        if cats is None and x:
            cats = x
        series_list.append((y, _series_label(workbook, series)))

    if not any(y for y, _ in series_list):
        return False

    n = max(len(y) for y, _ in series_list)
    xs = list(cats[:n]) if cats else list(range(n))

    if pct:
        padded = [y + [0] * (n - len(y)) for y, _ in series_list]
        totals = [sum(s[i] for s in padded) for i in range(n)]
        series_list = [
            ([s[i] / totals[i] * 100 if totals[i] else 0 for i in range(n)], lbl)
            for s, (_, lbl) in zip(padded, series_list)
        ]

    if stacked:
        ys_arrays = [y + [0] * (n - len(y)) for y, _ in series_list]
        labels = [lbl or "" for _, lbl in series_list]
        ax.stackplot(xs, *ys_arrays, labels=labels)
    else:
        for y, label in series_list:
            if not y:
                continue
            ax.fill_between(xs[: len(y)], y, alpha=0.4, label=label)
            ax.plot(xs[: len(y)], y)

    return True


def _render_scatter(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: For each series, extract X and Y values → match lengths by taking minimum → plot as scatter or line
    style = getattr(chart, "scatterStyle", "marker") or "marker"
    plotted = False
    for series in chart.series:
        x = _numeric(_read_ref(workbook, getattr(series, "xVal", None)))
        y = _numeric(_read_ref(workbook, getattr(series, "yVal", None)))
        label = _series_label(workbook, series)
        if not (x and y):
            continue
        n = min(len(x), len(y))
        if style in ("line", "lineMarker", "smooth", "smoothMarker"):
            ax.plot(x[:n], y[:n], marker="o", label=label)
        else:
            ax.scatter(x[:n], y[:n], label=label)
        plotted = True
    return plotted


def _render_bubble(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: For each series, extract X, Y, and optional bubble size values → match lengths
    # → normalize bubble sizes or use uniform size → plot as scatter with sized markers
    plotted = False
    for series in chart.series:
        x = _numeric(_read_ref(workbook, getattr(series, "xVal", None)))
        y = _numeric(_read_ref(workbook, getattr(series, "yVal", None)))
        sizes = _numeric(_read_ref(workbook, getattr(series, "bubbleSize", None)))
        label = _series_label(workbook, series)
        if not (x and y):
            continue
        n = min(len(x), len(y))
        if sizes and len(sizes) >= n:
            max_s = max(sizes[:n]) or 1
            s = [v / max_s * 500 for v in sizes[:n]]
        else:
            s = 100
        ax.scatter(x[:n], y[:n], s=s, alpha=0.6, label=label)
        plotted = True
    return plotted


def _render_radar(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: Extract filled/standard type → for each series, extract Y values → compute angular positions
    # → close the path by repeating first point → plot/fill on polar axes with optional category labels
    filled = getattr(chart, "type", "standard") == "filled"
    cats: list | None = None
    plotted = False

    for series in chart.series:
        y = _numeric(_read_ref(workbook, getattr(series, "val", None)))
        if cats is None:
            cats = _read_ref(workbook, getattr(series, "cat", None))
        label = _series_label(workbook, series)
        if not y:
            continue
        n = len(y)
        angles = [2 * math.pi * i / n for i in range(n)] + [0]
        values = y + [y[0]]
        if filled:
            ax.fill(angles, values, alpha=0.25, label=label)
        ax.plot(angles, values, label=label)
        if cats:
            ax.set_xticks([2 * math.pi * i / n for i in range(n)])
            ax.set_xticklabels([str(c) for c in cats[:n]])
        plotted = True

    return plotted


def _render_stock(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: Extract series in order (Open, High, Low, Close) → draw vertical lines for High-Low range
    # → draw colored bars for Open-Close spread (green if close ≥ open, red otherwise) → add category labels
    # Series order: Open, High, Low, Close (by append order in workbook)
    series_data: list[list] = []
    cats: list | None = None

    for series in chart.series:
        y = _numeric(_read_ref(workbook, getattr(series, "val", None)))
        if cats is None:
            x = _read_ref(workbook, getattr(series, "cat", None))
            if x:
                cats = x
        series_data.append(y)

    if len(series_data) < 2:
        return False

    # Extract in order: Open, High, Low, Close
    opens = series_data[0] if len(series_data) > 0 else []
    highs = series_data[1] if len(series_data) > 1 else []
    lows = series_data[2] if len(series_data) > 2 else []
    closes = series_data[3] if len(series_data) > 3 else []

    if not (highs and lows):
        return False

    n = min(len(highs), len(lows))
    xs = list(range(n))

    for i in xs:
        ax.vlines(i, lows[i], highs[i], color="black", linewidth=1.5)

    if opens and closes and len(opens) >= n and len(closes) >= n:
        for i in xs:
            color = "green" if closes[i] >= opens[i] else "red"
            ax.bar(
                i,
                abs(closes[i] - opens[i]),
                bottom=min(opens[i], closes[i]),
                color=color,
                width=0.4,
            )

    if cats:
        ax.set_xticks(xs)
        ax.set_xticklabels([str(c) for c in cats[:n]], rotation=45, ha="right")

    return True


def _render_surface(workbook: Workbook, chart: ChartType, ax) -> bool:
    # Flow: Collect all series as rows of Z values → pad ragged rows → create 2D meshgrid for X/Y
    # → plot 3D surface or 2D contour with colormap based on chart type
    import numpy as np

    rows: list[list] = []
    for series in chart.series:
        y = _numeric(_read_ref(workbook, getattr(series, "val", None)))
        if y:
            rows.append(y)

    if not rows:
        return False

    n_cols = max(len(r) for r in rows)
    Z = np.array([r + [0] * (n_cols - len(r)) for r in rows], dtype=float)

    if isinstance(chart, SurfaceChart3D):
        X, Y = np.meshgrid(range(n_cols), range(len(rows)))
        ax.plot_surface(X, Y, Z, cmap="viridis")
    else:
        im = ax.contourf(Z, cmap="viridis")
        plt.colorbar(im, ax=ax)

    return True


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------


def render_chart(workbook: Workbook, chart: ChartType, output_path: Path) -> bool:
    """Render an openpyxl chart to a PNG using matplotlib. Returns True if data was plotted."""
    fig = plt.figure(figsize=(8, 5))

    match chart:
        case SurfaceChart3D():
            # SurfaceChart3D renders as a true 3D surface; others render as 2D
            ax = fig.add_subplot(111, projection="3d")
        case RadarChart():
            ax = fig.add_subplot(111, projection="polar")
        case _:
            ax = fig.add_subplot(111)

    plotted = False
    match chart:
        case BarChart() | BarChart3D():
            plotted = _render_bar(workbook, chart, ax)
        case LineChart() | LineChart3D():
            plotted = _render_line(workbook, chart, ax)
        case PieChart() | PieChart3D() | ProjectedPieChart():
            plotted = _render_pie(workbook, chart, ax)
        case DoughnutChart():
            plotted = _render_doughnut(workbook, chart, ax)
        case AreaChart() | AreaChart3D():
            plotted = _render_area(workbook, chart, ax)
        case ScatterChart():
            plotted = _render_scatter(workbook, chart, ax)
        case BubbleChart():
            plotted = _render_bubble(workbook, chart, ax)
        case RadarChart():
            plotted = _render_radar(workbook, chart, ax)
        case StockChart():
            plotted = _render_stock(workbook, chart, ax)
        case SurfaceChart() | SurfaceChart3D():
            plotted = _render_surface(workbook, chart, ax)

    if plotted:
        try:
            handles, _ = ax.get_legend_handles_labels()
            if handles:
                ax.legend()
        except Exception:
            pass
        fig.savefig(output_path, bbox_inches="tight", dpi=100)

    plt.close(fig)
    return plotted
