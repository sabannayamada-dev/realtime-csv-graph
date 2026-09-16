from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any

os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "matplotlib-cache"))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import japanize_matplotlib  # noqa: F401  Japanese font bundled for headless/Linux runs.

from .data_io import numeric_series, read_data


def _resolve_column(frame, column: str | int) -> str:
    if isinstance(column, int):
        try:
            return str(frame.columns[column])
        except IndexError as exc:
            raise ValueError(f"列番号 {column} は範囲外です") from exc
    return str(column)


def generate_graph(config: dict[str, Any], base_dir: Path, output_dir: Path) -> Path:
    input_path = (base_dir / config["input"]).resolve()
    frame = read_data(input_path, header=config.get("header", 0))
    x_column = _resolve_column(frame, config.get("x", 0))
    x_values = numeric_series(frame, x_column)
    series_configs = config.get("series") or [
        {"column": str(column)} for column in frame.columns if str(column) != x_column
    ]
    if not series_configs:
        raise ValueError("Y軸に使用する系列が指定されていません")

    style = config.get("style", {})
    fig, ax = plt.subplots(
        figsize=tuple(style.get("figsize", [8.0, 5.2])),
        dpi=int(style.get("dpi", 160)),
    )
    plotted = 0
    for series in series_configs:
        column = _resolve_column(frame, series["column"])
        y_values = numeric_series(frame, column)
        valid = x_values.notna() & y_values.notna()
        if not valid.any():
            continue
        options = {
            "label": series.get("label", column),
            "color": series.get("color"),
            "linewidth": float(series.get("linewidth", 1.8)),
            "marker": series.get("marker", "o"),
            "markersize": float(series.get("markersize", 4.0)),
        }
        options = {key: value for key, value in options.items() if value is not None}
        if series.get("plot", "line") == "scatter":
            options.pop("linewidth", None)
            marker_size = options.pop("markersize", 4.0)
            ax.scatter(x_values[valid], y_values[valid], s=marker_size**2, **options)
        else:
            ax.plot(x_values[valid], y_values[valid], **options)
        plotted += 1
    if plotted == 0:
        raise ValueError("描画できる数値データがありません")

    ax.set_title(config.get("title", ""), fontsize=float(style.get("title_size", 14)))
    ax.set_xlabel(config.get("x_label", x_column), fontsize=float(style.get("label_size", 11)))
    ax.set_ylabel(config.get("y_label", ""), fontsize=float(style.get("label_size", 11)))
    if config.get("x_log"):
        ax.set_xscale("log")
    if config.get("y_log"):
        ax.set_yscale("log")
    if config.get("x_limits"):
        ax.set_xlim(*config["x_limits"])
    if config.get("y_limits"):
        ax.set_ylim(*config["y_limits"])
    ax.grid(bool(style.get("grid", True)), alpha=0.28)
    if config.get("legend", True):
        ax.legend()
    fig.tight_layout()

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / config.get("output", f"{config.get('id', 'graph')}.png")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=int(style.get("export_dpi", 300)), bbox_inches="tight")
    plt.close(fig)
    return output_path.resolve()
