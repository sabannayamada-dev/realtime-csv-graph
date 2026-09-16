from __future__ import annotations

import argparse
import json
from pathlib import Path

from .graph import generate_graph
from .report import generate_report


def run_pipeline(config_path: Path, output_override: Path | None = None) -> dict[str, object]:
    config_path = config_path.resolve()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    base_dir = config_path.parent
    configured_output = Path(config.get("output_dir", "outputs"))
    output_dir = (output_override or (base_dir / configured_output)).resolve()

    graph_paths: dict[str, Path] = {}
    for graph_config in config.get("graphs", []):
        graph_id = graph_config.get("id")
        if not graph_id:
            raise ValueError("すべてのグラフに id が必要です")
        if graph_id in graph_paths:
            raise ValueError(f"グラフIDが重複しています: {graph_id}")
        graph_paths[graph_id] = generate_graph(graph_config, base_dir, output_dir)

    report_path = None
    if config.get("report"):
        report_path = generate_report(config["report"], base_dir, output_dir, graph_paths)

    manifest = {
        "config": str(config_path),
        "output_dir": str(output_dir),
        "graphs": {key: str(value) for key, value in graph_paths.items()},
        "report": str(report_path) if report_path else None,
    }
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    manifest["manifest"] = str(manifest_path)
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser(description="CSV/Excelからグラフと実験レポートを生成します")
    parser.add_argument("config", type=Path, help="パイプライン設定JSON")
    parser.add_argument("--output-dir", type=Path, help="出力先を一時的に上書き")
    args = parser.parse_args()
    result = run_pipeline(args.config, args.output_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
