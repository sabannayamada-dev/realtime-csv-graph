# CSVグラフ・実験レポート作成ツール

CSV、Excel、LTspiceデータからグラフを作成し、実験レポート様式のWordファイルへまとめるPythonツールです。

## 2つの使い方

- `csv_graph_generator.py`: 従来のWindows画面版。ライブプレビューを確認しながらグラフを調整します。
- `experiment_report_pipeline`: 画面を使わない自動実行版。Codex Cloud、ConoHa、GitHub Actionsから実行できます。

自動実行版は元の画面版を変更せず追加しています。

## 自動実行版

```bash
python -m pip install -r requirements.txt
python -m experiment_report_pipeline.cli examples/pipeline_config.json
```

`outputs/sample` にグラフ画像、Wordレポート、`manifest.json` が作成されます。

## 設定ファイル

`examples/pipeline_config.json` をコピーし、次を変更します。

- `graphs[].input`: CSVまたはExcelファイル
- `graphs[].x`: X軸の列名または0始まりの列番号
- `graphs[].series`: Y軸系列、凡例、色、マーカー
- `report.title`: レポート名
- `report.metadata`: 実験日、氏名、学籍番号など
- `report.sections`: 目的、方法、結果、考察、参考文献など

本文では `[[FIG:図のID]]` と `[[TABLE:表のID]]` を使うと、図表番号へ自動変換されます。

## スマホから実行

GitHubのActions画面で `Build experiment report` を選び、`Run workflow` を押します。処理後、実行結果の `Artifacts` からWordファイルとグラフをダウンロードできます。

Codex Cloudへ指示する場合は、実験データと設定ファイルをリポジトリへ追加したうえで、例えば次のように依頼できます。

> `data/experiment.csv` を解析し、グラフと実験レポートを生成できるよう `configs/experiment.json` を作成してテストしてください。

## テスト

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

GitHubへ変更を送ると、同じテストが自動実行されます。

## 対応形式

- 入力: CSV、タブ区切りTXT、XLSX、XLS
- グラフ出力: PNG、PDF、SVGなどMatplotlibが対応する形式
- レポート出力: DOCX

実データ、個人情報、提出前のレポートは非公開リポジトリで管理してください。
