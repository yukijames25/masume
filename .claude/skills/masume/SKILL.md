---
name: masume
description: Read or fill a Japanese Excel form (Excel 方眼紙 / 神エクセル / 申請書 / 届出 / 精算書 .xlsx) without breaking its layout. Use when the user asks to fill in, 記入, 入力, 転記 into, or summarize an .xlsx form whose layout uses merged cells, □ checkboxes or 令和　年　月　日 placeholders.
---

# masume — Excel 方眼紙フォームの読み書き

`masume` は依存ゼロの Python ツール。テンプレートは絶対に上書きせず、新しいファイルに書き出す。
書き込むセル以外（罫線・結合・画像・印影・印刷設定）はバイト単位で保持される。

## 前提

このリポジトリ直下で実行するなら `PYTHONPATH=src`（PowerShell: `$env:PYTHONPATH="src"`）を付ける。`pip install -e .` 済みなら不要。

## 手順

1. **構造を読む**（まずこれ。生のセルを読まない）
   ```
   python -m masume md <form.xlsx>
   ```
   入力欄一覧（id・種類・セル範囲・選択肢・表の列）と論理レイアウトが出る。
   機械可読にしたいときは `python -m masume inspect <form.xlsx> --json`。

2. **雛形を作る**
   ```
   python -m masume template <form.xlsx> > values.json
   ```

3. **値を埋める**（ユーザー提供の情報だけを使う。不明な欄は空のまま残し、最後にユーザーへ確認する）
   - 日付は `YYYY-MM-DD` で書く。フォームが和暦なら自動で「令和8年11月19日」になる。
   - 選択欄（□）は選択肢名を文字列か配列で。
   - 表は列idをキーにした行オブジェクトの配列。行数上限は inspect の `max_rows`。
   - 検出されなかった欄は `"C12"` や `"シート名!C12"` のセル番地をキーにして直接書ける。

4. **書き込む**
   ```
   python -m masume fill <form.xlsx> values.json -o <form>_記入済.xlsx
   ```
   出力 JSON の `errors` / `unknown_keys` が空であることを確認し、`unfilled` に残った欄をユーザーに伝える。

   検出が外れた様式は `inspect --json > fields.json` を手で直して `fill --fields fields.json` で使い回せる。

5. **検証**: `python -m masume md <form>_記入済.xlsx` で記入結果を読み直して報告する。

## 注意

- `.xls`（旧形式）は非対応。Excel で .xlsx に保存し直してもらう。
- 数式セルには書き込まない（ツールが拒否する）。合計欄などは Excel が開いたとき再計算する。
- 個人情報を扱うので、ユーザーが求めない外部送信はしない。

MCP として使う場合は同梱の `.mcp.json`（`python -m masume.mcp_server`）を有効化すると
`inspect_form` / `form_to_markdown` / `fill_form` / `list_sheets` ツールが使える。
