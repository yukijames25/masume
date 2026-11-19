![masume — 神エクセルの申請書を、AIに壊さず書かせる。](assets/hero.png)

# masume（升目）

**Excel方眼紙（神エクセル）の申請書を、AIエージェントが“壊さずに”読み書きするためのツール。**
Python 標準ライブラリだけで動きます（依存ゼロ・約110KB）。

*masume fills unmodified Japanese "Excel 方眼紙" government/office forms (merged cells, □ checkboxes, 和暦 dates) for AI agents, without breaking the layout. Zero dependencies; CLI, MCP server, Claude Code skill and an offline single-file web version.*

紹介記事: [神エクセルの申請書を、AIに“壊さず”書かせるツールを作った（Zenn）](https://zenn.dev/yukidev25/articles/fe349484127386)

## インストール

```bash
pip install git+https://github.com/yukijames25/masume
masume --help
```

Python 3.9 以上。依存ライブラリはありません。インストールせずに試すなら、ブラウザ版 `web/index.html` を開くだけで使えます。

```
$ python -m masume inspect 休暇届.xlsx
休暇届!T5:AG5       日付(セル内)    [休暇届] 日付
休暇届!H8:S8        テキスト       [申請者] 所属
休暇届!H9:AG9       テキスト       [申請者] 氏名
休暇届!H12:AG12     選択         [休暇内容] 休暇の種類  選択肢=年次有給休暇/特別休暇/代休/その他
休暇届!H13:P13      日付         [休暇内容] 開始日
休暇届!B16:AG19     テキスト       [休暇内容] 理由
休暇届!B21:AG21     テキスト(セル内)  [休暇内容] 備考
```

```
$ python -m masume fill 休暇届.xlsx values.json -o 休暇届_記入済.xlsx
```

```json
{ "日付": "2026-11-19", "氏名": "山田 太郎", "休暇の種類": "年次有給休暇",
  "開始日": "2026-12-01", "理由": "私用のため", "備考": "特になし" }
```

→ `令和8年11月19日`、`■年次有給休暇　□特別休暇…`、`備考：特になし` のように、様式の書式に合わせて書き込みます。

## 何がうれしいか

- **テンプレートを加工しない**: `{{name}}` を仕込む必要なし。結合セルと罫線から「ラベル → 入力欄」を自動で見つけます。
- **壊さない**: 書き込むセル以外（罫線・結合・画像・印影・印刷設定）はバイト単位でそのまま。数式セルには書き込みません。
- **日本の書式がわかる**: 和暦（令和元年も）、元号リスト「明・大・昭・平・令」、□チェック（1セル／別セル）、「男 ・ 女」、「令和　年　月　日」、期間「～」、〒・電話・時刻の書式、1文字1マス、分割欄、＿＿下線、均等割付ラベル（氏　名）、ふりがな、明細表・学歴職歴表、縦書きのグループ見出し（請求者／必要な方）、記入例シートの除外、押印欄。
- **役所の様式の書き方にも対応**: 実在の自治体様式で見つけた書き方を取り込みました。大きな罫線枠の中のラベル、2行ラベル「所在地／（住所）」、セルに分かれた「令和[ ]年[ ]月[ ]日提出」、1文字ずつ結合した法人番号、Excelの**チェックボックス**（開発タブのフォームコントロール）、「事項｜変更前｜変更後」の比較表（id は `所在地（住所）/変更後`）、受付印・市記入欄の除外。
- **AIから使える**: MCPサーバー（依存ゼロの自前実装）と Claude Code スキルを同梱。
- **ローカル完結**: 社外秘・個人情報入りの様式でも外部に送りません。

## 使い方

```bash
# このフォルダで（インストールせずに使う場合）
set PYTHONPATH=src           # PowerShell なら: $env:PYTHONPATH="src"
python -m masume md       form.xlsx            # LLM向けMarkdown（入力欄一覧＋論理レイアウト）
python -m masume inspect  form.xlsx --json     # 入力欄をJSONで
python -m masume template form.xlsx > v.json   # 記入用の雛形
python -m masume fill     form.xlsx v.json -o out.xlsx

# インストール済みなら
masume inspect form.xlsx
masume fill form.xlsx v.json -o out.xlsx --brief   # --brief: 書いた内容を欄ごとに1行で表示
```

### 値の書き方

| 欄の種類 | 値の例 |
|---|---|
| テキスト | `"山田 太郎"` / 数値 `3` |
| 日付 | `"2026-11-19"`（和暦の様式なら自動で令和表記。「昭和・平成」の元号リスト、（満　歳）の年齢も自動） |
| 期間（～） | `"2026-12-01～2026-12-03"`（（　日間）も自動計算） |
| 選択（□） | `"振込"` / `["A", "B"]` / `"その他（祖父）"` |
| 選択（男 ・ 女） | `"女"` → 「男 ・ （女）」 |
| 書式入り欄 | 〒`"100-0001"`、電話`"03-1234-5678"`、時刻`"9:30"` → 空白に順に埋める |
| 分割欄 | `"5年3ヶ月"` / `[5, 3]` / `{"年": 5, "ヶ月": 3}`、電話の3分割、郵便番号の2分割 |
| 1文字1マス | `"1234567"`（口座番号・マイナンバー等） |
| 表 | `[{"日付": "11/1", "内容": "交通費", "金額": 2300}, ...]` |
| 検出されなかった欄 | セル番地をキーに: `"C12": "..."`, `"シート名!C12": "..."` |

### 記入後の確認

記入した値がラベルに見えて id が変わらないよう、記入前に欄の定義を保存しておき、確認のときに渡します。

```bash
python -m masume inspect form.xlsx --json > fields.json
python -m masume fill form.xlsx v.json -o out.xlsx
python -m masume md out.xlsx --fields fields.json     # 各欄の「現在値」を表示
```

`template` の値はすべて `null` です。null の欄は書き込まれないので、分かる欄だけ埋めれば済みます。

## ブラウザ版（インストール不要）

`web/index.html` を開くだけで使えます（1ファイル・オフライン可）。ページ自体が外部への通信を禁止（CSP `default-src 'none'`）しているため、
社外秘の様式も端末の外に出ません。検出と記入のロジックはPython版と同じで、`node web/test_core.mjs` で
全テスト様式と300個のランダムなワークブックについて、Python版と出力が一致するかを確かめています。

## Claude Code / MCP で使う

このフォルダを Claude Code で開くと `.claude/skills/masume`（スキル）と `.mcp.json`（MCPサーバー）が使えます。
他のプロジェクトから使うなら:

```bash
claude mcp add masume -- python -m masume.mcp_server   # pip install -e . 済みの場合
```

ツール: `inspect_form` / `form_to_markdown` / `fill_form` / `list_sheets`

## 制限

- `.xls`（旧形式）は `masume convert` で `.xlsx` にしてから使います（Excel か LibreOffice が必要）。
- 検出はヒューリスティックです。外れた欄はセル番地で直接指定するか、`inspect --json` の結果を直して `--fields` で使い回せます。
- 実在の自治体様式6件での結果は [docs/REAL_FORMS.md](docs/REAL_FORMS.md)。1ページの届出書・申請書は実用レベル、大きな様式や数値明細は手直しが必要です。

## 開発

```bash
python -m unittest discover -s tests       # テスト（openpyxl があれば互換チェックも）
python tests/make_fixtures.py              # テスト用様式の再生成（要 openpyxl）
```

なぜこれを作ったか（市場調査）は [docs/RESEARCH.md](docs/RESEARCH.md)。

## うまくいかないとき

検出が外れた様式、記入がずれた様式があれば、[Issue](https://github.com/yukijames25/masume/issues) で教えてください。
様式そのものを添付できない場合は、`masume md 様式.xlsx` の出力（個人情報を含まない記入前のもの）を貼ってもらえると助かります。

License: MIT
