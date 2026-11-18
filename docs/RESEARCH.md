# ブルーオーシャン調査メモ（2026-11-19 夜）

## 選定基準

1. **軽い**: 依存ライブラリなし・数百KB以内で作れる（PC容量を食わない）
2. **個人で作れる**: サーバー・DB・大規模データ収集が不要
3. **競合が薄い**: 「同じ切り口」の既存品が見当たらない
4. **痛みが本物**: 日本固有で、毎日誰かが困っている
5. **追い風**: AIエージェント（Claude Code / MCP / ChatGPT）の普及で需要が伸びる

## 候補と競合状況

| 候補 | 競合状況 | 判定 |
|---|---|---|
| 官公庁オープンデータのMCPサーバー | 国交省が公式MCPを公開（2025/11）。e-Stat・e-Gov法令・gBizINFO・東京都オープンデータ・祝日なども個人実装が多数 | レッド寄り |
| Claude Codeのログ→日報自動生成 | claude-job-reporter（PyPI）、Classmethodの日報スキル、ccconv など既に複数 | レッド寄り |
| Excel方眼紙→Markdown（**読む**） | MarkItDown、xlsx2md（ブラウザ版）、ExStruct、Aspose、Nutrient など | 競合あり |
| 汎用Excel MCP | negokaz、sbroenne、SuperPyonchiX、openpyxl系など多数 | レッド |
| `{{変数}}` 差し込み型テンプレ記入 | elysiatools の xlsx template filler 等。ただし**テンプレの事前加工が前提** | 一部競合 |
| **無加工の方眼紙フォームにAIが「書く」** | ラベルから入力欄を自動で見つけて非破壊で書き戻す**再利用できるツール**は見当たらなかった（下の「追加調査」を参照） | **空白地帯（需要は確認済み）** |

### 追加調査（2026-11-19 04:15）

より近いものが2つ見つかった。どちらもツールとしての競合ではなく、**需要があることの裏付け**と判断した。

- [AIで申請書を自動で記入しよう！ #ClaudeCode（Qiita, 2026-06-23）](https://qiita.com/barcarunrun/items/42c0d8e30ef638af2a61)
  - Claude Code に方眼紙のセル一覧を出させてMarkdownテンプレートを作り、openpyxl で書き戻す**手順の記事**。ライブラリではない
  - 欄の自動検出、和暦・□・〒の扱い、図形や画像の保持はない
  - → 「AIに申請書を書かせたい」需要はすでにあり、masume はその手順を道具にしたもの
- [engchina/no.1-production-ready-suite PR #2520（2026-10-19 マージ）](https://github.com/engchina/no.1-production-ready-suite/pull/2520)
  - RAG の前処理として、方眼紙のラベルと値を点数で組にする**読み取り専用**の機能。大きなアプリの一部で、未リリース
  - → 検出の発想は近いが、書き込みと日本の様式の扱いはない
- PyPI に `masume` という名前のパッケージは見当たらなかった（公開直前に再確認）

## 結論: 「masume（升目）」

> 役所・社内の Excel 方眼紙の申請書を、**加工せずそのまま** AIエージェントに読ませて**書かせる**ツール。

### 空白になっている理由

- 「読む」ツールは増えたが、生成AIの利用者が本当に困っているのは**提出物を作る（書く）**方。
- 汎用Excel MCPはセル番地でしか書けない。方眼紙では「氏名の欄はどこか」をAIが推測しても外しやすい。
- openpyxl 系で保存し直すと、図形・画像・印影・一部の書式が落ちることがある（Excel MCP各種のREADMEにも注意書きがある）。役所の様式でこれは致命的。
- 和暦（令和　年　月　日）・□チェック・＿＿下線・均等割付ラベル（氏　名）といった**日本特有の書式**を扱うツールがない。海外勢は参入しにくい。

### masume の差別化

1. **テンプレ無改造**: `{{name}}` のような事前加工が不要。結合セル・罫線からラベルと入力欄の組を推定。
2. **非破壊書き込み**: シートXMLの該当セルだけを文字列単位で差し替える。他の部品はバイト単位で同一（テストで保証）。
3. **日本の書式を理解**: 和暦変換（令和元年対応）、□→■、「令和　年　月　日」の穴埋め、＿＿の置換、明細表の行。
4. **ゼロ依存・約100KB**: Python標準ライブラリのみ。MCPサーバーも自前実装。
5. **ローカル完結**: 社外秘・個人情報を含む様式でもクラウドへ送らない（ブラウザ版も同じ方針）。

### 想定ユーザー

- 総務・経理・人事の担当者（休暇届・経費精算・稟議書）
- 自治体の窓口業務とその利用者（各種申請書）
- SIer・社内SE（レガシーExcel業務のAI化）
- Claude Code / ChatGPT のヘビーユーザー（「この申請書、埋めといて」）

### マネタイズ案（後で検討）

- 本体はOSS（MIT）で認知を取る
- 有料: よく使う様式の**フィールド定義パック**（自治体様式・社保・税務）、ブラウザ版Pro（一括処理・履歴）、法人サポート
- Claude の Skill / MCP ディレクトリ、ChatGPT Apps への掲載

### リスク

- Microsoft Copilot in Excel が方眼紙を理解するようになると価値が下がる → 「ローカル・無料・エージェントから使える」で差別化を続ける
- `.xls`（旧形式）の様式がまだ多い → 変換手順の案内、将来 LibreOffice 連携
- 検出はヒューリスティック → 様式ごとの「フィールド定義の保存・共有」機能で補う（次の開発候補）

## 参考にした情報源

- [国土交通省が「MCPサーバ」公開（@IT）](https://atmarkit.itmedia.co.jp/ait/spv/2511/14/news056.html)
- [japan-corporate-mcp](https://enterprisedna.co/directories/mcp/yamariki-hub-japan-corporate-mcp) / [e-Gov法令MCP](https://mcpmarket.com/server/japan-e-gov-legislation) / [tokyo-opendata MCP](https://mcpmarket.com/ko/server/tokyo-opendata)
- [Claude Code のセッションログから日報…（Classmethod）](https://dev.classmethod.jp/articles/claude-code-daily-report-timesheet-skill/) / [claude-job-reporter](https://pypi.org/project/claude-job-reporter/0.6.0/)
- [MarkItDownとClaude CodeでExcel方眼紙をリフォーマット（Classmethod）](https://dev.classmethod.jp/articles/markditdown-claude-code-excel-reformat/)
- [xlsx2md（igapyon）](https://github.com/igapyon/xlsx2md) / [ExStruct](https://pypi.org/project/exstruct/0.5.2/) / [Nutrient: Excel to Markdown for LLMs](https://www.nutrient.io/blog/excel-to-markdown-for-llms/)
- [Excel MCP Server（SuperPyonchiX）](https://glama.ai/mcp/servers/@SuperPyonchiX/excel_mcp_server) / [negokaz/excel-mcp-server](https://glama.ai/mcp/servers/@negokaz/excel-mcp-server/tree) / [sbroenne/mcp-server-excel](https://glama.ai/mcp/servers/@sbroenne/mcp-server-excel/blob/39990241b4069cca52e7d8d5f6665f37cffdd9d8/src/ExcelMcp.McpServer/README.md)
- [xlsx template filler（elysiatools）](https://elysiatools.com/en/tools/xlsx-template-filler)
- [SoftBank: 生成AIでExcelを構造化](https://www.softbank.jp/biz/blog/cloud-technology/articles/202512/genai-excel/)

※ 競合が「ない」と言い切れるのは今回検索した範囲まで。公開前に GitHub / npm / PyPI / Chrome ウェブストアで再確認すること。
