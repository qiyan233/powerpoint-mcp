# PowerPoint MCP Server（日本語版）

PowerPoint 自動化のための最高のオープンソース MCP Server。無料、サブスクリプション不要、クレジット制限なし。

このプロジェクトは、毎週 3〜4 時間もプレゼン資料作成に費やすことにうんざりした UCSD の大学院生によって作られました。

**Languages**
- [English](README.md)
- [中文](README.zh-CN.md)
- [日本語](README.ja.md)

[![PyPI version](https://badge.fury.io/py/powerpoint-mcp.svg)](https://badge.fury.io/py/powerpoint-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

---

## デモ動画

[![Watch the PowerPoint MCP Demo](video_thumbnail.jpg)](https://www.youtube.com/watch?v=5p24Vr36py8)

*Claude が “Fourier Transform and Fourier Series” の講義資料をゼロから構築する様子を確認できます。数式、図、アニメーションまで含まれています。*

---

## Windows 専用

**この MCP server は Windows 専用**です。`pywin32` の COM 自動化を使って Microsoft PowerPoint を直接制御しているため、双方向の読み書き、リアルタイム編集、LaTeX レンダリング、テンプレート解析、アニメーション操作を実現できます。

`python-pptx` のような macOS / Linux 系の代替手段では、PowerPoint のフル機能にはアクセスできません。AppleScript ベースの macOS 版 PR は歓迎されています。

---

## このプロジェクトが違う理由

PowerPoint 自動化ツールは数多くありますが、このプロジェクトが本当に価値を持つ理由は次の通りです。

**テンプレート優先設計**  
使いたいテンプレートを LLM に渡すだけで動きます。たとえば Claude に *「会社の Nvidia_Black_Green_2025 テンプレートを使って GPU 性能比較プレゼンを作って」* と頼めば、テンプレートを見つけ、レイアウトを解析し、適切に内容を配置してくれます。

**本当の双方向・リアルタイム編集**  
`python-pptx` のような書き込み専用ツールや、他の一般的な MCP 実装とは異なり、このプロジェクトは COM 自動化を使います。Claude は、いま開いているプレゼンを**読み取り**、PowerPoint を閉じずにそのまま編集できます。

**マルチモーダルなスライド解析**  
`slide_snapshot` は、視覚的コンテキスト（境界ボックス付きスクリーンショット）と詳細なテキスト / グラフ / 表の抽出結果を同時に LLM に渡します。つまり Claude はスライドの内容を本当に“見る”ことができます。

**実用的な LaTeX レンダリング**  
PowerPoint の内蔵数式エディタを操作して LaTeX を描画します。`<latex>E=mc^2</latex>` と書くだけで、PowerPoint ネイティブの数式オブジェクトとして挿入されます。

**トークンを節約する HTML 書式指定**  
`<b>`、`<i>`、`<red>`、`<ul>`、`<ol>` などの HTML タグで太字、斜体、色、箇条書きをまとめて指定できます。何度もツールを呼び出して書式を当てる必要はありません。

**段階表示に対応したアニメーション**  
PowerPoint の実際のアニメーションを制御でき、段落単位の逐次表示にも対応しています。内容を複数の shape に分割しなくても、bullet-by-bullet の表示が可能です。

**1 行でインストール、課金なし**
```bash
claude mcp add powerpoint -- uvx powerpoint-mcp
```
Claude Code、Cursor、GitHub Copilot、その他の MCP クライアントで利用できます。外部サービス不要、月額料金なし、クレジット制限なしです。

---

## 実際のワークフロー

以下は実際に試されたワークフローです。

**調査して作成**
> 「量子コンピューティングの最新動向を調べて、15 枚のプレゼンを作成して」

Claude は web search を使って情報を集め、引用付きで資料を組み立てます。

**データ分析と可視化**
> 「Titanic_dataset.csv を分析して詳細な EDA を行い、結果を説明するプレゼンを作って」

自由に Python / matplotlib を使ったプロットを作成し、そのままスライドの placeholder に挿入できます。

**コードベースのドキュメント化**
> 「このリポジトリ全体を解析して、技術アーキテクチャを説明するプレゼンを作って」

Claude はローカルファイルを読み、構造を理解し、システムを説明するスライドを生成できます。

**テンプレート駆動の社内資料**
> 「Nvidia_Black_Green_template を使って、nvidia_quarterly_sales_data.csv から Q4 売上報告資料を作成して」

テンプレートのレイアウトを自動検出し、placeholder を正しく識別して埋め込みます。

**学術用途 / LaTeX 多用**
> 「Fourier Series と Fourier Transforms を教える 20 枚の講義資料を数式付きで作って」

LaTeX は PowerPoint ネイティブ数式として挿入され、段階的な説明のためのアニメーションも付けられます。

**インタラクティブ学習**
> 「計算生物学の PAM / BLOSUM 行列に関するこのプレゼンを理解したい。各スライドを説明して、各セクションの後にクイズも出して」

Claude はスライド（グラフ、表、スピーカーノートを含む）を読み取り、説明やクイズ生成を行えます。

---

## クイックスタート

**ステップ 1: インストール**
```bash
claude mcp add powerpoint -- uvx powerpoint-mcp
```

**ステップ 2: PowerPoint を開く**
Microsoft PowerPoint を起動します（既存のプレゼンでも、空白の開始画面でも構いません）。

**ステップ 3: Claude に依頼する**
```text
会社のテンプレートを使って、機械学習の基礎を説明する 10 枚のプレゼンを作ってください。
必要な箇所には数式を入れ、各スライドの箇条書きは段落ごとにアニメーション表示してください。
```

これだけで、あとは Claude が処理します。

---

## インストール

### 前提条件
- Windows 10/11
- Microsoft PowerPoint インストール済み
- Python 3.10+

### Claude Code
```bash
# 単一プロジェクト
claude mcp add powerpoint -- uvx powerpoint-mcp

# グローバル利用（1 回インストールすれば全プロジェクトで使える）
claude mcp add powerpoint --scope user -- uvx powerpoint-mcp
```

### Cursor
1. Settings -> Tools & MCP -> New MCP Server をクリック
2. `~/.cursor/mcp.json` を開く
3. 以下を追加：
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
4. IDE を再起動します。

### Trae
1. プロジェクトルートに `.trae/mcp.json` を作成します
2. 以下を追加します：
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
3. 保存後に Trae を再起動し、MCP 対応の Agent モードで利用します

### Trae CN
1. 現在の Trae 中国版でも同じプロジェクト単位の `.trae/mcp.json` 形式を利用できます
2. プロジェクトルートの `.trae/mcp.json` に以下を追加します：
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
3. 保存後に Trae CN を再起動し、MCP 対応の Agent モードを使ってください。ビルドに `Builder with MCP` や `SOLO Coder` がある場合はそれらを優先してください

### Codex
この server は Windows 専用なので、Windows PowerShell では次を実行します：
```bash
codex mcp add powerpoint -- cmd /c uvx powerpoint-mcp
```

追加後は Codex CLI を再起動するか、`/mcp` を実行してサーバーが有効になっていることを確認してください。

### CodeBuddy
1. サイドバーのチャットパネルで CodeBuddy Settings をクリック
2. MCP タブを開く
3. Add MCP をクリック
4. 以下の JSON を追加：
```json
{
  "mcpServers": {
    "powerpoint": {
      "type": "stdio",
      "command": "uvx",
      "args": ["powerpoint-mcp"],
      "description": "PowerPoint automation on Windows"
    }
  }
}
```
5. 保存後、Try to Run で確認するか、Agent から直接呼び出します

### CodeBuddy CN
1. サイドバーのチャットパネル右上から CodeBuddy Settings を開きます
2. MCP タブへ切り替えます
3. Add MCP をクリックします
4. 以下の JSON を追加します：
```json
{
  "mcpServers": {
    "powerpoint": {
      "type": "stdio",
      "command": "uvx",
      "args": ["powerpoint-mcp"],
      "description": "PowerPoint automation on Windows"
    }
  }
}
```
5. 保存後、Try to Run で検証するか、Agent から直接利用します

### OpenCode
1. グローバル設定なら `~/.config/opencode/opencode.json`、プロジェクト単位ならルートの `opencode.json` を開く
2. 以下を追加：
```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "powerpoint": {
      "type": "local",
      "command": ["uvx", "powerpoint-mcp"]
    }
  }
}
```
3. OpenCode を再起動します。

### VS Code (GitHub Copilot)
1. `C:\Users\Your_User_Name\AppData\Roaming\Code\User\mcp.json` を開く
2. 以下を追加：
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
3. IDE を再起動します。

---

## ツール一覧

11 個のツールがあり、それぞれ明確な役割を持っています。

| Tool | 機能 |
|------|------|
| `manage_presentation` | プレゼンのオープン、クローズ、作成、保存を行う。すべてのワークフローの入口。 |
| `slide_snapshot` | スライド解析全般。スクリーンショット、境界ボックス、テキスト抽出、グラフデータ、表解析、リンク、コメント、ノートを取得。 |
| `switch_slide` | アクティブなプレゼンの指定スライドへ移動。 |
| `add_speaker_notes` | 公式 COM 手法でスピーカーノートを追加または置換。 |
| `populate_placeholder` | placeholder にテキスト（HTML 書式対応）、LaTeX 数式、matplotlib プロットを挿入。 |
| `manage_slide` | スライドの複製、削除、移動。 |
| `list_templates` | Personal / User / System テンプレートディレクトリから利用可能なテンプレートを列挙。 |
| `analyze_template` | テンプレートの layout、placeholder 位置、スクリーンショットを解析。 |
| `add_slide_with_layout` | テンプレート内の特定 layout を使って新規スライドを追加。 |
| `add_animation` | shape に入場アニメーションを追加し、段落単位の段階表示にも対応。 |
| `powerpoint_evaluate` | PowerPoint コンテキスト内で任意の Python コードを実行し、複雑な自動化に対応。 |

---

## コンテンツ書式の参考

`populate_placeholder` は、以下の HTML 風タグを受け付けます。

```text
<b>bold text</b>
<i>italic text</i>
<u>underlined text</u>
<red>colored text</red>  (blue, green, orange, purple, white, black も利用可能)
<ul><li>unordered list item</li></ul>
<ol><li>ordered list item</li></ol>
<latex>E = mc^2</latex>
```

matplotlib プロットを挿入する場合は、描画コードをそのまま `content` として渡し、`content_type="plot"` を指定してください。サーバー側がコードを実行して画像として挿入します。

---

## 仕組み

PowerPoint MCP は Windows COM 自動化（`pywin32`）を使って Microsoft PowerPoint を直接操作します。これは VBA マクロや Office アドインと同じ仕組みであり、数式エディタ、アニメーションエンジン、テンプレートシステム、リアルタイム読み取りなど、PowerPoint の完全な機能にアクセスできます。

サーバーは [FastMCP](https://github.com/jlowin/fastmcp) で構築されており、[PyPI](https://pypi.org/project/powerpoint-mcp/) で公開されているため、`uvx` で簡単にインストールできます。

スクリーンショットは自動的に `~/.powerpoint-mcp/` にタイムスタンプ付きファイル名で保存されます。

---

## 制限事項

- **Windows 専用**：COM 自動化は Windows 技術のため、現状 macOS / Linux には対応していません。
- **PowerPoint 必須**：LibreOffice はサポートしていません。
- **Python 3.10+ 必須**：サーバー実行に必要です。

---

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=ayushmaniar/powerpoint-mcp&type=date&legend=top-left)](https://www.star-history.com/#ayushmaniar/powerpoint-mcp&type=date&legend=top-left)

---

MIT License の完全なオープンソースです。役に立ったら GitHub で star を付けてもらえると嬉しいです。

**GitHub:** https://github.com/Ayushmaniar/powerpoint-mcp
