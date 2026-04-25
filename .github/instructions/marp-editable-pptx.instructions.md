---
name: 'Marp Editable PPTX 開発規約'
description: 'marp-to-editable-pptx の native-pptx モジュール開発規約。dom-walker.ts ・ slide-builder.ts の修正・テスト追加・ fixture 管理・ ADR 記録・ commit 規約・デグレ防止を扱うときに適用する。'
applyTo: 'src/native-pptx/**/*.ts, src/native-pptx/test-fixtures/**, src/native-pptx/README.md'
---

# Marp Editable PPTX 開発規約

## 言語ポリシー

`src/native-pptx/` 配下のソースコード・コメント・テストケース名・ドキュメントは **すべて英語で書く**。
日本語は ADR ログ（`src/native-pptx/README.md` の「バグ修正・意思決定の記録」セクション）にのみ使用する。

## 設計原則（これに反する実装は禁止）

**ブラウザが唯一の真実（Browser is the source of truth）**

- `getComputedStyle()` と `getBoundingClientRect()` の値を 1:1 で PPTX に写す
- Marp テーマ・CSS セレクタ・Markdown 構文を解析しない
- 要素固有のハードコード対応は、**ブラウザが既に描画済みだが PPTX 側の制限で再現できない場合のみ**許容する（例: SVG `<foreignObject>`, スライドページ番号）
- その場合の修正方法は「ブラウザのレンダリング結果をラスタ画像としてキャプチャする」のみ

## アーキテクチャ

| ファイル | 役割 | 修正するケース |
|---|---|---|
| `dom-walker.ts` | ブラウザ DOM から `SlideData[]` を抽出 | テキストが消える・抽出されない・余計な要素が混入する |
| `slide-builder.ts` | `SlideData[]` を PptxGenJS API 呼び出しに変換 | 座標変換ミス・PPTX 出力形式の問題・色変換の誤り |
| `index.ts` | パイプライン全体の制御 | 画像ラスタライズ・ブラウザライフサイクル |
| `utils.ts` | 変換ユーティリティ（px→inch, rgb→hex 等） | 単位変換の誤り |

> **注意**: `dom-walker.ts` はブラウザ内で `page.evaluate()` で実行されるため、webpack/esbuild のスコープ外。
> 変更後は必ず `node src/native-pptx/scripts/generate-dom-walker-script.js` で再コンパイルする。

## ビルドシーケンス

```powershell
# dom-walker.ts を変更した場合（必須）
node src/native-pptx/scripts/generate-dom-walker-script.js

# dom-walker.ts または index.ts を変更した後、gen-pptx.js（ローカルツール）を最新化する場合（必須）
node src/native-pptx/scripts/build-native-pptx-bundle.js
```

> **`npm run build` はこれらを実行しない**（VS Code 拡張の webpack bundle のみ生成する）。
> `dom-walker.ts` を変更したら必ず `generate-dom-walker-script.js` を実行すること。
> `build-native-pptx-bundle.js` は `gen-pptx.js` が参照する `lib/native-pptx.cjs` を生成するため、ローカルでの視覚比較をする前に必要になる。

## fixture 管理規約

### 機密・個人データの排除（公開リポジトリ）

`src/native-pptx/test-fixtures/pptx-export.md` は公開リポジトリにコミットされる。
以下を絶対に含めない：

- 開発者のローカルパス（`C:\Users\...`、`/home/...`）
- 顧客名・プロジェクト名・社内システム名・業務データ
- 社内 URL・IP アドレス・認証情報

再現スライドは必ず汎化する（`Sample Title`、`Alice` / `Bob`、`path/to/file.md` 等）。

> バグの本質は Marp テーマの CSS/DOM 構造にある。テキスト内容を変えても同一の条件で再現できる。再現しない場合はテキストパターン（特殊文字・長さ・禁則処理等）が原因なので最小再現テキストを使う。

### fixture を追加する際の手順

1. 問題のスライドだけで単独再現するか確認する（単独 deck で `gen-pptx.js` を実行）
2. `<style>` を追加する場合は `section` セレクタ等でスコープを絞る
3. fixture に追加後、全スライドの `compare-visuals.js` を実行して既存スライドが壊れないことを確認する

### README の枚数記載を必ず更新する（2箇所）

fixture にスライドを追加したら次の 2 箇所を必ず同時に更新する：

| ファイル | 更新箇所 |
|---|---|
| `README.md`（リポジトリルート） | `compare-NNN.png` の行と `All slide comparisons (N slides)` の枚数 |
| `src/native-pptx/README.md` | 「Canonical test deck」セクションの枚数記述と「Visual diff improvement loop」セクション |

> スライドを追加したのに README を更新しないことが繰り返し発生している。
> スライド追加のコミットには必ずこの 2 箇所の更新を含める。

## ADR 記録（修正のたびに必須）

`src/native-pptx/README.md` の「バグ修正・意思決定の記録」セクションに追記する。
**ADR を読まずに修正に入ると、過去に解決した問題を再発させる（実際に繰り返し発生している）。**

必須項目：
- 問題（症状）
- 根本原因（DOM 処理・CSS 解釈・座標計算の観点で）
- 修正（どのファイル・関数・ロジックを変えたか）
- テスト追加（追加した test case 名）
- なぜ単体テストや画像 diff で検知できなかったか

## テスト規約

- テストケース名は **英語**（`src/native-pptx/README.md` の言語ポリシーより）
- バグ修正時は必ず `dom-walker.test.ts` または `slide-builder.test.ts` に回帰テストを追加する
- テストケース名から「何を検証しているか」が分かる形にする
- `describe` ブロックは対象関数名をそのまま使う

## デグレ防止の 2 軸

修正後は必ず以下の **両方** を確認する：

| 軸 | 確認内容 |
|---|---|
| ① ルールベース単体テスト | `npx jest` が全件パスするか。過去に追加した回帰テストが壊れていないか |
| ② ビジュアル diff 傾向 | `compare-report.html` の **差分の種類** を目視確認。特に改行ズレ・重なり・欠落を確認する |

### 差分率だけで OK/NG を判断しない

- 折り返しによる行ズレは差分率がほぼ 0% のまま発生することがある
- ページはみ出しも差分率では検知できない
- 目視確認では「テキストの行数が HTML と一致しているか」を必ず確認する

## commit 規約

Conventional Commits を使う：

```
fix(<scope>): 説明
feat(<scope>): 説明
docs(<scope>): 説明
chore(<scope>): 説明
ci(<scope>): 説明
```

- scope は対象ファイル名（例: `dom-walker`, `slide-builder`, `compare-visuals`）
- 1 コミットは 1 つの問題の修正
- `dist/` 内のファイルをコミットしない
- `slides-ci.html` をコミットしない
- commit 対象は `.ts` / `.test.ts` / `pptx-export.md` / `README.md` の変更のみ

## ブランチ・PR 規約

- ブランチ名: `fix/説明-in-kebab-case`、`feat/説明-in-kebab-case`
- PR を経由して main にマージする（直プッシュしない）
- PR のタイトルは commit メッセージと同じ形式にする
- リリースするには PR に `release` ラベルを付ける

## やってはいけないこと

- `dist/` に出力されるファイルを git add する
- `slides-ci.html` を git add する
- 修正と無関係なファイルを変更する
- `npm run build` で bundle が更新されたと思い込む（`dom-walker.ts` 変更後は必ず再コンパイルする）
- LibreOffice をローカルにインストールする（PowerPoint COM で代替する）
- ブラウザの CSS レンダリング結果を上書きするような element-specific 処理を書く（設計原則違反）
- ADR を読まずに修正に入る
- 新しいツールや補助スクリプトを依頼なく作成する（`compare-visuals.js` / `gen-pptx.js` / `diagnose-pptx.js` で事足りる）
