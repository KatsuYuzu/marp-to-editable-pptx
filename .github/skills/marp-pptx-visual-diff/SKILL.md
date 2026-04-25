---
name: marp-pptx-visual-diff
description: 'marp-to-editable-pptx の visual fidelity 改善ループを Windows 上で進めるスキル。LibreOffice 不要。PowerPoint COM で PPTX→PNG 変換。設計原則（ブラウザ唯一の真実）・言語ポリシー（英語）・修正場所判断（dom-walker vs slide-builder）・fixture 機密データ排除・README 2箇所更新・bundle 再生成・compare-visuals 目視判定・ADR 記録まで一連で扱う。"PPTX の見た目がおかしい" "スライドがずれている" "テキストが消えている" "compare-visuals を回したい" "visual diff ループ" "改行によるズレ" "テキストの折り返しがずれる" "差分率だけでは判断できない" といった依頼で必ず使う。LibreOffice をインストールしてはならない。fixture に機密・個人データを含めてはならない。'
argument-hint: '症状の説明（例: "Slide 34 でバッジが浮いている"）または 対象スライド番号'
---

# marp-pptx Visual Diff 改善ループ（Windows / LibreOffice 不要）

## この環境の大前提

- **Windows 専用ワークフロー**。PPTX → PNG の変換は PowerPoint COM を使う。
- **LibreOffice をインストールしてはならない**。このリポジトリの存在理由は "LibreOffice 不要のエディタブル PPTX" であり、環境もそれを前提とする。
- CI は Ubuntu + LibreOffice で動く。ローカル確認は PowerPoint COM で代替し、CI と完全同一の数値を求めない。CI 側の `compare-061.png` 等は GitHub Actions 手動実行で取得する。

## 設計原則（修正に入る前に必ず確認）

**ブラウザが唯一の真実（Browser is the source of truth）**

- `getComputedStyle()` / `getBoundingClientRect()` の値を 1:1 で PPTX に写す
- テーマ・CSS セレクタ・Markdown 構文を解析してはならない
- 要素固有のハードコード対応は「ブラウザが既に描画済みだが PPTX 側の制限で再現できない場合のみ」許容される
  - 許容例: SVG `<foreignObject>`（PowerPoint が描画不可）、スライドページ番号（再番号付けが必要）
  - その場合の修正方法は「ブラウザのレンダリング結果を PNG キャプチャする」のみ
- この原則に違反する修正（CSS を解釈するコード・要素専用の分岐追加）は設計として誤り

## 修正場所の判断（dom-walker vs slide-builder）

| 症状 | 修正場所 |
|---|---|
| テキストが抽出されない・消える・余計な要素が混入する | `dom-walker.ts` |
| **テキストが 2個表示される（重複）**—同じテキストが PPTX に 2件以上指定されている | `dom-walker.ts`（レンダリング前のテキストノードを誤収集している可能性が高い） |
| 座標変換ミス・幅/高さの計算誤り | `dom-walker.ts` または `slide-builder.ts` |
| PPTX 出力形式の問題（マージン・色・フォント） | `slide-builder.ts` |
| 画像ラスタライズの条件漏れ | `index.ts` |

> `dom-walker.ts` はブラウザ内で実行される。変更後は必ず `generate-dom-walker-script.js` を再実行すること。

## 言語ポリシー

`src/native-pptx/` 配下のソースコード・コメント・テストケース名は **すべて英語**。
日本語は ADR ログ（`src/native-pptx/README.md` の「バグ修正・意思決定の記録」セクション）にのみ使用する。

---

## ループ全体の流れ

```
症状確認
  │
  ├─ まず `npx jest` を実行して現状を把握する
  │    │
  │    ├─ テストが落ちる → dom-walker.test.ts / slide-builder.test.ts で最小再現 → 修正 → テスト → bundle 再生成 → compare
  │    │
  │    └─ テストは通る（見た目だけおかしい）
  │         │
  │         ├─ Step 1: fixture に再現スライド追加（機密データ排除・README 2箇所更新）
  │         ├─ Step 2: 初回 compare は既存 bundle のまま実行。修正後（dom-walker.ts 変更後）に再生成
  │         ├─ Step 3: HTML 生成（--html --allow-local-files 必須）
  │         ├─ Step 4: PPTX 生成（gen-pptx.js）
  │         ├─ Step 5: compare-visuals で比較
  │         ├─ Step 5b: diff の「種類」を目視判定（差分率だけで判断しない・改行ズレに注意）
  │         ├─ Step 5c: ADR 確認 → 修正 → 2軸デグレチェック（テスト＋目視）
  │         ├─ Step 6: 修正（dom-walker.ts or slide-builder.ts を英語テスト付きで修正）
  │         └─ Step 7: ADR 記録
  │
  └─ compare は PASS だがユーザー報告あり
       │
       ├─ Step 5b で目視確認 → 問題なし → ユーザーに PPTX の実画面スクリーンショット提供を依頼
       └─ Step 5b で目視 NG → Step 5c（ADR 確認）→ Step 1（fixture 追加）→ Step 3 以降の通常ループへ
```

---

## Step 1: 再現スライドを fixture に追加

**新しいバグを発見したら必ず先に fixture を足す。既存スライドでバグが発覚した場合も、バグの再現に特化した最小再現スライドを末尾に追加する（既存スライドを直接修正しない）。**

```
src/native-pptx/test-fixtures/pptx-export.md
```

- 末尾の既存スライドの後に `---` で区切って追加する。
- スライドタイトルに番号とバグ内容を入れる（例: `# Slide 62: ...`）。

### ⚠️ README 2箇所を必ず同時に更新する（繰り返し発生した失敗）

スライドを追加したのに README 更新を忘れることが繰り返し発生している。以下の 2 箇所を同じコミットで必ず更新する：

| ファイル | 更新箇所 |
|---|---|
| `README.md`（リポジトリルート） | `<details>` 内の `compare-NNN.png` の行追加 と `All slide comparisons (N slides)` の枚数 |
| `src/native-pptx/README.md` | 「Canonical test deck」セクションの枚数（例: `63 slides`）と fixture ファイルの説明 |

CI の `screenshots.yml` が GitHub Pages の比較画像を自動更新するため、`compare-NNN.png` の `<img>` タグだけ追加しておけば画像自体は CI が生成する。

### fixture に取り込む際の必須ルール

#### 機密・個人データの排除（公開リポジトリ）

このファイルは公開リポジトリにコミットされる。**以下を絶対に含めない：**

- 開発者のローカルパス（`C:\Users\...`、`/home/...` 等）
- 顧客名・プロジェクト名・社内システム名
- 業務データ・実データ（ファイル名、金額、氏名、ID 等）
- 社内 URL・IP アドレス・認証情報

**再現スライドは十分に汎化する：**

| 元のデータ | fixture に書く内容 |
|---|---|
| `C:\Users\tanaka\project\slides.md` | `path/to/slides.md` |
| `株式会社〇〇 売上データ 2025` | `Sample Title` |
| 顧客名・担当者名 | `Alice` / `Bob` / `Item A` |
| 実際の業務フロー図 | 同じ CSS/レイアウト構造を持つ汎用ダイアグラム |
> **短い汎用英語単語（`input`、`data`、`item`、`label` 等）はそのまま使用してよい。汎化不要。**
不具合の本質は「CSS のレイアウト・DOM 構造」にある。テキスト内容を変えても再現するはず。再現しない場合はテキストパターン（特殊文字・長さ・禁則処理等）が原因なので、最小再現テキストを使う。

#### 範囲の特定（スライド個別 vs グローバル CSS）

fixture を追加する **前に** 以下を確認する：

1. **問題のスライドだけを単独の Markdown（1 枚 deck）にしてもバグが再現するか確認する**
   ```powershell
   # dist/repro-single.md を作り単独ビルドで確認
   npx marp dist/repro-single.md --html --allow-local-files --output dist/repro-single.html
   node src/native-pptx/tools/gen-pptx.js dist/repro-single.html dist/repro-single.pptx
   ```
   - 再現する → スライド個別の問題 → その CSS/DOM 構造だけ fixture に追加
   - 再現しない → 他のスライドの `<style>` やグローバル CSS が干渉している → 干渉元を特定してから fixture を設計する

2. **`<style>` をスライドページ先頭にスコープする**  
   テーマや共通 CSS を変更するような `<style>` ブロックを fixture に追加すると、後続の全スライドの見た目が変わる。
   - `<style>` を追加するときは必ず `section` セレクタ等でスコープを絞る
   - または Marp の `<!-- _class: xxx -->` を使って当該スライドにのみ適用する

3. **fixture 追加後に全スライドの compare を回して既存スライドが壊れていないことを確認する**
   ```powershell
   node src/native-pptx/tools/compare-visuals.js `
     src/native-pptx/test-fixtures/slides-ci.html `
     dist/compare-out.pptx
   # → compare-report.html で新規 FAIL が発生していないか全件確認
   ```
   新規 FAIL が出た場合はグローバル影響を疑い、追加した `<style>` や HTML 構造を見直す。

---

## Step 2: 必要なビルド

> **初回の症状確認比較（Step 3〜5）は既存 bundle のまま実行してよい。**
> bundle の再生成が必要になるのは「`dom-walker.ts` を修正した後」のみ。

```powershell
# DOM walker を変更したとき（dom-walker.ts）
node src/native-pptx/scripts/generate-dom-walker-script.js

# gen-pptx.js で使う standalone bundle を再生成するとき
# ← npm run build ではこれは再生成されない（よくあるトラップ）
node src/native-pptx/scripts/build-native-pptx-bundle.js
```

> **注意**: `npm run build` は VS Code 拡張の webpack バンドルを生成するが、
> `lib/native-pptx.cjs`（`gen-pptx.js` が依存する bundle）は生成しない。
> `dom-walker.ts` を変更したら必ず上の 2 コマンドを追加で実行する。

---

## Step 3: HTML 生成

```powershell
# --html と --allow-local-files は必須（省くとバッジ・mermaid・画像が欠ける）
npx marp src/native-pptx/test-fixtures/pptx-export.md `
  --html --allow-local-files `
  --output src/native-pptx/test-fixtures/slides-ci.html
```

`slides-ci.html` は `.gitignore` 対象。コミットしない。

---

## Step 4: PPTX 生成

```powershell
node src/native-pptx/tools/gen-pptx.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/compare-out.pptx
```

問題の原因を絞るときは 1 枚 deck で先に確認する：

```powershell
# 問題スライドだけ抜き出した markdown を一時作成して確認
npx marp dist/slide-repro.md --html --allow-local-files --output dist/slide-repro.html
node src/native-pptx/tools/gen-pptx.js dist/slide-repro.html dist/slide-repro.pptx
```

#### デバッグ出力（DOM 抽出 JSON）

```powershell
$env:MARP_PPTX_DEBUG = '1'
node src/native-pptx/tools/gen-pptx.js dist/slide-repro.html dist/slide-repro.pptx
# → dist/slide-repro.native-pptx.json に SlideData[] をダンプ
```

JSON で座標・テキストが正しく取れているか確認してから describe を書くと速い。

---

## Step 5: compare-visuals で比較

```powershell
# PPTX → PNG は PowerPoint COM（要インストール済み PowerPoint）
node src/native-pptx/tools/compare-visuals.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/compare-out.pptx
```

出力は必ず `dist/compare-slides-ci/` に落ちる（`src/` には絶対書かない）：

| ファイル | 内容 |
|---|---|
| `html-slide-NNN.png` | Marp HTML の参照スクリーンショット |
| `pptx-slide-NNN.png` | PowerPoint COM 出力 |
| `diff-slide-NNN.png` | pixelmatch 差分 |
| `compare-report.html` | per-slide 差分率サマリ |

`compare-report.html` を開いて FAIL / WARN のスライドを確認する。

---

## Step 5b: diff の「種類」を判定する（差分率だけで OK/NG を判断してはならない）

**差分率（RMSE・ピクセル差分%）はあくまで参考値。** プレゼンスライドとして実際に問題になるかどうかは、「何がどうずれているか」で判断する。

### 受け入れ可能な差分（修正不要）

| 種類 | 見分け方 |
|---|---|
| フォントのアンチエイリアス差 | diff-NNN.png で文字の輪郭だけが赤くなっている。レイアウトは一致している |
| サブピクセルレベルの位置ずれ | 1〜2px の均一なにじみ。要素全体が正しい位置に収まっている |
| OS/ブラウザ間のカーニング差 | 文字間隔が微妙に違うが行内に収まっている |
| 背景グラデーションの微差 | 差分が一様に薄く広がっている。図形や文字の境界ではない |

### 修正必須の差分（NG）

| 種類 | 見分け方 | 代表的な根本原因 |
|---|---|---|
| **レイアウト位置ずれ** | 要素がスライド内で横/縦にまるごとずれている。diff に「帯状の赤」が出る | 座標計算ミス、padding/margin の未考慮 |
| **図形・テキストの重なり** | 本来重ならないはずの要素が重なっている | z-order、座標オフセットの二重計上 |
| **折り返し・行数の不一致** | PPTX 側のテキストが行あふれ・折り返し増減している。diff で行末に縦の赤線 | テキストボックス幅/高さ不足、フォントサイズ変換ミス |
| **余分なテキスト要素の混入** | PPTX にはあるが HTML にはない文字列が重なっている（raw ソースコード等） | DOM walker がレンダリング前テキストノードを誤収集 |
| **図形・画像の欠落** | HTML 側にある要素が PPTX で消えている。diff にベタの赤いブロック | extract ロジックのスキップ条件の誤り |
| **図形の色・塗りのずれ** | 背景色や枠線色が大きく違う。diff で広い面積に強い赤 | backgroundColorの取得ミス、透明度の扱い |

### 判定の手順

1. `compare-report.html` で差分率の高いスライドを列挙する
2. 各スライドの `html-slide-NNN.png` と `pptx-slide-NNN.png` を **横並びで目視確認**する
3. `diff-slide-NNN.png` で「どこで差が出ているか」のパターンを分類する
4. **差分率が低い（FAILでない）スライドも1枚ずつ目視する**
   - 差分率が低くても「NG」に該当する問題は紛れている場合がある
   - 特にテキスト重なり・図形欠落は差分率に出ないことがある
5. NG に分類した問題を修正課題として Issue / todo に記録してから修正に入る

### 「差分率が低いからOK」は禁止

- RMSE は差分の「量」を示すが「種類」を示さない
- 図形が1つ完全に消えても、他のほとんどが一致していれば差分率は低く出る
- 逆にフォントレンダリングの差だけで FAIL 判定になることもある
- **全スライドを目視してから「このスライドは合格」と判断する**こと

> 目視で確認しても問題を特定できない場合は、ユーザーに「PPTX を PowerPoint で開いたときのスクリーンショット」を提供してもらい、PPTX の実表示と HTML が差异する箇所を特定する。

### ⚠️ 改行・折り返しによるズレは「差分率に出ない」ことがある

これは繰り返し見落とされてきた最重要クリティカル。

- テキストの折り返しが1行増減するだけで、以降の要素が全体的に縦にずれる
- 折り返しズレは **差分率がほぼ0%** のまま発生することがある（周辺ピクセルが同色なら差分が出ない）
- ページはみ出しも同様に差分率では検知できない
- 目視確認のときは「行数が HTML と一致しているか」「テキストが枠からはみ出していないか」を明示的に確認する

**目視チェックリスト（diff率に関わらず必ず確認）:**
- [ ] 各スライドのテキストの行数が HTML と一致しているか
- [ ] テキストボックスからのはみ出しがないか
- [ ] 箇条書きの続き行が次の要素と重なっていないか
- [ ] 絵文字・バッジ等インライン要素が同一行に留まっているか

---

## Step 5c: デグレ防止の2軸チェック

**修正前に ADR を確認し、修正後に2軸でデグレがないことを確認する。**

| 軸 | 何を確認するか |
|---|---|
| ① ルールベース単体テスト | `dom-walker.test.ts` / `slide-builder.test.ts` が全件パスするか。過去に追加した回帰テストが壊れていないか |
| ② ビジュアル diff 傾向 | `compare-report.html` で **差分率ではなく差分の種類** を確認。特に改行ズレ・重なり・欠落を目視で確認する |

### 修正前の必須確認

1. `src/native-pptx/README.md` の ADR ログを読んで、過去の意思決定と修正済みケースを把握する
2. 今回の修正が既存の ADR 判断と矛盾しないか確認する（矛盾する場合はリバートではなく新 ADR で上書きする）

> **ADR を読まずに修正に入るとデグレが繰り返される。** 過去に直した問題が再発した場合、その修正が ADR に記録されていなかったことが原因であることが多い。

---

## Step 6: 修正の方針

| 問題の種類 | 直す場所 |
|---|---|
| テキストが消える / ずれる / 余計な要素が混入する | `src/native-pptx/dom-walker.ts` |
| 座標変換・幅/高さ計算の誤り | `src/native-pptx/dom-walker.ts` または `src/native-pptx/slide-builder.ts` |
| PPTX 出力形式の問題（色・マージン・フォント） | `src/native-pptx/slide-builder.ts` |
| 画像ラスタライズの条件漏れ | `src/native-pptx/index.ts` |

### 修正後の手順

1. `dom-walker.test.ts` または `slide-builder.test.ts` に **英語で** 回帰テストを追加する
2. `dom-walker.ts` を変更した場合: `node src/native-pptx/scripts/generate-dom-walker-script.js` を実行する
3. Step 2 → Step 4 → Step 5 を再実行して diff が改善したことを確認する
4. ADR を `src/native-pptx/README.md` の ADR ログに追記する

### ADR に必ず書くこと

- 問題（症状）と根本原因（DOM処理・CSS解釈・座標計算の観点で）
- 修正（どのファイル・関数・ロジックを変えたか）
- テスト追加（追加した test case 名）
- **なぜ単体テストや画像 diff で検知できなかったか**（次の同種バグの早期発見に使う）

---

## Step 7: ADR 記録

`src/native-pptx/README.md` の末尾に追記する。必須項目：

```markdown
### ADR-N: 現象タイトル

**問題**
症状の簡潔な説明

**根本原因**
なぜそうなったか（DOM 処理・CSS 解釈・座標計算の観点で）

**修正**
どのファイル・関数・ロジックを変えたか

**テスト追加**
追加した test case 名
```

---

## 完了条件

- [ ] `npx jest` 全テスト通過
- [ ] `compare-report.html` に FAIL なし（修正前より diff 率が下がっていること）
- [ ] 全スライドを目視確認し、NG 差分がないことを確認した
- [ ] commit 対象: `.ts` / `.test.ts` / `pptx-export.md` / README の変更のみ
- [ ] `dist/` の生成物は commit しない
- [ ] `slides-ci.html` は commit しない
- [ ] ADR 追記済み

---

## 改善ループ完了後の報告フォーマット

改善ループが完了したら、必ず以下の形式で報告する。

```
## 改善ループ完了レポート

### 変更内容
- 修正ファイル: （例: src/native-pptx/dom-walker.ts）
- 変更概要: （1〜2行で）

### テスト結果
- 単体テスト: N 件パス（新規追加 N 件）

### 比較レポート（ローカル確認用）
比較レポート: dist\compare-slides-ci\compare-report.html
（ブラウザで開くと全スライドの HTML / PPTX 横並び比較と差分率が確認できます）

### 目視確認結果
- 対象スライド: N 枚
- FAIL（閾値超過）: N 枚
- 目視NG（テキスト重なり・欠落・レイアウトずれ等）: N 件
  - Slide NNN: （問題の説明）

### コミット
ブランチ: fix/...
コミット: （ハッシュ）
```

> **HTML パスではなくファイルシステムパスを報告する。**
> ユーザーがエクスプローラや `start` コマンドで直接開けるよう、
> `dist\compare-slides-ci\compare-report.html` 形式（バックスラッシュ）で記載する。

---

## やってはいけないこと

- LibreOffice を winget / msiexec / 管理者権限でインストールしようとする
- `npm run build` だけ実行して bundle が更新済みと思い込む
- `dist/` 内のファイルを git add する
- pptx-export.md への再現スライド追加を飛ばして直接修正に入る
- worktree を作ったまま放置する（使い終わったら `git worktree remove` と `git worktree prune`）
- **差分率が低いから OK と判断する**（特に改行ズレ・折り返しは差分率に出ない）
- **ADR を読まずに修正に入る**（過去の意思決定を無視した修正はデグレを招く）
- **開発者のローカルパス・業務データ・機密データを pptx-export.md に直接書き込む**（公開リポジトリ。必ず汎化する）
- **修正と無関係なファイルを追加・変更する**（commit 対象は `.ts` / `.test.ts` / `pptx-export.md` / README の変更のみ）
- **比較ツールや補助スクリプトを新規作成する**（既存の `compare-visuals.js` / `gen-pptx.js` / `diagnose-pptx.js` で事足りる。新規作成を求められていない限り作らない）
