---
name: marp-pptx-visual-diff
description: 'marp-to-editable-pptx の visual fidelity 改善ループを Windows 上で進めるスキル。LibreOffice 不要。PowerPoint COM で PPTX→PNG 変換。dom-walker.ts か slide-builder.ts の修正、bundle 再生成、compare-visuals によるスライド比較、ADR 記録まで一連で扱う。"PPTX の見た目がおかしい" "スライドがずれている" "テキストが消えている" "compare-visuals を回したい" "visual diff ループ" といった依頼で必ず使う。LibreOffice をインストールしてはならない。'
argument-hint: '対象スライド番号 または 症状の説明'
---

# marp-pptx Visual Diff 改善ループ（Windows / LibreOffice 不要）

## この環境の大前提

- **Windows 専用ワークフロー**。PPTX → PNG の変換は PowerPoint COM を使う。
- **LibreOffice をインストールしてはならない**。このリポジトリの存在理由は "LibreOffice 不要のエディタブル PPTX" であり、環境もそれを前提とする。
- CI は Ubuntu + LibreOffice で動く。ローカル確認は PowerPoint COM で代替し、CI と完全同一の数値を求めない。CI 側の `compare-061.png` 等は GitHub Actions 手動実行で取得する。

---

## ループ全体の流れ

```
症状確認
  │
  ├─ 単体テストが落ちる → dom-walker.test.ts / slide-builder.test.ts で最小再現 → 修正 → テスト → bundle 再生成 → compare
  │
  └─ 見た目がおかしいが単体テストは通る → fixture に再現スライド追加 → compare で確認 → 修正 → ループ
```

---

## Step 1: 再現スライドを fixture に追加

**新しいバグを発見したら必ず先に fixture を足す。**

```
src/native-pptx/test-fixtures/pptx-export.md
```

- 末尾の既存スライドの後に `---` で区切って追加する。
- スライドタイトルに番号とバグ内容を入れる（例: `# Slide 62: ...`）。
- `README.md` (repo ルート) の `compare-NNN.png` 行と枚数も同時に更新する（過去の失敗教訓）。
- `src/native-pptx/README.md` のスライド枚数記載も更新する。

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

---

## Step 6: 修正の方針

| 問題の種類 | 直す場所 |
|---|---|
| テキストが消える / ずれる | `src/native-pptx/dom-walker.ts` |
| 座標変換の誤り | `src/native-pptx/dom-walker.ts` または `src/native-pptx/slide-builder.ts` |
| PPTX 出力形式の問題 | `src/native-pptx/slide-builder.ts` |

修正後は必ず：

1. `dom-walker.test.ts` または `slide-builder.test.ts` に回帰テスト追加
2. Step 2 → Step 4 → Step 5 を再実行して diff が改善したことを確認
3. ADR を `src/native-pptx/README.md` の ADR ログに追記

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
- [ ] commit 対象: `.ts` / `.test.ts` / `pptx-export.md` / README の変更のみ
- [ ] `dist/` の生成物は commit しない
- [ ] `slides-ci.html` は commit しない
- [ ] ADR 追記済み

---

## やってはいけないこと

- LibreOffice を winget / msiexec / 管理者権限でインストールしようとする
- `npm run build` だけ実行して bundle が更新済みと思い込む
- `dist/` 内のファイルを git add する
- pptx-export.md への再現スライド追加を飛ばして直接修正に入る
- worktree を作ったまま放置する（使い終わったら `git worktree remove` と `git worktree prune`）
