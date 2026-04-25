---
marp: true
paginate: true
html: true
theme: default
---

<!--
  PPTX Export Comprehensive Test Slides
  ===========================================
  This file is a test MD that covers all known edge cases and expected issues.

  Generation steps:
    marp --html src/native-pptx/test-fixtures/pptx-export.md -o out/test-pptx-export.html
    node scripts/gen-pptx.js out/test-pptx-export.html out/test-pptx-export.pptx
    node scripts/compare-visuals.js out/test-pptx-export.html out/test-pptx-export.pptx
-->

<!-- _class: lead -->

# PPTX Export<br>Comprehensive Test Slides

## All Known Bugs & Edge Cases — Full Coverage

---

# Slide 2: Basic headings and paragraphs

## H2 Heading

### H3 Heading (blue)

#### H4 Heading

This is a normal paragraph text. Verify that writing a long line without breaks does not cause errors.

<!-- Japanese text intentionally kept for CJK rendering test: mixed Japanese/English -->

日本語と English が混在していても問題ないことを確認します。

---

# Slide 3: Inline decorations

**Bold text** and _italic text_ and `inline code` and ~~strikethrough~~ test.

**Bold** followed by _italic_, and **bold with _nested italic_ inside** is also tested.

[Hyperlink](https://github.com/marp-team/marp-vscode) placed on the same line as text.

Multiple `code` spans `included` in a sentence.

---

# Slide 4: Emoji test ← Important

**Emoji at the beginning of text:**

✅ Sentence starting with a checkmark  
🚀 Sentence starting with a rocket  
⚠️ Sentence starting with a warning sign

**Emoji in the middle of text:**

Alpha zeta ✅ beta gamma delta  
Zeta 🚀 eta theta. Iota kappa  
Lambda ⚠️ mu nu xi omicron  
Pi ⭐ rho sigma tau

**Emoji at the end of text:**

Alpha complete 🎉  
Beta gamma passed ✅

**Consecutive emoji:**

🎉🚀✅⚠️👍 Five consecutive emoji in a row

---

# Slide 5: Emoji + Bold + Italic compound

Alpha beta — gamma delta **✅ Zeta item**.

**🚀 Eta** through _⚠️ Theta check_ — iota kappa.

Emoji in lists:

- ✅ Lambda item
- 🔄 Mu item
- ⏳ Nu item
- ❌ Xi item

---

# Slide 6: Paragraph + Code + Blockquote

```typescript
function greet(name: string): string {
  return `Hello, ${name}!`
}
```

> This is a blockquote. It also tests
> the case of a long quote with line breaks.

---

# Slide 7: Table

| Column 1 (left-aligned)     | Column 2 (center-aligned) | Column 3 (right-aligned) |
| :-------------------------- | :-----------------------: | -----------------------: |
| Cell A1                     |          Cell A2          |                  Cell A3 |
| **Bold** cell               |        `Code` cell        |            _Italic_ cell |
| Cell with long text content |          Medium           |                    Short |
| Last row                    |        Last row 2         |               Last row 3 |

---

# Slide 8: Bullet lists

- Top-level item 1
  - Level 2 item 1-1
  - Level 2 item 1-2
    - Level 3 item 1-2-1
    - Level 3 item 1-2-2
  - Level 2 item 1-3
- Top-level item 2
- Top-level item 3 (**bold** included)

1. Numbered list item 1
2. Numbered list item 2
   1. Numbered nested 2-1
   2. Numbered nested 2-2
3. Numbered list item 3

---

# Slide 9: Code block (syntax highlighting)

```python
# Python sample
import asyncio

async def fetch_data(url: str) -> dict:
    """Fetch data asynchronously"""
    async with aiohttp.ClientSession() as session:
        async with session.get(url) as response:
            return await response.json()
```

```sql
SELECT u.name, COUNT(o.id) AS order_count
FROM users u
LEFT JOIN orders o ON u.id = o.user_id
WHERE u.created_at >= '2024-01-01'
GROUP BY u.id
ORDER BY order_count DESC;
```

---

# Slide 10: Image (same directory, no filter)

![w:300 Test icon](./test-icon.png) Normal image beside text

Image placed below:

![w:400](./test-icon.png)

---

# Slide 11: Image (CSS filter)

Verify the following are processed correctly compared to a normal image:

![w:300 Grayscale](./test-icon.png)
No filter (left)

![w:300 filter:grayscale(1)](./test-icon.png)
Grayscale (right)

---

<!-- _backgroundImage: url('./test-icon.png') -->

# Slide 12: Background image (full slide)

This background image has no filter. Verify that text is not duplicated.

Body text. Verify it is correctly placed on top of the background.

---

<!-- _backgroundImage: url('./test-icon.png') -->
<!-- _backgroundColor: rgba(0,0,0,0.4) -->

# Slide 13: Background image + Grayscale filter

<style scoped>
section {
  color: white;
}
</style>

This background should have a grayscale filter applied.  
**Important**: Verify that text is NOT displayed twice.  
The background should be gray, and this text should appear only once.

---

# Slide 14: Inline-only div (important edge case)

<div style="border: 2px solid #336; border-radius: 8px; padding: 16px; background: #f0f4ff;">
This container provides <strong>val-7 items</strong> ready to use right away.<br>
They cover the full set: alpha, beta, gamma, delta, epsilon, zeta, eta.
</div>

The div above is an "inline-only container". Text should be displayed correctly inside the frame.

---

# Slide 15: Card component (border + background)

<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
<div style="border: 1px solid #ccc; border-radius: 12px; padding: 16px; background: #f8f9ff;">

### 🎯 Card A

A mix of text, **bold**, and `code`.  
Emoji are also included.

</div>
<div style="border: 1px solid #ccc; border-radius: 12px; padding: 16px; background: #fff8f0;">

### 🔧 Card B

- List item 1
- List item 2 ✅

</div>
</div>

---

# Slide 16: Card component (border only, no background)

<div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px;">
<div style="border: 2px solid #0066cc; border-radius: 8px; padding: 12px;">

**Blue border card**  
No background

</div>
<div style="border: 2px solid #cc6600; border-radius: 8px; padding: 12px;">

**Orange border card**  
No background

</div>
<div style="border: 2px solid #006600; border-radius: 8px; padding: 12px;">

**Green border card**  
No background

</div>
</div>

---

# Slide 17: Nested containers (3 levels)

<div style="border: 2px solid #333; padding: 16px; background: #f5f5f5; border-radius: 8px;">

**Outer container** — has background and border

<div style="border: 1px solid #999; padding: 12px; background: white; border-radius: 4px; margin-top: 8px;">

Middle container — white background

<div style="background: #fff3cd; padding: 8px; border-radius: 4px; margin-top: 6px;">
Inner container — yellow background
</div>

</div>
</div>

---

# Slide 18: Step badges (circled numbers)

<div style="display: flex; flex-direction: column; gap: 12px;">
<div style="display: flex; align-items: center; gap: 10px;">
<span style="display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px; border-radius: 50%; background: #0066cc; color: white; font-weight: bold;">1</span>
<span>Label-1: Alpha beta gamma</span>
</div>
<div style="display: flex; align-items: center; gap: 10px;">
<span style="display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px; border-radius: 50%; background: #0066cc; color: white; font-weight: bold;">2</span>
<span>Create a <strong>delta item</strong> and save it as <code>item.dat</code></span>
</div>
<div style="display: flex; align-items: center; gap: 10px;">
<span style="display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px; border-radius: 50%; background: #0066cc; color: white; font-weight: bold;">3</span>
<span>Check result ✅</span>
</div>
</div>

---

# Slide 19: Text and image mixed container

<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; align-items: start;">
<div>

### Text side

**Label-A**: Alpha beta gamma  
**Label-B**: Delta epsilon  
**Label-C**: Zeta eta ✅

Theta iota kappa lambda.

</div>
<div style="border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">

![w:100%](./mock-screenshot.png)

</div>
</div>

---

# Slide 20: Header & Footer

<style scoped>
header { display: block; }
footer { display: block; }
</style>

<!-- _header: "Test Header Text" -->
<!-- _footer: "Test Footer — Page number: {{ paginate }}" -->

This slide verifies that the header and footer are positioned correctly.

Body text.

---

<!-- Japanese text intentionally kept for CJK rendering test: special characters and fullwidth symbols -->

# Slide 21: Special characters & HTML entities

Verify the following special characters display correctly:

- `<`, `>`, `&`, `"`, `'` (HTML entities)
- Fullwidth symbols: 「括弧」『二重括弧』【太括弧】〔亀甲括弧〕
- Math-like: x² + y² = z², α・β・γ
- Long vowel mark: ー (katakana) included in words
- Repetition mark: 〜がテ゛スト〜
- Zero-width character check: 通常テキスト (verify no invisible characters are mixed in)

---

<!-- Japanese text intentionally kept for CJK rendering test: long paragraph wrapping -->

# Slide 22: Long paragraph text (wrapping test)

これは非常に長い段落テキストのテストです。日本語のテキストが長くなった場合に正しく折り返し表示されるか確認します。英語テキストも混在させます：The quick brown fox jumps over the lazy dog. この文章には特に意味がありません。段落が長くなっても途中でクリップされないことを確認します。改行なしでどこまで続くか試してみます。そしてまだ続きます。長い長いテキストがどこかの時点で折り返されることを期待します。

Next paragraph. Also checking the spacing between paragraphs.

Third paragraph. Verify that three paragraphs are displayed independently.

---

# Slide 23: Content overflow at bottom of slide

<style scoped>
section { font-size: 20px; }
</style>

## Main section

Normal content.

<div style="margin-top: 8px; display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
<div style="border: 1px solid #ccc; border-radius: 8px; padding: 12px; background: #f0f8ff;">
<strong>Left card</strong><br>
Content line 1<br>
Content line 2<br>
Content line 3
</div>
<div style="border: 1px solid #ccc; border-radius: 8px; padding: 12px; background: #fff0f0;">
<strong>Right card</strong><br>
Content line 1<br>
Content line 2<br>
Content line 3
</div>
</div>

<div style="margin-top: 8px; padding: 10px; background: #fffde7; border-radius: 6px; border: 1px solid #f0c040; font-size: 16px;">
⚠️ This is a note box at the bottom of the slide. For overflow testing. Do not replace <code>item.cfg</code> — append to it instead.
</div>

---

# Slide 24: Blockquote variations

> Single-line blockquote

> Multi-line  
> blockquote  
> test

> Blockquote containing **bold**, _italic_, and `code`

---

# Slide 25: Table (complex cells)

| Step | Role       | Input              | Output     |
| :--: | :--------: | :----------------- | :--------- |
|  1   | Role-A     | Input-A            | Output-A   |
|  2   | **Role-B** | Input-B + Input-C  | Output-B   |
|  3   | Role-C     | Output-B           | Output-C   |
|  4   | `Role-D`   | Output-C + Input-A | Output-D   |
|  5   | Role-E ✅  | Input-D + Output-B | Output-E   |

---

# Slide 26: Code block (background color, long lines)

```javascript
// Very long single line of code — wrapping/clipping test
const veryLongVariableName = someFunction({
  key1: 'value1',
  key2: 'value2',
  key3: 'value3',
  key4: 'value4',
})

// Code with comments
const config = {
  apiKey: process.env.API_KEY, // Loaded from environment variable
  timeout: 30000, // Timeout: 30 seconds
  retries: 3, // Retry count
}
```

---

# Slide 27: Container with image (same-dir image test)

<div style="display: grid; grid-template-columns: 1.2fr 0.8fr; gap: 16px; align-items: start; margin-top: 12px;">
<div>

### Left column (text)

- Item 1 ✅
- **Important item 2**
- `Item with code 3`
- Normal item 4

Additional text below.

</div>
<div style="border: 1px solid #ddd; border-radius: 8px; padding: 4px;">

![w:100%](./mock-screenshot.png)

</div>
</div>

---

# Slide 28: Empty container / minimal content

Empty div (border only, no child elements):

<div style="border: 2px dashed #999; height: 60px; border-radius: 8px;"></div>

Text-only div (no child elements, inline only):

<div style="border: 1px solid #369; border-radius: 6px; padding: 12px; background: #e8f4ff;">Text only. Contains <strong>bold</strong> and <em>italic</em>.</div>

---

# Slide 29: Incremental font sizes

<div style="font-size: 28px;">Large text (28px)</div>
<div style="font-size: 22px;">Normal text (22px)</div>
<div style="font-size: 18px;">Slightly smaller text (18px)</div>
<div style="font-size: 14px;">Small text (14px)</div>
<div style="font-size: 11px;">Very small text (11px) — Check the minimum pt size in PPTX</div>

---

# Slide 30: Color variations

<div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; margin-top: 8px;">
<div style="background: #1a73e8; color: white; padding: 12px; border-radius: 6px;">Blue bg<br>White text</div>
<div style="background: #34a853; color: white; padding: 12px; border-radius: 6px;">Green bg<br>White text</div>
<div style="background: #ea4335; color: white; padding: 12px; border-radius: 6px;">Red bg<br>White text</div>
<div style="background: #fbbc04; color: #333; padding: 12px; border-radius: 6px;">Yellow bg<br>Dark text</div>
<div style="background: #9c27b0; color: white; padding: 12px; border-radius: 6px;">Purple bg<br>White text</div>
<div style="background: #ff5722; color: white; padding: 12px; border-radius: 6px;">Orange bg<br>White text</div>
</div>

---

# Slide 31: SVG shapes (embedded SVG)

An SVG is embedded below:

<svg width="200" height="100" viewBox="0 0 200 100">
  <rect x="10" y="10" width="80" height="60" rx="8" fill="#0066cc" />
  <circle cx="140" cy="40" r="35" fill="#cc6600" />
  <text x="50" y="45" text-anchor="middle" fill="white" font-size="14">BOX</text>
  <text x="140" y="45" text-anchor="middle" fill="white" font-size="14">CIR</text>
</svg>

Verify the SVG is correctly converted as an image.

---

# Slide 32: Header + all-elements compound test (final check)

## Heading H2 ✅

**Bold** _Italic_ `code` ~~strikethrough~~ [Link](https://example.com) Emoji🎉

<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 12px;">
<div style="border: 1px solid #ccc; border-radius: 8px; padding: 12px;">

### Left card 🔷

- Item A ✅
- Item B 🔧
- **Bold item** C

```js
const x = 1
```

</div>
<div>

| Col A | Col B |
| :---- | ----: |
| Val 1 |   100 |
| Val 2 |   200 |

> Blockquote text

</div>
</div>

---

# Slide 33: 👉 Emoji between text (tight list) ← Important bug verification

Alpha beta gamma delta

- Item chain test: "Item-A👉Item-B👉Item-C👉Item-D👉Item-E👉Item-F👉Item-G👉Item-H👉Item-I👉Item-J👉Item-K"
- Zeta eta theta description item
- A👉B👉C (tight list, direct img)
- Flow A 👉 Flow B 👉 Flow C

**Also verify outside bullet lists:**

Text👉Text👉Text (no spaces before/after)

Previous text 👉 Next text (with spaces)

Verification in lists:

1. Step A👉Step B👉Step C
2. Step 👉 Eta

---

# Slide 34: Inline badges (border-radius rounded) ← Important bug verification

**Status badges (span with rounded background):**

Current status: <span style="background:#e53e3e;color:white;padding:2px 10px;border-radius:12px;font-size:0.85em">Status-A</span>
Next step: <span style="background:#3182ce;color:white;padding:2px 10px;border-radius:12px;font-size:0.85em">Status-B</span>
Complete: <span style="background:#38a169;color:white;padding:2px 10px;border-radius:12px;font-size:0.85em">Status-C ✅</span>

**Labeled badges (center-aligned with display:inline-flex):**

<span style="display:inline-flex;align-items:center;background:#c05621;color:white;padding:4px 12px;border-radius:16px;font-weight:bold;margin-right:4px">LVL-H</span><span style="display:inline-flex;align-items:center;background:#dd6b20;color:white;padding:4px 12px;border-radius:16px;font-weight:bold;margin-right:4px">LVL-M</span><span style="display:inline-flex;align-items:center;background:#2f855a;color:white;padding:4px 12px;border-radius:16px;font-weight:bold">LVL-L</span>

**Circle badges (step numbers, perfect circle):**

<span style="display:inline-flex;align-items:center;justify-content:center;background:#3182ce;color:white;width:28px;height:28px;border-radius:50%;font-weight:bold;margin-right:6px">1</span> Step-1 <span style="display:inline-flex;align-items:center;justify-content:center;background:#3182ce;color:white;width:28px;height:28px;border-radius:50%;font-weight:bold;margin-right:6px">2</span> Create **item** file <span style="display:inline-flex;align-items:center;justify-content:center;background:#38a169;color:white;width:28px;height:28px;border-radius:50%;font-weight:bold;margin-right:6px">✅</span> Check result

---

<!-- _backgroundImage: url('./test-icon.png') -->

# Slide 35: Background image filter (![bg grayscale]) ← Important bug verification

**This background uses \_backgroundImage (CSS background-image).**
Background image test without filter. Verify that text is not displayed twice.

---

<!-- _class: lead -->

# Slide 36: SVG + Badge + Emoji compound test

**Flow 👉**

<span style="display:inline-flex;align-items:center;justify-content:center;background:#6b46c1;color:white;width:40px;height:40px;font-size:0.8em;margin-right:0">Step-A</span><span style="display:inline-flex;align-items:center;justify-content:center;background:#38a169;color:white;width:40px;height:40px;border-radius:50%;margin-left:0"> ✓</span>

1. Step-A 👉 Step-B 👉 Step-C
2. Check <span style="background:#c05621;color:white;padding:1px 8px;border-radius:8px;font-size:0.8em">Status-X</span>

---

# Slide 37: Mermaid (flowchart)

<script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>

<div class="mermaid">
flowchart LR
  A[Alpha beta gamma] --> B[Delta epsilon zeta]
  B --> C[Eta theta iota]
  C --> A
</div>

---

# Slide 38: Mermaid (sequence diagram)

<div class="mermaid">
sequenceDiagram
  Node-A->>Node-B: Input
  Node-B-->>Node-A: Output
</div>

---

# Slide 39: Mermaid (multiple diagrams, 2-column layout)

<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
<div>

**Flowchart:**

<div class="mermaid">
flowchart LR
  A[Step-A] --> B[Step-B] --> C[Step-C]
</div>

</div>
<div>

**Sequence diagram:**

<div class="mermaid">
sequenceDiagram
  Node-A->>Node-B: Input
  Node-B-->>Node-A: Output
</div>

</div>
</div>

---

<!-- ============================================================
     Slides 40+: Additional test cases (generalized patterns from production materials)
     ============================================================ -->

<style>
/* === CSS Custom Properties test === */
:root {
  --brand: #2563eb;
  --brand-soft: #dbeafe;
  --accent: #ea580c;
  --ink: #1e293b;
  --surface: #f8fafc;
}

/* === Pseudo-element bar decoration test === */
section.decorated::before {
  content: '';
  display: block;
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 12px;
  background: var(--brand);
}
section.decorated::after {
  content: '';
  display: block;
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 8px;
  background: var(--accent);
}

/* === Gradient section background test === */
section.gradient-bg {
  background: linear-gradient(180deg, #e0f2fe 0%, #ffffff 40%);
}
section.gradient-accent {
  background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
}

/* === Marker-style strong test === */
.marker-highlight strong {
  background: linear-gradient(transparent 62%, #fff2a8 62%);
}

/* === Heading border-left decoration test === */
section.border-heading h2 {
  border-left: 5px solid var(--brand);
  padding-left: 12px;
}

/* === Step badge (scoped style version) === */
.step-badge {
  display: inline-block;
  min-width: 34px;
  padding: 4px 10px;
  margin-right: 10px;
  border-radius: 999px;
  background: var(--brand);
  color: #ffffff;
  font-size: 18px;
  text-align: center;
}

/* === Chat bubble === */
.chat-container {
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.msg-bubble {
  max-width: 80%;
  padding: 10px 14px;
  border-radius: 16px;
  box-shadow: 0 2px 6px rgba(0,0,0,0.08);
  font-size: 15px;
  line-height: 1.5;
}
.msg-bubble.user {
  align-self: flex-end;
  background: #dcf8c6;
  color: #1e293b;
}
.msg-bubble.assistant {
  align-self: flex-start;
  background: #ffffff;
  border: 1px solid #e2e8f0;
  color: #1e293b;
}

/* === Note box === */
.note-box {
  border-left: 6px solid var(--accent);
  background: #fff7ed;
  padding: 10px 14px;
  border-radius: 0 8px 8px 0;
  font-size: 15px;
}

/* === Card + box-shadow === */
.shadow-card {
  border-radius: 12px;
  padding: 16px;
  background: white;
  box-shadow: 0 10px 24px rgba(0,0,0,0.06);
}

/* === Flow diagram pattern === */
.flow-parent {
  display: grid;
  grid-template-columns: 1fr auto 1fr;
  align-items: center;
  gap: 12px;
}
.flow-arrow {
  font-size: 24px;
  color: var(--brand);
}
.flow-children {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 8px;
}

/* === letter-spacing === */
.ls-heading h1, .ls-heading h2 {
  letter-spacing: 0.04em;
}
</style>

<!-- _class: decorated gradient-bg -->

# Slide 40: CSS Custom Properties + Pseudo-element bars

**::before (top blue bar) and ::after (bottom orange bar) should be displayed.**

This slide uses `var(--brand)` and `var(--accent)`. Test whether pseudo-elements are captured in screenshots or ignored.

Text color specified with `var(--ink)`: <span style="color: var(--ink);">This text is dark.</span>

---

<!-- _class: gradient-accent border-heading -->

# Slide 41: Gradient background + border-left heading

## Section heading (with left border)

Background is `linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%)`. Heading h2 has `border-left: 5px solid var(--brand)` applied.

- Point 1: CSS Custom Properties should resolve correctly
- Point 2: gradient should fall back or be screenshotted in PPTX
- Point 3: border-left should appear as a blue line

---

# Slide 42: Marker-style highlight (gradient strong)

<div class="marker-highlight">

This text has a **marker-style highlight** applied. The style `linear-gradient(transparent 62%, #fff2a8 62%)` colours only the bottom 38% yellow.

**Alpha labels** applied with this technique. Delta epsilon zeta theta iota.

</div>

---

# Slide 43: Scoped style step badge (pill shape)

<p><span class="step-badge">01</span> Alpha beta gamma</p>
<p><span class="step-badge">02</span> Delta epsilon item</p>
<p><span class="step-badge">03</span> Zeta eta and <strong>theta iota</strong></p>

**The badges above should render as pills (border-radius: 999px) with blue background and white text.**

---

# Slide 44: nth-child table column colouring

<style scoped>
table th:nth-child(3), table td:nth-child(3) { background: #dbeafe !important; color: #1e40af !important; }
table th:nth-child(4), table td:nth-child(4) { background: #fee2e2 !important; color: #991b1b !important; }
table th:nth-child(5), table td:nth-child(5) { background: #d1fae5 !important; color: #065f46 !important; }
section { font-size: 20px; }
</style>

| Item   | Category |   Mode-1    |   Mode-2    | Result |
| :----- | :------- | :---------: | :---------: | :----: |
| Item-A | Cat-A    | Pattern-1   | Pattern-2   |   ◎    |
| Item-B | Cat-B    | Pattern-3   | Pattern-4   |   ○    |
| Item-C | Cat-C    | Pattern-1   | Pattern-2   |   ◎    |
| Item-D | Cat-D    | Pattern-5   | Pattern-6   |   △    |

**Column 3 (blue), column 4 (red), and column 5 (green) should be colour-coded.**

---

# Slide 45: Chat bubble UI

<div class="chat-container">
<div class="msg-bubble user">Alpha beta gamma delta epsilon?</div>
<div class="msg-bubble assistant">Zeta eta, two options.<br><strong>1. Method-A</strong>: theta iota kappa<br><strong>2. Method-B</strong>: lambda mu nu</div>
<div class="msg-bubble user">Xi omicron pi rho sigma tau?</div>
<div class="msg-bubble assistant">A <code>upsilon</code>-based phi pattern is common. It works well for chi psi omega items.</div>
</div>

---

# Slide 46: box-shadow cards

<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 16px;">
<div class="shadow-card">

### 📊 Label-A

- Label-B: **val-1234**
- Label-C: **val-98**
- Label-D: val-2 /uu

</div>
<div class="shadow-card" style="border-left: 4px solid #ea580c;">

### ⚠️ Note

Watch out for limit — **val-60 items** per /uu max.
Exceeding will return a val-429 result.

</div>
</div>

---

# Slide 47: Note box + strong + code mixed

<div class="note-box">
<strong>ℹ️ Note:</strong> Add value to <code>item.cfg</code>. <strong>Do not replace</strong> existing items — append instead.
</div>

<div style="margin-top: 16px;">

Normal paragraph text follows the note box.

</div>

<div class="note-box" style="border-left-color: #dc2626; background: #fef2f2;">
<strong>⚠️ Warning:</strong> This action is <strong>irreversible</strong>. Please back up before running.
</div>

---

# Slide 48: Flow diagram pattern (parent → arrow → children)

<div style="text-align: center; margin-bottom: 12px;">

<div style="display: inline-block; padding: 10px 24px; background: var(--brand); color: white; border-radius: 8px; font-weight: bold;">Input item</div>

</div>

<div style="text-align: center; font-size: 24px; color: var(--brand); margin: 8px 0;">▼</div>

<div class="flow-children">
<div style="border: 2px solid var(--brand); border-radius: 8px; padding: 10px; text-align: center; background: var(--brand-soft);">
<strong>Step A</strong><br>Alpha process
</div>
<div style="border: 2px solid var(--accent); border-radius: 8px; padding: 10px; text-align: center; background: #fff7ed;">
<strong>Step B</strong><br>Beta format
</div>
<div style="border: 2px solid #16a34a; border-radius: 8px; padding: 10px; text-align: center; background: #dcfce7;">
<strong>Step C</strong><br>Gamma output
</div>
</div>

---

<!-- _class: ls-heading -->

# Slide 49: letter-spacing test

## Heading with letter-spacing: 0.04em

Normal paragraph text (no letter-spacing). The difference in character spacing from the heading should be visible.

|     | Setting | Effect        |
| :-- | :------ | :------------ |
| h1  | 0.04em  | Wider spacing |
| h2  | 0.04em  | Wider spacing |
| p   | none    | Default       |

---

![bg cover blur:5px brightness:0.3 grayscale](./test-icon.png)

# Slide 50: Marp bg filter syntax test

<style scoped>
section { color: white; }
h1 { color: white; }
</style>

**`![bg cover blur:5px brightness:0.3 grayscale]` — Multiple filters applied simultaneously.**

Background image should appear blurred, darkened, and greyscale. Text is white.

---

![bg right:30%](./test-icon.png)

# Slide 51: Split background test

**`![bg right:30%]` — Image occupies right 30%, content uses left 70%.**

- Verify split layout is reproduced correctly
- Text should be on the left side
- Image should occupy the right 30%

---

<!-- _class: decorated -->

# Slide 52: Negative margin + fixed-width container

<div style="display: flex; gap: 16px; align-items: flex-start;">
<div style="flex: 1;">

### Text content

Normal flow text. The image on the right slightly overlaps upward.

- Item A
- Item B ✅

</div>
<div style="width: 300px; margin-top: -20px; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">

![w:100%](./test-icon.png)

</div>
</div>

---

# Slide 53: HTML inline flex side-by-side images

<div style="display: flex; gap: 16px; align-items: center;">
<div style="flex: 1;">

![w:100%](./test-icon.png)

</div>
<div style="align-content: center; font-size: 24px; color: var(--brand);">→</div>
<div style="flex: 1;">

![w:100%](./mock-screenshot.png)

</div>
</div>

<div style="text-align: center; margin-top: 8px; color: #64748b; font-size: 14px;">Left: Before → Right: After</div>

---

# Slide 54: h3 + step badge (heading with mixed badge)

<style scoped>
.step2 {
  display: inline-block;
  min-width: 32px;
  padding: 3px 10px;
  margin-right: 8px;
  border-radius: 999px;
  background: #0f6cbd;
  color: white;
  font-size: 18px;
  text-align: center;
}
</style>

<h3><span class="step2">1</span> Alpha beta gamma</h3>
<h3><span class="step2">2</span> Delta epsilon zeta</h3>
<h3><span class="step2">3</span> Eta theta iota kappa</h3>

---

# Slide 55: section::before/after (banner suppression test)

<style scoped>
section::before {
  content: "";
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 16px;
  background: #0f6cbd;
}
section::after {
  content: "";
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 10px;
  background: #2563eb;
}
</style>

Verify that the PPTX for this slide does **not display top/bottom banners**.

section::before and section::after with content:"" should **not** be extracted as banners.

---

# Slide 56: h2 border-left vertical decoration + ZWJ emoji

<style scoped>
section h2 {
  color: #2980b9;
  border-left: 4px solid #3498db;
  padding-left: 10px;
}
strong {
  background-color: #f1c40f;
  padding: 2px 4px;
  border-radius: 4px;
}
</style>

## Vertical decoration heading: text must not overlap the bar

Text should start to the **right** of the border-left bar.
Developer: 🧑‍💻 (ZWJ composed emoji)
Family: 👨‍👩‍👧‍👦 (family ZWJ sequence)

- Working on **alpha beta** items
- The 4px border-left should render as a vertical bar with no text overlap

---

# Slide 57: Section-scoped colour themes (class-scoped h2 border)

<style scoped>
section h2 {
  color: #27ae60;
  border-left: 4px solid #27ae60;
  padding-left: 10px;
}
section.topic-alt h2 {
  color: #8e44ad;
  border-left: 4px solid #8e44ad;
}
table { font-size: 18px; width: 100%; }
th { background-color: #27ae60; color: white; }
</style>

## Theme-colour vertical decoration

| Item       | Mode-1    | Mode-2    |
| ---------- | --------- | --------- |
| **Item-A** | State-1   | State-2   |
| **Item-B** | State-3   | State-4   |
| **Item-C** | State-5   | State-6   |

---

# Slide 58: Solid background strong highlight

<style scoped>
strong {
  background-color: #f1c40f;
  padding: 2px 4px;
  border-radius: 4px;
}
img:not(.emoji) {
  border: 4px solid #333;
}
</style>

## Solid-colour strong tag

Inline highlighting of **alpha items** and **val-42 labels** in paragraphs.
Solid colour (#f1c40f) background — not a gradient — should render correctly.

- **method-A** approach for beta check
- **Label-H / Label-M** item data

---

# Slide 59: Compound background filters (blur + brightness + grayscale)

<style scoped>
section h2 {
  border-left: 4px solid #e74c3c;
  padding-left: 10px;
  color: #c0392b;
}
</style>

## Background image with filters

![bg cover blur:5px brightness:3 grayscale](https://via.placeholder.com/1280x720/3498db/ffffff?text=BG)

Three filters applied simultaneously to the background image: blur + brightness + grayscale.
Foreground text is displayed with a border-left vertical decoration heading.

---

<!-- _class: decorated -->

# Slide 60: global section::before + user class (no banner test)

<style scoped>
section::before {
  content: "";
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 16px;
  background: #16324f;
}
section::after {
  content: "";
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 10px;
  background: #0f6cbd;
}
</style>

This slide has **class="decorated"** and also defines `section::before` / `section::after` as scoped styles.

Because the same pseudo-element also applies to a classless slide (slide 55), the rule is treated as a **global theme decoration** and **no banner should appear** at the top or bottom.

- Verify: no dark navy bar at top of slide
- Verify: no blue bar at bottom of slide
- Only the content text should be visible

---

# Slide 61: Grid + flex agenda (direct text node bug regression)

<!-- Regression for ADR-15: each agenda item is display:flex and contains
     an inline-flex badge plus a direct child TEXT_NODE. The PPTX export must
     render the description text to the right of the badge with no overlap. -->

<style scoped>
.agenda-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 28px;
  margin-top: 24px;
}

.agenda-item {
  display: flex;
  align-items: flex-start;
  gap: 14px;
  font-size: 22px;
  line-height: 1.5;
}

.agenda-num {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  background: #0f6cbd;
  color: #fff;
  border-radius: 50%;
  width: 32px;
  height: 32px;
  font-size: 15px;
  font-weight: 700;
  flex-shrink: 0;
}
</style>

<div class="agenda-grid">
  <div class="agenda-item"><span class="agenda-num">1</span> Topic-1</div>
  <div class="agenda-item"><span class="agenda-num">2</span> Topic-2</div>
  <div class="agenda-item"><span class="agenda-num">3</span> Topic-3</div>
  <div class="agenda-item"><span class="agenda-num">4</span> Topic-4</div>
  <div class="agenda-item"><span class="agenda-num">5</span> Topic-5</div>
  <div class="agenda-item"><span class="agenda-num">6</span> Topic-6</div>
</div>

---

# Slide 62: Trailing-space hard line break inside list item

<!-- Regression for <br> handling in extractListItemEl:
     Two trailing spaces before newline produce <br> inside <li>.
     The PPTX must render two visible lines within the same bullet. -->

- Line one  
  Line two (same bullet, separated by trailing-space line break)
- Normal item (no line break)
- Third bullet  
  Also has a second line

---

# Slide 63: Image between list items (no blank lines)

<!-- Regression for image-in-list split logic in walkElements:
     When a Markdown image appears between list items with no blank lines,
     markdown-it parses it as a lazy continuation of the first <li>.
     The PPTX must show items in order: First item → image → Second item. -->

- First item
![w:200](./test-icon.png)
- Second item

---

# Slide 64: Text nodes alongside inline-block badges in block container

<!-- ADR-22 regression: when a display:block div contains both direct text
     nodes and inline-block elements (badges), the shallow recovery pass
     must recover the text nodes. Before fix, only <strong> text was kept;
     surrounding text nodes were dropped. -->

<style scoped>
.tl-fix { display: flex; flex-direction: column; gap: 0; margin-top: 12px; }
.tl-fix-row { display: flex; align-items: flex-start; gap: 8px; padding: 5px 0; border-bottom: 1px dotted #ccc; }
.tl-fix-row:last-child { border-bottom: none; }
.tl-fix-yr { min-width: 54px; font-weight: 700; font-size: 12px; color: #888; }
.tl-fix-body { font-size: 16px; line-height: 1.5; }
.tl-fix-badge { display: inline-block; border-radius: 3px; padding: 1px 6px; font-size: 13px; font-weight: 700; margin-left: 4px; }
.tl-fix-badge.blue { background: #eef5fc; color: #0f6cbd; }
.tl-fix-badge.green { background: #edf8f1; color: #27ae60; }
</style>

<div class="tl-fix">
<div class="tl-fix-row">
<div class="tl-fix-yr">Row-1</div>
<div class="tl-fix-body">Alpha item and <strong>beta gamma</strong>. Delta epsilon<span class="tl-fix-badge blue">Tag-A</span></div>
</div>
<div class="tl-fix-row">
<div class="tl-fix-yr">Row-2</div>
<div class="tl-fix-body">Zeta nu eta, <strong>theta iota</strong> kappa<span class="tl-fix-badge green">Tag-B</span></div>
</div>
<div class="tl-fix-row">
<div class="tl-fix-yr">Row-3</div>
<div class="tl-fix-body">Alpha text and <strong>beta gamma</strong> delta item</div>
</div>
</div>

Expected: All text visible including "Alpha item and", ". Delta epsilon",
"Zeta nu eta, " and " kappa" (text nodes adjacent to badges).

---

# Slide 65: Dense table inside grid panel with narrow cells

<!-- ADR-22 regression: table CSS padding should be used as PPTX cell margin
     instead of a fixed value, reducing font-metric wrapping in dense tables. -->

<style scoped>
section { font-size: 16px; }
.panel-grid-fix { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 8px; }
.panel-fix { border: 1px solid #ccc; border-radius: 8px; padding: 10px 12px; }
.panel-fix table { margin-top: 0; width: 100%; font-size: 13px; }
.panel-fix th, .panel-fix td { padding: 2px 4px; }
.panel-fix th { background: #e8f0f8; }
</style>

<div class="panel-grid-fix">
<div class="panel-fix">

| Col-A | Col-B | Col-C | Col-D | Col-E |
|---|---|:---:|:---:|---|
| Cat-A | Item-1 | n-1 | **val-p1** | (r1–r9) |
| Cat-A | Item-2 | n-2 | **val-p2** | (r2–r9) |
| Cat-B | Item-3 | n-3 | **val-p1** | (r2–r8) |
| Cat-B | Item-4 | n-4 | **val-p2** | (r2–r9) |
| Cat-C | Item-5 | n-4 | **val-p1** | (r2–r9) |

</div>
<div class="panel-fix">

| Col-A | Col-B | Col-C | Col-D | Col-E |
|---|---|:---:|:---:|---|
| Cat-A | Item-6 | n-1 | **val-p1** | (r1–r9) |
| Cat-A | Item-7 | n-2 | **val-p3** | (r2–r8) |
| Cat-B | Item-8 | n-3 | **val-p1** | (r2–r9) |
| Cat-B | Item-9 | n-2 | **val-p2** | (r2–r9) |
| Cat-C | Item-10 | n-2 | **val-p2** | (r2–r9) |

</div>
</div>

Expected: Table text should not wrap within cells. Tight padding (2px 4px) matches HTML.

---

# Slide 66: Step list with inline-block tag badges

<!-- ADR-22 regression: step-body divs contain text nodes + inline-block
     span tags. Same root cause as Slide 64: text nodes dropped in shallow
     recovery when display:block container has inline-block children. -->

<style scoped>
section { font-size: 13px; }
.step-grid-fix { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-top: 10px; }
.step-col-fix { border: 1px solid #ccc; border-radius: 10px; padding: 12px 14px; }
.step-title-fix { font-weight: 700; padding-bottom: 5px; border-bottom: 3px solid #0f6cbd; margin-bottom: 8px; color: #0f6cbd; }
.step-list-fix { display: flex; flex-direction: column; gap: 5px; }
.step-fix { display: flex; align-items: flex-start; gap: 7px; }
.step-num-fix { min-width: 21px; height: 21px; border-radius: 50%; color: #fff; font-weight: 700; display: flex; align-items: center; justify-content: center; background: #0f6cbd; font-size: 12px; }
.step-body-fix { line-height: 1.45; }
.stag-fix { display: inline-block; border-radius: 3px; padding: 0 5px; font-weight: 700; }
.stag-fix.g { background: #edf8f1; color: #27ae60; }
.stag-fix.b { background: #eef5fc; color: #0f6cbd; }
</style>

<div class="step-grid-fix">
<div class="step-col-fix">
<div class="step-title-fix">Group A</div>
<div class="step-list-fix">
<div class="step-fix"><div class="step-num-fix">1</div><div class="step-body-fix">Alpha task and beta item<span class="stag-fix g">Tag-A</span></div></div>
<div class="step-fix"><div class="step-num-fix">2</div><div class="step-body-fix">Gamma delta item epsilon<span class="stag-fix g">Tag-B</span></div></div>
<div class="step-fix"><div class="step-num-fix">3</div><div class="step-body-fix">Zeta eta theta item<span class="stag-fix b">Tag-C</span></div></div>
</div>
</div>
<div class="step-col-fix">
<div class="step-title-fix">Group B</div>
<div class="step-list-fix">
<div class="step-fix"><div class="step-num-fix">1</div><div class="step-body-fix">Iota kappa lambda</div></div>
<div class="step-fix"><div class="step-num-fix">2</div><div class="step-body-fix">Mu nu xi omicron<span class="stag-fix b">Tag-D</span></div></div>
<div class="step-fix"><div class="step-num-fix">3</div><div class="step-body-fix">Pi rho sigma tau<span class="stag-fix g">Tag-E</span></div></div>
</div>
</div>
</div>

Expected: Step descriptions fully visible, e.g. "Alpha task and beta item"
before the Tag-A badge.

---

# Slide 67: Right-aligned text in flex row (wrapping regression — nowrap width test)

<!-- ADR-22 regression: flex-child divs with white-space:nowrap have tight
     bounding boxes. PPTX font metrics may be slightly wider, causing text
     to wrap. The width is extended by 10% to absorb font-metric variance. -->

<style scoped>
.nw-grid { border: 1px solid #ccc; border-radius: 8px; overflow: hidden; margin-top: 12px; }
.nw-row { display: flex; align-items: center; padding: 9px 14px; border-bottom: 1px solid #ccc; gap: 10px; }
.nw-row:last-child { border-bottom: none; background: #f0f7ff; }
.nw-op { font-weight: 700; color: #888; min-width: 16px; text-align: center; }
.nw-label { color: #888; font-size: 15px; flex: 1; line-height: 1.4; }
.nw-val { font-weight: 800; font-size: 22px; letter-spacing: -0.5px; white-space: nowrap; }
.nw-val.blue { color: #0f6cbd; }
</style>

<div class="nw-grid">
<div class="nw-row">
<div class="nw-op"> </div>
<div class="nw-label">Label-A<br><span>val-X × val-N (label-1)</span></div>
<div class="nw-val">val-1000<span style="font-weight:400; color:#888;">/uu</span></div>
</div>
<div class="nw-row">
<div class="nw-op">+</div>
<div class="nw-label">Label-B</div>
<div class="nw-val blue">val-0–val-100<span style="font-weight:400; color:#888;">/uu</span></div>
</div>
<div class="nw-row">
<div class="nw-op" style="color:#0f6cbd;">=</div>
<div class="nw-label" style="font-weight:700; color:#0f6cbd;">Label-C (label-2)</div>
<div class="nw-val" style="font-weight:900; color:#0f6cbd;">val-1000–val-2000<span style="font-weight:400; color:#888;">/uu</span></div>
</div>
</div>

Expected: nw-val text should stay on one line without wrapping.

---

# Slide 68: Narrow emoji-icon div in flex row with adjacent text (ADR-30 regression)

<!-- ADR-30 regression: a flex-row child containing only emoji (min-width: 18px,
     text-align: center) had emojiWidthOverride extend its text box to the
     parent's right edge.  The centred emoji then overlapped the adjacent text
     div.  Fix: skip emojiWidthOverride when a visible sibling follows. -->

<style scoped>
.icon-row-col { display: flex; flex-direction: column; gap: 5px; margin-top: 8px; }
.icon-row { display: flex; align-items: flex-start; gap: 7px; }
.icon-cell { min-width: 18px; text-align: center; }
</style>

<div class="icon-row-col">
<div class="icon-row">
<div class="icon-cell">🔒</div>
<div>Alpha beta gamma delta</div>
</div>
<div class="icon-row">
<div class="icon-cell">📝</div>
<div>Delta epsilon zeta eta</div>
</div>
<div class="icon-row">
<div class="icon-cell">🔗</div>
<div>Theta iota kappa lambda</div>
</div>
</div>

Expected: each emoji icon stays left-aligned in its 18 px cell; text div starts immediately to the right without overlap.
