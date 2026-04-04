"use strict";
var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/native-pptx/index.ts
var index_exports = {};
__export(index_exports, {
  generateNativePptx: () => generateNativePptx
});
module.exports = __toCommonJS(index_exports);
var import_promises = require("node:fs/promises");
var import_node_url2 = require("node:url");
var import_puppeteer_core = __toESM(require("puppeteer-core"));

// src/native-pptx/dom-walker-script.generated.ts
var DOM_WALKER_SCRIPT = '"use strict";\nvar DomWalker = (() => {\n  var __defProp = Object.defineProperty;\n  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;\n  var __getOwnPropNames = Object.getOwnPropertyNames;\n  var __hasOwnProp = Object.prototype.hasOwnProperty;\n  var __export = (target, all) => {\n    for (var name in all)\n      __defProp(target, name, { get: all[name], enumerable: true });\n  };\n  var __copyProps = (to, from, except, desc) => {\n    if (from && typeof from === "object" || typeof from === "function") {\n      for (let key of __getOwnPropNames(from))\n        if (!__hasOwnProp.call(to, key) && key !== except)\n          __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });\n    }\n    return to;\n  };\n  var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);\n\n  // src/native-pptx/dom-walker.ts\n  var dom_walker_exports = {};\n  __export(dom_walker_exports, {\n    extractSlides: () => extractSlides\n  });\n  function extractSlides(root = document) {\n    function findBackgroundColor(section) {\n      const style = getComputedStyle(section);\n      const bg = style.backgroundColor;\n      if (bg && bg !== "transparent" && bg !== "rgba(0, 0, 0, 0)") {\n        return bg;\n      }\n      const bgImage = style.backgroundImage;\n      if (bgImage && bgImage !== "none") {\n        const colorMatches = bgImage.match(\n          /rgba?\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+(?:\\s*,\\s*[\\d.]+)?\\s*\\)/g\n        );\n        if (colorMatches && colorMatches.length > 0) {\n          for (let i = colorMatches.length - 1; i >= 0; i--) {\n            const c = colorMatches[i];\n            if (c !== "rgba(0, 0, 0, 0)") {\n              const alphaMatch = c.match(\n                /rgba\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*([\\d.]+)\\s*\\)/\n              );\n              if (!alphaMatch || parseFloat(alphaMatch[1]) > 0.1) {\n                return c;\n              }\n            }\n          }\n        }\n      }\n      return "rgb(255, 255, 255)";\n    }\n    function extractTextStyle(style) {\n      let textAlign = style.textAlign || "left";\n      if (textAlign === "start") textAlign = "left";\n      else if (textAlign === "end") textAlign = "right";\n      if (textAlign === "left" && style.justifyContent === "center")\n        textAlign = "center";\n      return {\n        color: style.color,\n        fontSize: parseFloat(style.fontSize) || 16,\n        fontFamily: style.fontFamily,\n        fontWeight: parseInt(style.fontWeight, 10) || 400,\n        textAlign,\n        lineHeight: parseFloat(style.lineHeight) || 0,\n        letterSpacing: parseFloat(style.letterSpacing) || 0\n      };\n    }\n    function isEmojiImg(imgEl) {\n      const alt = imgEl.alt ?? "";\n      return !!(imgEl.classList?.contains("emoji") || imgEl.src && (imgEl.src.includes("twemoji") || imgEl.src.includes("/emoji/")) || alt.length > 0 && alt.length <= 8 && /\\p{Extended_Pictographic}/u.test(alt));\n    }\n    function extractTextRuns(element, skipInlineBadges = false) {\n      const runs = [];\n      const elementStyle = getComputedStyle(element);\n      const elementBg = elementStyle.backgroundColor;\n      const elementHasBg = !!elementBg && elementBg !== "transparent" && elementBg !== "rgba(0, 0, 0, 0)";\n      function lastIsBreak() {\n        return runs.length > 0 && runs[runs.length - 1].breakLine === true;\n      }\n      function pushText(text, style, bg) {\n        const segments = text.split("\\n");\n        for (let i = 0; i < segments.length; i++) {\n          const seg = segments[i];\n          if (seg !== "") {\n            const run = {\n              text: seg,\n              color: style.color,\n              fontSize: parseFloat(style.fontSize) || 16,\n              fontFamily: style.fontFamily,\n              bold: parseInt(style.fontWeight, 10) >= 600,\n              italic: style.fontStyle === "italic",\n              underline: style.textDecorationLine?.includes("underline"),\n              strikethrough: style.textDecorationLine?.includes("line-through")\n            };\n            if (bg) run.backgroundColor = bg;\n            runs.push(run);\n          }\n          if (i < segments.length - 1 && !lastIsBreak()) {\n            runs.push({ text: "", breakLine: true });\n          }\n        }\n      }\n      for (const node of Array.from(element.childNodes)) {\n        if (node.nodeType === Node.TEXT_NODE) {\n          const text = node.textContent ?? "";\n          if (text.trim() === "") {\n            const newlineCount = (text.match(/\\n/g) ?? []).length;\n            for (let i = 0; i < newlineCount; i++) {\n              runs.push({ text: "", breakLine: true });\n            }\n            continue;\n          }\n          pushText(text, elementStyle, elementHasBg ? elementBg : void 0);\n        } else if (node.nodeType === Node.ELEMENT_NODE) {\n          const el = node;\n          const tag = el.tagName.toLowerCase();\n          if (tag === "br") {\n            if (!lastIsBreak()) runs.push({ text: "", breakLine: true });\n            continue;\n          }\n          if (tag === "a") {\n            const href = el.href;\n            const childRuns = extractTextRuns(el);\n            childRuns.forEach((r) => {\n              if (!r.breakLine) r.hyperlink = href;\n            });\n            runs.push(...childRuns);\n            continue;\n          }\n          if (tag === "img") {\n            const imgEl = el;\n            const alt = imgEl.alt ?? "";\n            if (isEmojiImg(imgEl) && alt) {\n              pushText(alt, getComputedStyle(el), void 0);\n            }\n            continue;\n          }\n          const elStyle = getComputedStyle(el);\n          if (/^(block|flex|grid|list-item|table)/.test(elStyle.display)) {\n            if (!lastIsBreak() && runs.length > 0) {\n              runs.push({ text: "", breakLine: true });\n            }\n            runs.push(...extractTextRuns(el));\n          } else {\n            const bg = elStyle.backgroundColor;\n            const hasBg = bg && bg !== "transparent" && bg !== "rgba(0, 0, 0, 0)";\n            const alphaZero = hasBg && (() => {\n              const m = bg.match(/,\\s*([\\d.]+)\\s*\\)$/);\n              return m ? parseFloat(m[1]) === 0 : false;\n            })();\n            const isBadge = hasBg && !alphaZero && (elStyle.display === "inline-block" || elStyle.display === "inline-flex" || elStyle.display === "inline-grid");\n            if (isBadge) {\n              if (skipInlineBadges) {\n                continue;\n              }\n              const childRuns2 = extractTextRuns(el, false);\n              childRuns2.forEach((r) => {\n                if (!r.breakLine && !r.backgroundColor) r.backgroundColor = bg;\n              });\n              runs.push(...childRuns2);\n              continue;\n            }\n            const childRuns = extractTextRuns(el, skipInlineBadges);\n            if (hasBg && !alphaZero) {\n              childRuns.forEach((r) => {\n                if (!r.breakLine && !r.backgroundColor) r.backgroundColor = bg;\n              });\n            }\n            runs.push(...childRuns);\n          }\n        }\n      }\n      while (runs.length > 0 && runs[runs.length - 1].breakLine) {\n        runs.pop();\n      }\n      while (runs.length > 0 && runs[0].breakLine) {\n        runs.shift();\n      }\n      return runs;\n    }\n    function computeLeadingOffset(badgeShapes, containerRect, slideRect) {\n      if (badgeShapes.length === 0) return 0;\n      const containerSSLeft = containerRect.left - slideRect.left;\n      const leading = badgeShapes.filter((b) => b.x <= containerSSLeft + 8);\n      if (leading.length === 0) return 0;\n      const rightEdge = leading.reduce(\n        (max, b) => Math.max(max, b.x + b.width),\n        containerSSLeft\n      );\n      return Math.max(0, rightEdge - containerSSLeft);\n    }\n    function extractListItems(list, level = 0) {\n      const items = [];\n      for (const child of Array.from(list.children)) {\n        const tag = child.tagName.toLowerCase();\n        if (tag === "li") {\n          const runs = [];\n          const nestedItems = [];\n          for (const node of Array.from(child.childNodes)) {\n            if (node.nodeType === Node.TEXT_NODE) {\n              const text = node.textContent ?? "";\n              if (text.trim() === "") continue;\n              const liStyle = getComputedStyle(child);\n              const liBg = liStyle.backgroundColor;\n              const liHasBg = !!liBg && liBg !== "transparent" && liBg !== "rgba(0, 0, 0, 0)";\n              const segments = text.split("\\n");\n              for (let i = 0; i < segments.length; i++) {\n                const seg = segments[i];\n                if (seg !== "") {\n                  const run = {\n                    text: seg,\n                    color: liStyle.color,\n                    fontSize: parseFloat(liStyle.fontSize) || 16,\n                    fontFamily: liStyle.fontFamily,\n                    bold: parseInt(liStyle.fontWeight, 10) >= 600,\n                    italic: liStyle.fontStyle === "italic"\n                  };\n                  if (liHasBg) run.backgroundColor = liBg;\n                  runs.push(run);\n                }\n                if (i < segments.length - 1) {\n                  runs.push({ text: "", breakLine: true });\n                }\n              }\n            } else if (node.nodeType === Node.ELEMENT_NODE) {\n              const el = node;\n              const childTag = el.tagName.toLowerCase();\n              if (childTag === "ul" || childTag === "ol") {\n                nestedItems.push(...extractListItems(el, level + 1));\n              } else if (childTag === "img") {\n                const imgEl = el;\n                const alt = imgEl.alt ?? "";\n                if (isEmojiImg(imgEl) && alt) {\n                  const liStyle = getComputedStyle(child);\n                  runs.push({\n                    text: alt,\n                    color: liStyle.color,\n                    fontSize: parseFloat(liStyle.fontSize) || 16,\n                    fontFamily: liStyle.fontFamily,\n                    bold: parseInt(liStyle.fontWeight, 10) >= 600,\n                    italic: liStyle.fontStyle === "italic"\n                  });\n                }\n              } else {\n                runs.push(...extractTextRuns(el));\n              }\n            }\n          }\n          if (runs.length > 0) {\n            const combinedText = runs.map((r) => r.text).join("");\n            items.push({ text: combinedText.trim(), level, runs });\n          }\n          items.push(...nestedItems);\n        }\n      }\n      return items;\n    }\n    function extractCodeRuns(codeEl) {\n      const runs = [];\n      const defaultStyle = getComputedStyle(codeEl);\n      function walk(node) {\n        if (node.nodeType === Node.TEXT_NODE) {\n          const text = node.textContent ?? "";\n          if (text === "") return;\n          const parent = node.parentElement ?? codeEl;\n          const style = getComputedStyle(parent);\n          runs.push({\n            text,\n            color: style.color,\n            fontSize: parseFloat(style.fontSize) || parseFloat(defaultStyle.fontSize) || 16,\n            fontFamily: style.fontFamily || defaultStyle.fontFamily,\n            bold: parseInt(style.fontWeight, 10) >= 600,\n            italic: style.fontStyle === "italic"\n          });\n        } else if (node.nodeType === Node.ELEMENT_NODE) {\n          for (const child of Array.from(node.childNodes)) {\n            walk(child);\n          }\n        }\n      }\n      walk(codeEl);\n      return runs;\n    }\n    function extractTableData(table) {\n      const rows = [];\n      const colWidths = [];\n      for (const tr of Array.from(table.querySelectorAll("tr"))) {\n        const cells = [];\n        const isFirstRow = rows.length === 0;\n        for (const td of Array.from(tr.querySelectorAll("th, td"))) {\n          const style = getComputedStyle(td);\n          if (isFirstRow) {\n            colWidths.push(td.offsetWidth);\n          }\n          cells.push({\n            text: td.textContent ?? "",\n            runs: extractTextRuns(td),\n            isHeader: td.tagName.toLowerCase() === "th",\n            style: {\n              color: style.color,\n              backgroundColor: style.backgroundColor,\n              fontSize: parseFloat(style.fontSize) || 16,\n              fontFamily: style.fontFamily,\n              fontWeight: parseInt(style.fontWeight, 10) || 400,\n              textAlign: style.textAlign || "left",\n              borderColor: style.borderColor\n            }\n          });\n        }\n        rows.push({ cells });\n      }\n      return { rows, colWidths };\n    }\n    function extractInlineBadgeShapes(container, slideRect) {\n      const badges = [];\n      for (const el of Array.from(container.querySelectorAll("*"))) {\n        const s = getComputedStyle(el);\n        if (s.display !== "inline-block" && s.display !== "inline-flex" && s.display !== "inline-grid")\n          continue;\n        const bg = s.backgroundColor;\n        if (!bg || bg === "transparent") continue;\n        const alphaMatch = bg.match(/,\\s*([\\d.]+)\\s*\\)$/);\n        if (alphaMatch && parseFloat(alphaMatch[1]) === 0) continue;\n        const iRect = el.getBoundingClientRect();\n        if (iRect.width === 0 || iRect.height === 0) continue;\n        const br = parseFloat(s.borderRadius) || 0;\n        const badgeRuns = extractTextRuns(el);\n        badgeRuns.forEach((r) => {\n          if (!r.breakLine) r.backgroundColor = void 0;\n        });\n        const hasBadgeText = badgeRuns.some(\n          (r) => !r.breakLine && r.text.trim() !== ""\n        );\n        badges.push({\n          type: "container",\n          children: [],\n          ...hasBadgeText ? { runs: badgeRuns } : {},\n          x: iRect.left - slideRect.left,\n          y: iRect.top - slideRect.top,\n          width: iRect.width,\n          height: iRect.height,\n          style: {\n            backgroundColor: bg,\n            ...br > 0 ? { borderRadius: br } : {}\n          }\n        });\n      }\n      return badges;\n    }\n    function extractNestedImages(el, slideRect) {\n      const images = [];\n      for (const img of Array.from(el.querySelectorAll("img"))) {\n        const imgEl = img;\n        const rect = imgEl.getBoundingClientRect();\n        if (rect.width === 0 || rect.height === 0) continue;\n        const s = getComputedStyle(imgEl);\n        if (s.display === "none" || s.visibility === "hidden") continue;\n        if (imgEl.classList?.contains("emoji") || imgEl.src?.includes("twemoji") || imgEl.src?.includes("/emoji/"))\n          continue;\n        const cssFilter = s.filter && s.filter !== "none" ? s.filter : void 0;\n        images.push({\n          type: "image",\n          src: imgEl.src,\n          naturalWidth: imgEl.naturalWidth,\n          naturalHeight: imgEl.naturalHeight,\n          x: rect.left - slideRect.left,\n          y: rect.top - slideRect.top,\n          width: rect.width,\n          height: rect.height,\n          ...cssFilter ? { cssFilter, pageX: rect.left, pageY: rect.top } : {}\n        });\n      }\n      return images;\n    }\n    function walkElements(parent, slideRect) {\n      const elements = [];\n      for (const child of Array.from(parent.children)) {\n        const style = getComputedStyle(child);\n        if (style.display === "none" || style.visibility === "hidden") continue;\n        if (child.dataset?.marpitPresenterNotes !== void 0)\n          continue;\n        if (child.dataset?.marpitAdvancedBackgroundContainer !== void 0)\n          continue;\n        const rect = child.getBoundingClientRect();\n        if (rect.width === 0 || rect.height === 0) continue;\n        const tag = child.tagName.toLowerCase();\n        const parentIsFlexOrGrid = /^(flex|inline-flex|grid|inline-grid)/.test(\n          getComputedStyle(parent).display\n        );\n        if (!parentIsFlexOrGrid && tag !== "img" && tag !== "svg" && style.display === "inline")\n          continue;\n        const base = {\n          x: rect.left - slideRect.left,\n          y: rect.top - slideRect.top,\n          width: rect.width,\n          height: rect.height\n        };\n        if (/^h[1-6]$/.test(tag)) {\n          const borderBottomWidth = parseFloat(style.borderBottomWidth) || 0;\n          const borderLeftWidth = parseFloat(style.borderLeftWidth) || 0;\n          const headingBadgeShapes = extractInlineBadgeShapes(child, slideRect);\n          const headingLeadingOffset = computeLeadingOffset(\n            headingBadgeShapes,\n            rect,\n            slideRect\n          );\n          if (headingBadgeShapes.length > 0) elements.push(...headingBadgeShapes);\n          const headingRuns = extractTextRuns(\n            child,\n            headingBadgeShapes.length > 0\n          );\n          if (headingBadgeShapes.length === 0 || headingRuns.some((r) => !r.breakLine && r.text.trim() !== "")) {\n            elements.push({\n              type: "heading",\n              level: parseInt(tag[1], 10),\n              runs: headingRuns,\n              ...base,\n              x: base.x + headingLeadingOffset,\n              width: Math.max(10, base.width - headingLeadingOffset),\n              style: extractTextStyle(style),\n              ...borderBottomWidth > 0 ? {\n                borderBottom: {\n                  width: borderBottomWidth,\n                  color: style.borderBottomColor\n                }\n              } : {},\n              ...borderLeftWidth > 0 ? {\n                borderLeft: {\n                  width: borderLeftWidth,\n                  color: style.borderLeftColor\n                }\n              } : {}\n            });\n          }\n          elements.push(...extractNestedImages(child, slideRect));\n        } else if (tag === "p") {\n          const paraBadgeShapes = extractInlineBadgeShapes(child, slideRect);\n          const paraLeadingOffset = computeLeadingOffset(\n            paraBadgeShapes,\n            rect,\n            slideRect\n          );\n          if (paraBadgeShapes.length > 0) elements.push(...paraBadgeShapes);\n          const runs = extractTextRuns(child, paraBadgeShapes.length > 0);\n          if (runs.some((r) => !r.breakLine && r.text.trim() !== "")) {\n            elements.push({\n              type: "paragraph",\n              runs,\n              ...base,\n              x: base.x + paraLeadingOffset,\n              width: Math.max(10, base.width - paraLeadingOffset),\n              style: extractTextStyle(style)\n            });\n          }\n          elements.push(...extractNestedImages(child, slideRect));\n        } else if (tag === "ul" || tag === "ol") {\n          elements.push({\n            type: "list",\n            ordered: tag === "ol",\n            items: extractListItems(child),\n            ...base,\n            style: extractTextStyle(style)\n          });\n          elements.push(...extractNestedImages(child, slideRect));\n        } else if (tag === "table") {\n          const { rows: tableRows, colWidths } = extractTableData(child);\n          elements.push({\n            type: "table",\n            rows: tableRows,\n            ...colWidths.length > 0 ? { colWidths } : {},\n            ...base,\n            style: extractTextStyle(style)\n          });\n          elements.push(...extractNestedImages(child, slideRect));\n        } else if (tag === "pre") {\n          const innerSvg = child.querySelector("svg");\n          if (innerSvg) {\n            try {\n              const svgStr = new XMLSerializer().serializeToString(innerSvg);\n              const b64 = btoa(unescape(encodeURIComponent(svgStr)));\n              const dataUrl = `data:image/svg+xml;base64,${b64}`;\n              elements.push({\n                type: "image",\n                src: dataUrl,\n                naturalWidth: base.width,\n                naturalHeight: base.height,\n                ...base,\n                // Request rasterization: Mermaid SVGs may use <foreignObject>\n                // for text labels which PowerPoint cannot render from SVG data.\n                // pageX/pageY are intentionally omitted: rasterizeSlideTargets\n                // computes the absolute clip from the slide-relative x/y after\n                // navigating to the correct slide (avoids stale bespoke-transform\n                // coordinates).\n                rasterize: true\n              });\n            } catch {\n              const code = child.querySelector("code");\n              const codeTarget = code ?? child;\n              elements.push({\n                type: "code",\n                text: codeTarget.textContent ?? "",\n                language: code?.className?.replace("language-", "") ?? "",\n                runs: extractCodeRuns(codeTarget),\n                ...base,\n                style: {\n                  ...extractTextStyle(style),\n                  backgroundColor: style.backgroundColor\n                }\n              });\n            }\n          } else {\n            const code = child.querySelector("code");\n            const codeTarget = code ?? child;\n            elements.push({\n              type: "code",\n              text: codeTarget.textContent ?? "",\n              language: code?.className?.replace("language-", "") ?? "",\n              runs: extractCodeRuns(codeTarget),\n              ...base,\n              style: {\n                ...extractTextStyle(style),\n                backgroundColor: style.backgroundColor\n              }\n            });\n          }\n        } else if (tag === "img") {\n          const img = child;\n          if (img.classList?.contains("emoji") || img.src?.includes("twemoji") || img.src?.includes("/emoji/"))\n            continue;\n          const imgFilter = style.filter && style.filter !== "none" ? style.filter : void 0;\n          elements.push({\n            type: "image",\n            src: img.src,\n            naturalWidth: img.naturalWidth,\n            naturalHeight: img.naturalHeight,\n            ...base,\n            // Store page-absolute coords when cssFilter is set so the export\n            // tool can screenshot the rendered (filtered) region via Puppeteer.\n            ...imgFilter ? { cssFilter: imgFilter, pageX: rect.left, pageY: rect.top } : {}\n          });\n        } else if (tag === "blockquote") {\n          const borderWidth = parseFloat(style.borderLeftWidth) || 0;\n          const borderColor = style.borderLeftColor;\n          const bqBadgeShapes = extractInlineBadgeShapes(child, slideRect);\n          const bqLeadingOffset = computeLeadingOffset(\n            bqBadgeShapes,\n            rect,\n            slideRect\n          );\n          if (bqBadgeShapes.length > 0) elements.push(...bqBadgeShapes);\n          elements.push({\n            type: "blockquote",\n            runs: extractTextRuns(child, bqBadgeShapes.length > 0),\n            ...base,\n            x: base.x + bqLeadingOffset,\n            width: Math.max(10, base.width - bqLeadingOffset),\n            style: extractTextStyle(style),\n            ...borderWidth > 0 ? { borderLeft: { width: borderWidth, color: borderColor } } : {}\n          });\n          elements.push(...extractNestedImages(child, slideRect));\n        } else if (tag === "svg") {\n          try {\n            const svgStr = new XMLSerializer().serializeToString(child);\n            const b64 = btoa(unescape(encodeURIComponent(svgStr)));\n            const dataUrl = `data:image/svg+xml;base64,${b64}`;\n            elements.push({\n              type: "image",\n              src: dataUrl,\n              naturalWidth: base.width,\n              naturalHeight: base.height,\n              ...base\n            });\n          } catch {\n          }\n        } else if (tag === "header" || tag === "footer") {\n          const hfBadgeShapes = extractInlineBadgeShapes(child, slideRect);\n          const hfLeadingOffset = computeLeadingOffset(\n            hfBadgeShapes,\n            rect,\n            slideRect\n          );\n          if (hfBadgeShapes.length > 0) elements.push(...hfBadgeShapes);\n          elements.push({\n            type: tag,\n            runs: extractTextRuns(child, hfBadgeShapes.length > 0),\n            ...base,\n            x: base.x + hfLeadingOffset,\n            width: Math.max(10, base.width - hfLeadingOffset),\n            style: extractTextStyle(style)\n          });\n          elements.push(...extractNestedImages(child, slideRect));\n        } else {\n          const borderTopWidth = parseFloat(style.borderTopWidth) || 0;\n          const borderTopStyle = style.borderTopStyle;\n          const hasBorder = borderTopWidth > 0 && borderTopStyle !== "none";\n          const borderRadius = parseFloat(style.borderRadius) || 0;\n          const borderLeftWidth = parseFloat(style.borderLeftWidth) || 0;\n          const borderLeftStyle = style.borderLeftStyle;\n          const hasBorderLeft = borderLeftWidth > 0 && borderLeftStyle !== "none" && !hasBorder;\n          const boxShadow = style.boxShadow;\n          const hasBoxShadow = !!boxShadow && boxShadow !== "none";\n          const hasBackground = !!style.backgroundColor && style.backgroundColor !== "transparent" && style.backgroundColor !== "rgba(0, 0, 0, 0)" && !style.backgroundColor.match(\n            /rgba\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*0(?:\\.0+)?\\s*\\)/\n          );\n          const blockChildren = walkElements(child, slideRect);\n          const containerStyle = {\n            backgroundColor: style.backgroundColor,\n            ...hasBorder ? { borderWidth: borderTopWidth, borderColor: style.borderTopColor } : {},\n            ...borderRadius > 0 ? { borderRadius } : {},\n            ...hasBorderLeft ? {\n              borderLeft: {\n                width: borderLeftWidth,\n                color: style.borderLeftColor\n              }\n            } : {},\n            ...hasBoxShadow ? { boxShadow: true } : {}\n          };\n          if (blockChildren.length > 0) {\n            elements.push({\n              type: "container",\n              children: blockChildren,\n              ...base,\n              style: containerStyle\n            });\n          } else {\n            if (hasBackground || hasBorder || hasBorderLeft || hasBoxShadow) {\n              elements.push({\n                type: "container",\n                children: [],\n                ...base,\n                style: containerStyle\n              });\n            }\n            const runs = extractTextRuns(child);\n            if (runs.some((r) => !r.breakLine && r.text.trim() !== "")) {\n              const valign = style.alignItems === "center" || style.justifyContent === "center" || style.verticalAlign === "middle" ? "middle" : "top";\n              const paddingTop = parseFloat(style.paddingTop) || 0;\n              const paddingRight = parseFloat(style.paddingRight) || 0;\n              const paddingBottom = parseFloat(style.paddingBottom) || 0;\n              const paddingLeft = parseFloat(style.paddingLeft) || 0;\n              elements.push({\n                type: "paragraph",\n                runs,\n                ...base,\n                style: {\n                  ...extractTextStyle(style),\n                  ...paddingTop || paddingRight || paddingBottom || paddingLeft ? { paddingTop, paddingRight, paddingBottom, paddingLeft } : {}\n                },\n                valign\n              });\n            }\n            elements.push(...extractNestedImages(child, slideRect));\n          }\n        }\n      }\n      return elements;\n    }\n    function extractPseudoElements(section, slideRect) {\n      const shapes = [];\n      for (const pseudo of ["::before", "::after"]) {\n        const ps = getComputedStyle(section, pseudo);\n        const rawContent = ps.content;\n        if (!rawContent || rawContent === "none" || rawContent === "normal")\n          continue;\n        const stripped = rawContent.replace(/^["\']|["\']$/g, "").trim();\n        if (stripped === "") continue;\n        const bg = ps.backgroundColor;\n        if (!bg || bg === "transparent" || bg === "rgba(0, 0, 0, 0)") continue;\n        const w = parseFloat(ps.width) || 0;\n        const h = parseFloat(ps.height) || 0;\n        if (w === 0 && h === 0) continue;\n        const position = ps.position;\n        let x = 0;\n        let y = 0;\n        if (position === "absolute" || position === "fixed") {\n          const top = parseFloat(ps.top);\n          const left = parseFloat(ps.left);\n          const bottom = parseFloat(ps.bottom);\n          if (!isNaN(top)) y = top;\n          else if (!isNaN(bottom)) y = slideRect.height - bottom - h;\n          if (!isNaN(left)) x = left;\n        }\n        const effectiveW = w || slideRect.width;\n        const effectiveH = h || 0;\n        if (effectiveH <= 0) continue;\n        shapes.push({\n          type: "container",\n          children: [],\n          x,\n          y,\n          width: effectiveW,\n          height: effectiveH,\n          style: {\n            backgroundColor: bg\n          }\n        });\n      }\n      return shapes;\n    }\n    const allSections = Array.from(root.querySelectorAll("section")).filter(\n      (section) => {\n        if (section.parentElement?.closest("section")) return false;\n        if (section.parentElement?.tagName.toLowerCase() === "foreignobject") {\n          return true;\n        }\n        return section.hasAttribute("data-marpit-pagination");\n      }\n    );\n    const slideGroups = /* @__PURE__ */ new Map();\n    for (const [index, section] of allSections.entries()) {\n      const key = section.getAttribute("data-marpit-pagination") ?? section.getAttribute("id") ?? String(index);\n      const layer = section.getAttribute("data-marpit-advanced-background");\n      if (!slideGroups.has(key)) slideGroups.set(key, {});\n      const entry = slideGroups.get(key);\n      if (layer === "content") {\n        entry.content = section;\n      } else if (layer === "background") {\n        entry.background = section;\n      } else if (!layer) {\n        entry.content = section;\n      }\n    }\n    return Array.from(slideGroups.values()).map(\n      ({ content, background }, slideIdx) => {\n        const section = content ?? background;\n        const sectionRect = section.getBoundingClientRect();\n        const sectionStyle = getComputedStyle(section);\n        const backgroundImages = [];\n        if (background) {\n          const figures = background.querySelectorAll("figure");\n          for (const fig of Array.from(figures)) {\n            const figStyle = getComputedStyle(fig);\n            if (!figStyle.backgroundImage || figStyle.backgroundImage === "none")\n              continue;\n            const urlMatch = figStyle.backgroundImage.match(\n              /url\\(["\']?([^"\')]+)["\']?\\)/\n            );\n            if (!urlMatch) continue;\n            const figRect = fig.getBoundingClientRect();\n            const cssFilter = figStyle.filter && figStyle.filter !== "none" ? figStyle.filter : void 0;\n            backgroundImages.push({\n              url: urlMatch[1],\n              x: figRect.left - sectionRect.left,\n              y: figRect.top - sectionRect.top,\n              width: figRect.width || sectionRect.width,\n              height: figRect.height || sectionRect.height,\n              ...cssFilter ? { cssFilter } : {},\n              pageX: figRect.left,\n              pageY: figRect.top\n            });\n          }\n        }\n        if (backgroundImages.length === 0) {\n          const bgImg = sectionStyle.backgroundImage;\n          if (bgImg && bgImg !== "none") {\n            const urlMatch = bgImg.match(/url\\(["\']?([^"\')]+)["\']?\\)/);\n            if (urlMatch) {\n              backgroundImages.push({\n                url: urlMatch[1],\n                x: 0,\n                y: 0,\n                width: sectionRect.width,\n                height: sectionRect.height,\n                pageX: sectionRect.left,\n                pageY: sectionRect.top,\n                fromCssFallback: true\n              });\n            } else if (/gradient\\s*\\(/.test(bgImg)) {\n              backgroundImages.push({\n                url: "",\n                // placeholder \u2014 will be replaced by rasterized data URL\n                x: 0,\n                y: 0,\n                width: sectionRect.width,\n                height: sectionRect.height,\n                pageX: sectionRect.left,\n                pageY: sectionRect.top,\n                fromCssFallback: true\n              });\n            }\n          }\n        }\n        return {\n          width: sectionRect.width,\n          height: sectionRect.height,\n          background: findBackgroundColor(section),\n          backgroundImages,\n          elements: [\n            // Pseudo-element bars (::before/::after) go behind content\n            ...extractPseudoElements(section, sectionRect),\n            ...walkElements(section, sectionRect)\n          ],\n          notes: (section.querySelector("[data-marpit-presenter-notes]") ?? root.querySelector(`.bespoke-marp-note[data-index="${slideIdx}"]`))?.textContent?.trim() ?? ""\n        };\n      }\n    );\n  }\n  return __toCommonJS(dom_walker_exports);\n})();\n\nglobalThis.extractSlides = DomWalker.extractSlides;\n';

// src/native-pptx/slide-builder.ts
var import_node_url = require("node:url");
var import_pptxgenjs = __toESM(require("pptxgenjs"));

// src/native-pptx/utils.ts
function rgbToHex(rgb) {
  if (!rgb) return "000000";
  const match = rgb.match(
    /rgba?\(\s*(-?\d+)\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*(?:,\s*[\d.]+\s*)?\)/
  );
  if (!match) return "000000";
  const r = parseInt(match[1], 10);
  const g = parseInt(match[2], 10);
  const b = parseInt(match[3], 10);
  return [r, g, b].map((c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, "0")).join("").toUpperCase();
}
var genericFontMap = {
  "sans-serif": "Calibri",
  serif: "Cambria",
  monospace: "Courier New",
  cursive: "Calibri",
  fantasy: "Calibri",
  "ui-sans-serif": "Calibri",
  "ui-serif": "Cambria",
  "ui-monospace": "Courier New"
};
var systemAliases = /* @__PURE__ */ new Set([
  "-apple-system",
  "blinkmacsystemfont",
  "system-ui",
  "ui-rounded",
  "ui-monospace"
]);
var macOnlyFonts = /* @__PURE__ */ new Set([
  "sfmono-regular",
  "sf mono",
  "menlo",
  "sf pro text",
  "sf pro display"
]);
var knownSystemFonts = {
  "segoe ui": "Segoe UI",
  "helvetica neue": "Arial",
  helvetica: "Arial"
};
var japaneseFontPattern = /(noto sans jp|noto sans cjk jp|noto sans|yu gothic ui|yu gothic|meiryo|biz udpgothic|biz udgothic|ms pgothic|ms gothic|hiragino sans)/i;
var japaneseTextPattern = /[\u3040-\u30ff\u3400-\u9fff\uf900-\ufaff]/;
function cleanFontFamily(css, sampleText) {
  if (!css) return "Calibri";
  const families = css.split(",");
  let fallback;
  const candidates = [];
  const hasJapaneseText = typeof sampleText === "string" && japaneseTextPattern.test(sampleText);
  for (const raw of families) {
    const name = raw.trim().replace(/^["']|["']$/g, "");
    if (!name) continue;
    const lower = name.toLowerCase();
    if (systemAliases.has(lower)) continue;
    if (macOnlyFonts.has(lower)) continue;
    const known = knownSystemFonts[lower];
    if (known) {
      candidates.push(known);
      continue;
    }
    const generic = genericFontMap[lower];
    if (generic) {
      if (!fallback) fallback = generic;
      continue;
    }
    candidates.push(name);
  }
  if (hasJapaneseText) {
    const japaneseCandidate = candidates.find(
      (candidate) => japaneseFontPattern.test(candidate)
    );
    if (japaneseCandidate) return japaneseCandidate;
    const nonProprietary = candidates.find(
      (c) => !/\b\d{2,4}[A-Z]{2,}/i.test(c)
      // skip names with numeric+suffix codes like "35HSJPDOC"
    );
    if (nonProprietary) return nonProprietary;
    return "Meiryo";
  }
  if (candidates.length > 0) {
    return candidates[0];
  }
  return fallback ?? "Calibri";
}
function pxToInches(px) {
  return px / 96;
}
function pxToPoints(px) {
  return px * 0.75;
}
function isTransparent(color) {
  if (!color) return true;
  if (color === "transparent") return true;
  const match = color.match(
    /rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*([\d.]+)\s*\)/
  );
  if (match) return parseFloat(match[1]) <= 0.01;
  return false;
}
function sanitizeText(text) {
  return text.replace(
    // eslint-disable-next-line no-control-regex
    /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\u200B\u200C\u2060]/g,
    ""
  );
}

// src/native-pptx/slide-builder.ts
function resolveImageSource(url) {
  if (url.startsWith("data:")) return { data: url };
  if (url.startsWith("file:")) return { path: (0, import_node_url.fileURLToPath)(url) };
  return { path: url };
}
function computeLineSpacing(style) {
  const { lineHeight, fontSize } = style;
  if (!lineHeight || !fontSize || lineHeight <= 0 || fontSize <= 0)
    return void 0;
  const m = lineHeight / fontSize;
  if (m < 0.5 || m > 4) return void 0;
  return Math.round(m * 100) / 100;
}
function computeCharSpacing(style) {
  const ls = style.letterSpacing;
  if (!ls || Math.abs(ls) < 0.1) return void 0;
  return Math.round(pxToPoints(ls) * 100) / 100;
}
function computeTextInset(style) {
  const pt = (style.paddingTop ?? 0) * 0.75;
  const pr = (style.paddingRight ?? 0) * 0.75;
  const pb = (style.paddingBottom ?? 0) * 0.75;
  const pl = (style.paddingLeft ?? 0) * 0.75;
  return pt || pr || pb || pl ? [pt, pr, pb, pl] : 0;
}
function buildPptx(slides) {
  const pptx = new import_pptxgenjs.default();
  const slideW = slides[0]?.width ?? 1280;
  const slideH = slides[0]?.height ?? 720;
  pptx.defineLayout({
    name: "MARP",
    width: pxToInches(slideW),
    height: pxToInches(slideH)
  });
  pptx.layout = "MARP";
  for (const slideData of slides) {
    const slide = pptx.addSlide();
    const bgColor = isTransparent(slideData.background) ? "FFFFFF" : rgbToHex(slideData.background);
    const bgImages = slideData.backgroundImages ?? [];
    const firstBg = bgImages[0];
    const isFullSlide = firstBg && !firstBg.cssFilter && firstBg.x <= 1 && firstBg.y <= 1 && Math.abs(firstBg.width - slideData.width) <= 2 && Math.abs(firstBg.height - slideData.height) <= 2;
    if (isFullSlide && bgImages.length === 1) {
      slide.background = resolveImageSource(firstBg.url);
    } else {
      slide.background = { fill: bgColor };
      for (const bg of bgImages) {
        const x = pxToInches(bg.x);
        const y = pxToInches(bg.y);
        const w = pxToInches(bg.width);
        const h = pxToInches(bg.height);
        const imgOpts = {
          x,
          y,
          w,
          h,
          ...resolveImageSource(bg.url)
        };
        slide.addImage(imgOpts);
      }
    }
    for (const el of slideData.elements) {
      placeElement(slide, el, slideData.width, slideData.height);
    }
    if (slideData.notes) {
      slide.addNotes(slideData.notes);
    }
  }
  return pptx;
}
var TEXT_ELEMENT_TYPES = /* @__PURE__ */ new Set([
  "heading",
  "paragraph",
  "list",
  "blockquote",
  "code",
  "table",
  "header",
  "footer"
]);
function placeElement(slide, el, slideW = 0, slideH = 0) {
  const x = pxToInches(el.x);
  const y = pxToInches(el.y);
  const w = pxToInches(el.width);
  const rawH = pxToInches(el.height);
  const h = slideH > 0 && TEXT_ELEMENT_TYPES.has(el.type) ? Math.min(rawH, Math.max(0.01, pxToInches(slideH) - y)) : rawH;
  switch (el.type) {
    case "heading": {
      const headingBorderW = el.borderLeft && el.borderLeft.width > 0 ? pxToInches(el.borderLeft.width) : 0;
      if (headingBorderW > 0) {
        slide.addShape("rect", {
          x,
          y,
          w: headingBorderW,
          h,
          fill: { color: rgbToHex(el.borderLeft.color) },
          line: { color: rgbToHex(el.borderLeft.color) }
        });
      }
      slide.addText(
        el.runs.map((r) => toTextProps(r)),
        {
          x: x + headingBorderW,
          y,
          w: Math.max(0.01, w - headingBorderW),
          h,
          margin: 0,
          valign: "top",
          align: el.style.textAlign,
          lineSpacingMultiple: computeLineSpacing(el.style),
          paraSpaceBefore: 0,
          charSpacing: computeCharSpacing(el.style)
        }
      );
      if (el.borderBottom && el.borderBottom.width > 0) {
        const bh = pxToInches(el.borderBottom.width);
        slide.addShape("rect", {
          x,
          y: y + h,
          w,
          h: bh,
          fill: { color: rgbToHex(el.borderBottom.color) },
          line: { color: rgbToHex(el.borderBottom.color) }
        });
      }
      break;
    }
    case "paragraph":
      slide.addText(
        el.runs.map((r) => toTextProps(r)),
        {
          x,
          y,
          w,
          h,
          margin: computeTextInset(el.style),
          valign: el.valign ?? "top",
          align: el.style.textAlign,
          lineSpacingMultiple: computeLineSpacing(el.style),
          paraSpaceBefore: 0,
          charSpacing: computeCharSpacing(el.style)
        }
      );
      break;
    case "header":
    case "footer":
      slide.addText(
        el.runs.map((r) => toTextProps(r)),
        {
          x,
          y,
          w,
          h,
          margin: 0,
          valign: "top",
          align: el.style.textAlign,
          lineSpacingMultiple: computeLineSpacing(el.style),
          paraSpaceBefore: 0,
          charSpacing: computeCharSpacing(el.style)
        }
      );
      break;
    case "blockquote":
      if (el.borderLeft && el.borderLeft.width > 0) {
        const bw = pxToInches(el.borderLeft.width);
        slide.addShape("rect", {
          x,
          y,
          w: bw,
          h,
          fill: { color: rgbToHex(el.borderLeft.color) }
        });
        slide.addText(
          el.runs.map((r) => toTextProps(r)),
          {
            x: x + bw,
            y,
            w: w - bw,
            h,
            margin: 0,
            valign: "top",
            align: el.style.textAlign,
            lineSpacingMultiple: computeLineSpacing(el.style),
            paraSpaceBefore: 0,
            charSpacing: computeCharSpacing(el.style)
          }
        );
      } else {
        slide.addText(
          el.runs.map((r) => toTextProps(r)),
          {
            x,
            y,
            w,
            h,
            margin: 0,
            valign: "top",
            align: el.style.textAlign,
            lineSpacingMultiple: computeLineSpacing(el.style),
            paraSpaceBefore: 0,
            charSpacing: computeCharSpacing(el.style)
          }
        );
      }
      break;
    case "list":
      slide.addText(
        el.items.flatMap(
          (item, index) => toListTextProps(item, el.ordered, index < el.items.length - 1)
        ),
        {
          x,
          y,
          w,
          h,
          margin: 0,
          valign: "top",
          align: el.style.textAlign,
          lineSpacingMultiple: computeLineSpacing(el.style),
          paraSpaceBefore: 0,
          charSpacing: computeCharSpacing(el.style)
        }
      );
      break;
    case "table":
      slide.addTable(
        el.rows.map(
          (row) => row.cells.map((cell) => {
            if (cell.runs && cell.runs.length > 0) {
              const cellOpts2 = {
                align: cell.style.textAlign
              };
              if (!isTransparent(cell.style.backgroundColor)) {
                cellOpts2.fill = { color: rgbToHex(cell.style.backgroundColor) };
              }
              if (cell.style.borderColor && !isTransparent(cell.style.borderColor)) {
                cellOpts2.border = {
                  pt: 1,
                  color: rgbToHex(cell.style.borderColor)
                };
              }
              return {
                text: cell.runs.map((r) => ({
                  text: sanitizeText(r.text),
                  options: {
                    color: rgbToHex(r.color),
                    fontSize: pxToPoints(r.fontSize ?? cell.style.fontSize),
                    fontFace: cleanFontFamily(
                      r.fontFamily ?? cell.style.fontFamily,
                      r.text
                    ),
                    bold: r.bold ?? cell.isHeader ?? cell.style.fontWeight >= 600,
                    italic: r.italic
                  }
                })),
                options: cellOpts2
              };
            }
            const cellOpts = {
              bold: cell.isHeader || cell.style.fontWeight >= 600,
              color: rgbToHex(cell.style.color),
              fontSize: pxToPoints(cell.style.fontSize),
              fontFace: cleanFontFamily(cell.style.fontFamily, cell.text),
              align: cell.style.textAlign
            };
            if (!isTransparent(cell.style.backgroundColor)) {
              cellOpts.fill = { color: rgbToHex(cell.style.backgroundColor) };
            }
            if (cell.style.borderColor && !isTransparent(cell.style.borderColor)) {
              cellOpts.border = {
                pt: 1,
                color: rgbToHex(cell.style.borderColor)
              };
            }
            return { text: sanitizeText(cell.text), options: cellOpts };
          })
        ),
        {
          x,
          y,
          w,
          autoPage: false,
          // Preserve HTML column proportions when per-column widths are available
          ...el.colWidths && el.colWidths.length > 0 && el.colWidths.every((cw) => cw > 0) ? {
            colW: el.colWidths.map((cw) => pxToInches(cw))
          } : {}
        }
      );
      break;
    case "code": {
      if (!isTransparent(el.style.backgroundColor)) {
        slide.addShape("rect", {
          x,
          y,
          w,
          h,
          fill: { color: rgbToHex(el.style.backgroundColor) }
        });
      }
      slide.addText(sanitizeText(el.text), {
        x,
        y,
        w,
        h,
        margin: 0,
        fontFace: "Courier New",
        fontSize: pxToPoints(el.style.fontSize),
        color: rgbToHex(el.style.color),
        valign: "top",
        paraSpaceBefore: 0
      });
      break;
    }
    case "image": {
      const imgOpts = {
        x,
        y,
        w,
        h,
        ...resolveImageSource(el.src)
      };
      slide.addImage(imgOpts);
      break;
    }
    case "container": {
      const bg = el.style?.backgroundColor;
      const borderWidth = el.style?.borderWidth ?? 0;
      const borderColor = el.style?.borderColor;
      const borderRadius = el.style?.borderRadius ?? 0;
      const borderLeft = el.style?.borderLeft;
      const hasBoxShadow = el.style?.boxShadow === true;
      const hasBackground = !isTransparent(bg);
      const hasBorder = borderWidth > 0 && !!borderColor && !isTransparent(borderColor);
      const lineStyle = hasBorder ? { color: rgbToHex(borderColor), width: pxToPoints(borderWidth) } : hasBoxShadow ? { color: "CCCCCC", width: 0.5 } : void 0;
      if (hasBackground || hasBorder || hasBoxShadow) {
        const shapeType = borderRadius > 0 ? "roundRect" : "rect";
        const minDim = Math.min(el.width, el.height);
        const rectRadius = borderRadius > 0 ? Math.min(0.5, borderRadius / (minDim / 2)) : void 0;
        slide.addShape(shapeType, {
          x,
          y,
          w,
          h,
          fill: hasBackground ? { color: rgbToHex(bg) } : { type: "none" },
          ...lineStyle ? { line: lineStyle } : {},
          ...rectRadius !== void 0 ? { rectRadius } : {}
        });
      }
      if (borderLeft && borderLeft.width > 0) {
        const bw = pxToInches(borderLeft.width);
        slide.addShape("rect", {
          x,
          y,
          w: bw,
          h,
          fill: { color: rgbToHex(borderLeft.color) },
          line: { color: rgbToHex(borderLeft.color) }
        });
      }
      if (el.runs && el.runs.length > 0 && el.runs.some((r) => !r.breakLine && r.text.trim() !== "")) {
        slide.addText(
          el.runs.map((r) => toTextProps(r)),
          {
            x,
            y,
            w,
            h,
            margin: 0,
            valign: "middle",
            align: "center",
            lineSpacingMultiple: 1,
            paraSpaceBefore: 0
          }
        );
      }
      if (hasBackground) {
        const bgHex = rgbToHex(bg);
        for (const child of el.children ?? []) {
          if ("runs" in child && Array.isArray(child.runs)) {
            for (const r of child.runs) {
              if (!r.breakLine && r.backgroundColor && rgbToHex(r.backgroundColor) === bgHex) {
                r.backgroundColor = void 0;
              }
            }
          }
        }
      }
      for (const child of el.children ?? []) {
        placeElement(slide, child, slideW, slideH);
      }
      break;
    }
  }
}
function toTextProps(run) {
  if (run.breakLine) {
    return { text: "", options: { breakLine: true } };
  }
  const text = sanitizeText(run.text);
  const highlight = run.backgroundColor ? rgbToHex(run.backgroundColor) : void 0;
  return {
    text,
    options: {
      color: rgbToHex(run.color),
      fontSize: pxToPoints(run.fontSize ?? 16),
      fontFace: cleanFontFamily(run.fontFamily, run.text),
      bold: run.bold,
      italic: run.italic,
      underline: run.underline ? { style: "sng" } : void 0,
      strike: run.strikethrough ? "sngStrike" : void 0,
      hyperlink: run.hyperlink ? { url: run.hyperlink } : void 0,
      highlight
    }
  };
}
function toListTextProps(item, ordered = false, breakAfter = false) {
  const bulletOption = ordered ? { type: "number", style: "arabicPeriod" } : true;
  if (item.runs.length === 0) {
    return [
      {
        text: sanitizeText(item.text) || " ",
        options: {
          bullet: bulletOption,
          indentLevel: item.level,
          breakLine: breakAfter
        }
      }
    ];
  }
  return item.runs.map((run, i) => ({
    text: sanitizeText(run.text),
    options: {
      ...i === 0 ? { bullet: bulletOption, indentLevel: item.level } : {},
      ...i === item.runs.length - 1 && breakAfter ? { breakLine: true } : {},
      color: rgbToHex(run.color),
      fontSize: pxToPoints(run.fontSize ?? 16),
      fontFace: cleanFontFamily(run.fontFamily, run.text),
      bold: run.bold,
      italic: run.italic
    }
  }));
}

// src/native-pptx/index.ts
async function generateNativePptx(opts) {
  const { htmlPath, browserPath, width = 1280, height = 720 } = opts;
  let browser;
  try {
    browser = await import_puppeteer_core.default.launch({
      executablePath: browserPath,
      headless: true,
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu"
      ]
    });
    const page = await browser.newPage();
    await page.setViewport({ width, height });
    const fileUrl = (0, import_node_url2.pathToFileURL)(htmlPath).href;
    await page.goto(fileUrl, { waitUntil: "networkidle0" });
    await page.addStyleTag({
      content: ".bespoke-marp-osc,[data-bespoke-marp-osc],.bespoke-marp-note{display:none!important}"
    });
    await page.addScriptTag({ content: DOM_WALKER_SCRIPT });
    const slides = await page.evaluate(
      () => globalThis.extractSlides()
    );
    await rasterizeSlideTargets(page, buildFilteredBgJobs(slides));
    await rasterizeSlideTargets(page, buildCssFallbackBgJobs(slides));
    await rasterizeSlideTargets(page, buildFilteredContentImageJobs(slides));
    await rasterizeSlideTargets(page, buildRasterizeImageJobs(slides));
    await rasterizeSlideTargets(page, buildPartialBgJobs(slides));
    if (opts.debugJsonPath) {
      await (0, import_promises.writeFile)(
        opts.debugJsonPath,
        JSON.stringify(slides, null, 2),
        "utf-8"
      );
    }
    const pptx = buildPptx(slides);
    const output = await pptx.write({ outputType: "nodebuffer" });
    return Buffer.from(output);
  } finally {
    await browser?.close();
  }
}
var NAVIGATION_SETTLE_MS = 300;
var POST_RASTERIZE_SETTLE_MS = 100;
async function rasterizeSlideTargets(page, jobs) {
  if (jobs.length === 0) return;
  for (const { slideIdx, targets, setup, teardown } of jobs) {
    await page.evaluate((n) => {
      window.location.hash = "#" + n;
    }, slideIdx + 1);
    await new Promise((r) => setTimeout(r, NAVIGATION_SETTLE_MS));
    let slideOriginX = 0;
    let slideOriginY = 0;
    if (targets.some((t) => t.slideRelative)) {
      const origin = await page.evaluate((n) => {
        const sec = Array.from(
          document.querySelectorAll("section")
        ).find((s) => s.getAttribute("data-marpit-pagination") === String(n));
        if (!sec) return { x: 0, y: 0 };
        const r = sec.getBoundingClientRect();
        return { x: r.left, y: r.top };
      }, slideIdx + 1);
      slideOriginX = origin.x;
      slideOriginY = origin.y;
    }
    try {
      if (setup) await setup(page, slideIdx);
      for (const { clip, slideRelative, onCapture } of targets) {
        const effectiveClip = slideRelative ? {
          x: Math.round(slideOriginX + clip.x),
          y: Math.round(slideOriginY + clip.y),
          width: clip.width,
          height: clip.height
        } : clip;
        if (effectiveClip.width <= 0 || effectiveClip.height <= 0) continue;
        try {
          const raw = await page.screenshot({
            type: "png",
            clip: effectiveClip
          });
          onCapture(
            "data:image/png;base64," + Buffer.from(raw).toString("base64")
          );
        } catch {
        }
      }
    } finally {
      if (teardown) await teardown(page, slideIdx);
    }
  }
  await page.evaluate(() => {
    window.location.hash = "#1";
  });
  await new Promise((r) => setTimeout(r, POST_RASTERIZE_SETTLE_MS));
}
var ADVANCED_LAYERS_SELECTOR = 'section[data-marpit-advanced-background="content"], section[data-marpit-advanced-background="pseudo"]';
async function hideAdvancedLayers(page) {
  await page.evaluate((sel) => {
    document.querySelectorAll(sel).forEach(
      (el) => el.style.setProperty(
        "visibility",
        "hidden",
        "important"
      )
    );
  }, ADVANCED_LAYERS_SELECTOR);
}
async function restoreAdvancedLayers(page) {
  await page.evaluate((sel) => {
    document.querySelectorAll(sel).forEach((el) => el.style.removeProperty("visibility"));
  }, ADVANCED_LAYERS_SELECTOR);
}
async function hideSectionChildren(page, slideIdx) {
  await page.evaluate((n) => {
    const sec = document.querySelector(`section[data-marpit-pagination="${n}"]`);
    if (!sec) return;
    Array.from(sec.children).forEach(
      (el) => el.style.setProperty(
        "visibility",
        "hidden",
        "important"
      )
    );
  }, slideIdx + 1);
}
async function restoreSectionChildren(page, slideIdx) {
  await page.evaluate((n) => {
    const sec = document.querySelector(`section[data-marpit-pagination="${n}"]`);
    if (!sec) return;
    Array.from(sec.children).forEach(
      (el) => el.style.removeProperty("visibility")
    );
  }, slideIdx + 1);
}
function buildFilteredBgJobs(slides) {
  return slides.flatMap((s, i) => {
    const bgs = (s.backgroundImages ?? []).filter((b) => b.cssFilter);
    if (bgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg) => ({
            clip: {
              x: Math.round(bg.x),
              y: Math.round(bg.y),
              width: Math.round(bg.width),
              height: Math.round(bg.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl;
              delete bg.cssFilter;
            }
          })
        ),
        setup: (p) => hideAdvancedLayers(p),
        teardown: (p) => restoreAdvancedLayers(p)
      }
    ];
  });
}
function buildCssFallbackBgJobs(slides) {
  return slides.flatMap((s, i) => {
    const bgs = (s.backgroundImages ?? []).filter((b) => b.fromCssFallback);
    if (bgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg) => ({
            clip: {
              x: 0,
              y: 0,
              width: Math.round(s.width),
              height: Math.round(s.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl;
              delete bg.fromCssFallback;
            }
          })
        ),
        setup: (p, idx) => hideSectionChildren(p, idx),
        teardown: (p, idx) => restoreSectionChildren(p, idx)
      }
    ];
  });
}
function collectFilteredContentImages(elements) {
  const result = [];
  for (const el of elements ?? []) {
    if (el.type === "image" && el.cssFilter) result.push(el);
    if ("children" in el && Array.isArray(el.children)) {
      result.push(...collectFilteredContentImages(el.children));
    }
  }
  return result;
}
function buildFilteredContentImageJobs(slides) {
  return slides.flatMap((s, i) => {
    const imgs = collectFilteredContentImages(s.elements);
    if (imgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: imgs.map(
          (img) => ({
            clip: {
              x: Math.round(img.x),
              y: Math.round(img.y),
              width: Math.round(img.width),
              height: Math.round(img.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              img.src = dataUrl;
              delete img.cssFilter;
            }
          })
        )
      }
    ];
  });
}
function collectRasterizeImages(elements) {
  const result = [];
  for (const el of elements ?? []) {
    if (el.type === "image" && el.rasterize) result.push(el);
    if ("children" in el && Array.isArray(el.children)) {
      result.push(...collectRasterizeImages(el.children));
    }
  }
  return result;
}
function buildRasterizeImageJobs(slides) {
  return slides.flatMap((s, i) => {
    const imgs = collectRasterizeImages(s.elements);
    if (imgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: imgs.map(
          (img) => ({
            clip: {
              x: Math.round(img.x),
              y: Math.round(img.y),
              width: Math.round(img.width),
              height: Math.round(img.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              img.src = dataUrl;
              delete img.rasterize;
            }
          })
        )
      }
    ];
  });
}
function buildPartialBgJobs(slides) {
  return slides.flatMap((s, i) => {
    const bgs = (s.backgroundImages ?? []).filter((b) => {
      if (b.cssFilter || b.fromCssFallback) return false;
      if (b.url.startsWith("data:")) return false;
      const isFullSlide = b.x <= 1 && b.y <= 1 && Math.abs(b.width - s.width) <= 2 && Math.abs(b.height - s.height) <= 2;
      return !isFullSlide;
    });
    if (bgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg) => ({
            clip: {
              x: Math.round(bg.x),
              y: Math.round(bg.y),
              width: Math.round(bg.width),
              height: Math.round(bg.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl;
            }
          })
        ),
        setup: (p) => hideAdvancedLayers(p),
        teardown: (p) => restoreAdvancedLayers(p)
      }
    ];
  });
}
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  generateNativePptx
});
