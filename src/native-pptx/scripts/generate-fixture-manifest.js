#!/usr/bin/env node
const fs = require('node:fs')
const path = require('node:path')

const root = path.resolve(__dirname, '..', '..', '..')
const markdownPath = path.join(
  root,
  'src',
  'native-pptx',
  'test-fixtures',
  'pptx-export.md',
)
const outputPath = path.join(
  path.dirname(markdownPath),
  'fixture-manifest.json',
)
const markdown = fs.readFileSync(markdownPath, 'utf8')
const headings = []
let inFence = false
for (const line of markdown.split(/\r?\n/)) {
  if (/^```/.test(line)) {
    inFence = !inFence
    continue
  }
  if (!inFence && /^# /.test(line)) headings.push(line.slice(2).trim())
}

function classify(title) {
  const value = title.toLowerCase()
  const categories = []
  if (/emoji|cjk|symbol|character|japanese|superscript|subscript/.test(value))
    categories.push('typography')
  if (/badge|pill|step|circle|tag/.test(value)) categories.push('badge')
  if (/code|syntax|typescript|sql|indent/.test(value)) categories.push('code')
  if (/list|bullet|ordered|nested/.test(value)) categories.push('list')
  if (/table|cell|colspan|rowspan/.test(value)) categories.push('table')
  if (/image|background|svg|mermaid|math|katex|filter/.test(value))
    categories.push('media')
  if (/grid|flex|container|card|box-shadow|layout/.test(value))
    categories.push('layout')
  if (/header|footer|paginate|pseudo/.test(value))
    categories.push('decoration')
  if (/wrap|overflow|fitting|auto-scal|nowrap|constrained/.test(value))
    categories.push('overflow')
  if (/display contents/.test(value)) categories.push('transparent-wrapper')
  return categories.length > 0 ? categories : ['basic']
}

const slides = headings.map((title, index) => {
  const match = title.match(/^Slide\s+(\d+)(b)?(?:\s*:\s*|\s*$)/i)
  const sequence = match
    ? `F${match[1].padStart(3, '0')}${(match[2] ?? '').toUpperCase()}`
    : 'F001'
  return {
    page: index + 1,
    id: sequence,
    title,
    categories: classify(title),
  }
})

if (slides.length !== 96) {
  throw new Error(`Expected 96 fixture slides, found ${slides.length}`)
}
if (new Set(slides.map((slide) => slide.id)).size !== slides.length) {
  throw new Error('Fixture slide IDs must be unique')
}

fs.writeFileSync(outputPath, JSON.stringify({ slides }, null, 2) + '\n')
console.log(`Generated ${slides.length} fixture manifest entries`)
