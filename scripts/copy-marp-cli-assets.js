// Copy marp-cli runtime assets to dist/.
// marp-cli reads bespoke.js and watch.js via fs.readFile(path.resolve(__dirname, file)).
// After webpack bundles the code, __dirname points to dist/, so these files
// must be present alongside dist/extension.js.
const { cpSync, mkdirSync } = require('node:fs')
const path = require('node:path')

const SRC = path.resolve(__dirname, '../node_modules/@marp-team/marp-cli/lib')
const DEST = path.resolve(__dirname, '../dist')
const ASSETS = ['bespoke.js', 'watch.js']

mkdirSync(DEST, { recursive: true })
for (const file of ASSETS) {
  cpSync(path.join(SRC, file), path.join(DEST, file))
  console.log(`Copied: ${file} → dist/${file}`)
}
