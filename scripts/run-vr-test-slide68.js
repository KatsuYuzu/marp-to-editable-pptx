#!/usr/bin/env node
/**
 * Run VR test pipeline for pptx-export.md test fixture.
 * Steps: 1) gen-pptx  2) compare-visuals (slide 68 focused)
 */
const { execSync, spawnSync } = require('child_process')
const fs = require('fs')
const path = require('path')

const root = path.resolve(__dirname, '..')
process.chdir(root)

// Ensure images exist in dist/vr-test
const distVr = path.join(root, 'dist', 'vr-test')
const testIcon = path.join(root, 'src', 'native-pptx', 'test-fixtures', 'test-icon.png')
if (!fs.existsSync(path.join(distVr, 'test-icon.png'))) {
  fs.copyFileSync(testIcon, path.join(distVr, 'test-icon.png'))
}
if (!fs.existsSync(path.join(distVr, 'mock-screenshot.png'))) {
  fs.copyFileSync(testIcon, path.join(distVr, 'mock-screenshot.png'))
}
console.log('Images ready')

// Step 1: Generate PPTX
const htmlPath = path.join(distVr, 'test-pptx-export.html')
const pptxPath = path.join(distVr, 'test-pptx-export.pptx')

if (!fs.existsSync(htmlPath)) {
  console.error('HTML not found:', htmlPath)
  process.exit(1)
}

console.log('Generating PPTX...')
const genResult = spawnSync('node', [
  path.join(root, 'src', 'native-pptx', 'tools', 'gen-pptx.js'),
  htmlPath,
  pptxPath
], { encoding: 'utf-8', timeout: 120000, stdio: 'pipe' })

if (genResult.status !== 0) {
  console.error('gen-pptx failed:', genResult.stderr || genResult.stdout)
  process.exit(1)
}
console.log('PPTX generated:', fs.statSync(pptxPath).size, 'bytes')

// Step 2: Compare visuals
console.log('Running compare-visuals...')
const cmpResult = spawnSync('node', [
  path.join(root, 'src', 'native-pptx', 'tools', 'compare-visuals.js'),
  htmlPath,
  pptxPath
], { encoding: 'utf-8', timeout: 600000, stdio: 'pipe' })

if (cmpResult.stdout) console.log(cmpResult.stdout)
if (cmpResult.stderr) console.error(cmpResult.stderr)
console.log('Exit code:', cmpResult.status)
