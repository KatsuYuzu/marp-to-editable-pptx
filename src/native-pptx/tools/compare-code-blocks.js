/**
 * Block-level visual comparison: crops individual code blocks from HTML
 * screenshots and PPTX exports, then compares them pixel-by-pixel.
 * 
 * This gives much more granular results than full-slide comparison since
 * small typography differences don't get diluted by the rest of the slide.
 * 
 * Requires: PowerPoint (COM), Chrome, lib/native-pptx.cjs (npm run build)
 * Usage: node src/native-pptx/tools/compare-code-blocks.js [slides-ci.html]
 */
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const puppeteer = require('puppeteer-core');

const htmlPath = path.resolve(
  process.argv[2] || 'src/native-pptx/test-fixtures/slides-ci.html'
);
const outDir = path.resolve('dist/compare-code-blocks');
const pptxPath = path.resolve('dist/compare-code-blocks.pptx');

// PowerPoint COM → PNG export via VBScript
const VBS_EXPORT = `
Set ppt = CreateObject("PowerPoint.Application")
ppt.Visible = False
Set pres = ppt.Presentations.Open(WScript.Arguments(0), True, False, False)
Dim outFolder
outFolder = WScript.Arguments(1)
Dim slideIdx
slideIdx = CInt(WScript.Arguments(2))
pres.Slides(slideIdx).Export outFolder & "\\pptx-slide-" & slideIdx & ".png", "PNG", 1280, 720
pres.Close
ppt.Quit
`;

async function main() {
  if (!fs.existsSync(htmlPath)) {
    console.error('HTML not found:', htmlPath);
    process.exit(1);
  }

  // Setup output directory
  if (fs.existsSync(outDir)) fs.rmSync(outDir, { recursive: true });
  fs.mkdirSync(outDir, { recursive: true });

  // 1. Generate PPTX
  console.log('Generating PPTX...');
  const bundlePath = path.resolve('lib/native-pptx.cjs');
  const { generateNativePptx } = require(bundlePath);
  const buffer = await generateNativePptx({
    htmlPath,
    browserPath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  });
  fs.writeFileSync(pptxPath, buffer);
  console.log('PPTX written:', pptxPath);

  // 2. Extract code block positions from HTML via Puppeteer
  console.log('Extracting code block positions from HTML...');
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    headless: 'new',
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 720 });
  await page.goto('file:///' + htmlPath.replace(/\\/g, '/'), { waitUntil: 'networkidle0' });

  // Find all code blocks with their slide index and position
  const codeBlocks = await page.evaluate(() => {
    const sections = document.querySelectorAll('section');
    const blocks = [];
    
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i];
      const slideRect = section.getBoundingClientRect();
      // Skip nested sections (they appear within parent section areas)
      if (slideRect.width <= 0 || slideRect.height <= 0) continue;
      
      const marpPres = section.querySelectorAll('marp-pre');
      const pres = section.querySelectorAll('pre');
      const allPres = [...marpPres, ...pres];
      
      for (let j = 0; j < allPres.length; j++) {
        const pre = allPres[j];
        // Skip duplicates (pre inside marp-pre would be selected by both)
        if (pre.closest('marp-pre') && pre.tagName !== 'MARP-PRE') continue;
        
        const preRect = pre.getBoundingClientRect();
        if (preRect.width <= 0 || preRect.height <= 0) continue;
        
        blocks.push({
          sectionIndex: i,
          sectionId: section.id,
          blockIndex: j,
          // Clip coordinates relative to slide
          clip: {
            x: Math.max(0, Math.round(preRect.left - slideRect.left)),
            y: Math.max(0, Math.round(preRect.top - slideRect.top)),
            width: Math.min(Math.round(preRect.width), 1280),
            height: Math.min(Math.round(preRect.height), 720),
          },
          // Absolute page coordinates for Puppeteer screenshot
          pageClip: {
            x: Math.round(preRect.left),
            y: Math.round(preRect.top),
            width: Math.round(preRect.width),
            height: Math.round(preRect.height),
          },
        });
      }
    }
    return blocks;
  });

  console.log(`Found ${codeBlocks.length} code blocks`);

  // 3. Screenshot HTML code blocks
  console.log('Screenshotting HTML code blocks...');
  for (const block of codeBlocks) {
    const fileName = `html-s${block.sectionId}-b${block.blockIndex}.png`;
    await page.screenshot({
      path: path.join(outDir, fileName),
      clip: block.pageClip,
    });
  }
  await browser.close();

  // 4. Export PPTX slides via PowerPoint COM
  console.log('Exporting PPTX slides via PowerPoint...');
  const vbsPath = path.join(outDir, 'export-slide.vbs');
  fs.writeFileSync(vbsPath, VBS_EXPORT);

  // Get unique slide indices (1-based for PowerPoint)
  const uniqueSlides = [...new Set(codeBlocks.map(b => b.sectionIndex + 1))];
  
  for (const slideNum of uniqueSlides) {
    try {
      execSync(
        `cscript //nologo "${vbsPath}" "${pptxPath}" "${outDir}" ${slideNum}`,
        { timeout: 30000, stdio: 'pipe' }
      );
    } catch (e) {
      console.error(`  Failed to export slide ${slideNum}:`, e.message?.substring(0, 100));
    }
  }

  // 5. Crop PPTX screenshots to code block regions and compare
  console.log('Comparing code block regions...');
  const { PNG } = require('pngjs');
  const pixelmatch = require('pixelmatch');

  const results = [];
  
  for (const block of codeBlocks) {
    const htmlFile = path.join(outDir, `html-s${block.sectionId}-b${block.blockIndex}.png`);
    const pptxSlideFile = path.join(outDir, `pptx-slide-${block.sectionIndex + 1}.png`);
    
    if (!fs.existsSync(htmlFile) || !fs.existsSync(pptxSlideFile)) {
      results.push({
        slide: block.sectionId,
        block: block.blockIndex,
        status: 'SKIP',
        reason: !fs.existsSync(pptxSlideFile) ? 'PPTX export missing' : 'HTML screenshot missing',
      });
      continue;
    }

    // Read HTML block image
    const htmlPng = PNG.sync.read(fs.readFileSync(htmlFile));
    
    // Read and crop PPTX slide image to the same region
    const pptxSlidePng = PNG.sync.read(fs.readFileSync(pptxSlideFile));
    const { x: cx, y: cy, width: cw, height: ch } = block.clip;
    
    // Ensure crop is within bounds
    const cropW = Math.min(cw, pptxSlidePng.width - cx);
    const cropH = Math.min(ch, pptxSlidePng.height - cy);
    
    if (cropW <= 0 || cropH <= 0) {
      results.push({ slide: block.sectionId, block: block.blockIndex, status: 'SKIP', reason: 'crop out of bounds' });
      continue;
    }
    
    // Crop PPTX image
    const pptxCropped = new PNG({ width: cropW, height: cropH });
    for (let row = 0; row < cropH; row++) {
      for (let col = 0; col < cropW; col++) {
        const srcIdx = ((cy + row) * pptxSlidePng.width + (cx + col)) * 4;
        const dstIdx = (row * cropW + col) * 4;
        pptxCropped.data[dstIdx] = pptxSlidePng.data[srcIdx];
        pptxCropped.data[dstIdx + 1] = pptxSlidePng.data[srcIdx + 1];
        pptxCropped.data[dstIdx + 2] = pptxSlidePng.data[srcIdx + 2];
        pptxCropped.data[dstIdx + 3] = pptxSlidePng.data[srcIdx + 3];
      }
    }
    
    // Save cropped PPTX for inspection
    const pptxCroppedFile = path.join(outDir, `pptx-s${block.sectionId}-b${block.blockIndex}.png`);
    fs.writeFileSync(pptxCroppedFile, PNG.sync.write(pptxCropped));
    
    // Resize HTML image to match crop dimensions for comparison
    // (HTML screenshot is at exact pixel coordinates, PPTX might have slight size differences)
    const compareW = Math.min(htmlPng.width, cropW);
    const compareH = Math.min(htmlPng.height, cropH);
    
    if (compareW <= 0 || compareH <= 0) {
      results.push({ slide: block.sectionId, block: block.blockIndex, status: 'SKIP', reason: 'size mismatch' });
      continue;
    }

    // Create same-size images for comparison
    const imgA = new PNG({ width: compareW, height: compareH });
    const imgB = new PNG({ width: compareW, height: compareH });
    
    // Copy HTML region
    for (let row = 0; row < compareH; row++) {
      for (let col = 0; col < compareW; col++) {
        const srcIdx = (row * htmlPng.width + col) * 4;
        const dstIdx = (row * compareW + col) * 4;
        imgA.data[dstIdx] = htmlPng.data[srcIdx];
        imgA.data[dstIdx + 1] = htmlPng.data[srcIdx + 1];
        imgA.data[dstIdx + 2] = htmlPng.data[srcIdx + 2];
        imgA.data[dstIdx + 3] = htmlPng.data[srcIdx + 3];
      }
    }
    
    // Copy PPTX region
    for (let row = 0; row < compareH; row++) {
      for (let col = 0; col < compareW; col++) {
        const srcIdx = (row * cropW + col) * 4;
        const dstIdx = (row * compareW + col) * 4;
        imgB.data[dstIdx] = pptxCropped.data[srcIdx];
        imgB.data[dstIdx + 1] = pptxCropped.data[srcIdx + 1];
        imgB.data[dstIdx + 2] = pptxCropped.data[srcIdx + 2];
        imgB.data[dstIdx + 3] = pptxCropped.data[srcIdx + 3];
      }
    }
    
    // Pixelmatch comparison
    const diff = new PNG({ width: compareW, height: compareH });
    const numDiffPixels = pixelmatch(
      imgA.data, imgB.data, diff.data,
      compareW, compareH,
      { threshold: 0.3 }
    );
    
    const totalPixels = compareW * compareH;
    const diffRate = (numDiffPixels / totalPixels * 100);
    
    // Save diff image
    const diffFile = path.join(outDir, `diff-s${block.sectionId}-b${block.blockIndex}.png`);
    fs.writeFileSync(diffFile, PNG.sync.write(diff));
    
    const status = diffRate < 5 ? 'PASS' : diffRate < 15 ? 'WARN' : 'FAIL';
    results.push({
      slide: block.sectionId,
      block: block.blockIndex,
      status,
      diffRate: diffRate.toFixed(2) + '%',
      diffPixels: numDiffPixels,
      totalPixels,
      dimensions: `${compareW}×${compareH}`,
    });
  }

  // 6. Print results
  console.log('\n=== Code Block Comparison Results ===');
  console.log('Threshold: PASS < 5%, WARN < 15%, FAIL >= 15%\n');
  
  for (const r of results) {
    const icon = r.status === 'PASS' ? '✓' : r.status === 'WARN' ? '⚠' : r.status === 'FAIL' ? '✗' : '?';
    console.log(`  ${icon} Slide ${r.slide} block ${r.block}: ${r.status} ${r.diffRate || ''} ${r.reason || ''} ${r.dimensions || ''}`);
  }
  
  const passes = results.filter(r => r.status === 'PASS').length;
  const warns = results.filter(r => r.status === 'WARN').length;
  const fails = results.filter(r => r.status === 'FAIL').length;
  const skips = results.filter(r => r.status === 'SKIP').length;
  console.log(`\nSummary: ${passes} PASS, ${warns} WARN, ${fails} FAIL, ${skips} SKIP`);
  console.log(`Output: ${outDir}`);
  
  // Write JSON results
  fs.writeFileSync(
    path.join(outDir, 'results.json'),
    JSON.stringify(results, null, 2)
  );
}

main().catch(e => { console.error(e); process.exit(1); });
