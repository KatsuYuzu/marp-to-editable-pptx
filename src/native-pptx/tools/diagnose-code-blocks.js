/**
 * Diagnostic: Extract code block metrics from slides-ci.html via Puppeteer.
 * Shows fontSize, lineHeight, rect dimensions, and computed vs actual rendering.
 */
const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

const htmlPath = path.resolve(__dirname, '../test-fixtures/slides-ci.html');
const htmlUrl = 'file:///' + htmlPath.replace(/\\/g, '/');

// Target slides with code blocks
const TARGET_SLIDES = [87, 90, 91];
// Section IDs in slides-ci.html (these map from Marp's numbered headings)
const TARGET_SECTION_IDS = ['87', '91', '92'];

async function main() {
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    headless: 'new',
    args: ['--no-sandbox'],
  });

  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 720 });
  await page.goto(htmlUrl, { waitUntil: 'networkidle0' });

  // Debug: check how many sections
  const count = await page.evaluate(() => document.querySelectorAll('section').length);
  console.error('Section count:', count);

  // Debug: check target sections for pre elements
  const debug = await page.evaluate(() => {
    const s87 = document.querySelector('section[id="87"]');
    if (!s87) return { error: 'section#87 not found' };
    const html = s87.innerHTML;
    const childTags = Array.from(s87.children).map(c => c.tagName + '.' + (c.className || '').substring(0, 30));
    // Look for pre/code anywhere in the section hierarchy
    const allPre = s87.querySelectorAll('pre');
    const allCode = s87.querySelectorAll('code');
    // Check if there's a marp-auto-scaling wrapper
    const autoScale = s87.querySelectorAll('marp-auto-scaling, [data-auto-scaling]');
    return {
      htmlLength: html.length,
      htmlFirst500: html.substring(0, 500),
      childTags,
      preCount: allPre.length,
      codeCount: allCode.length,
      autoScaleCount: autoScale.length,
    };
  });
  console.error('Section #87 debug:', JSON.stringify(debug, null, 2));

  // Get all slides' code block diagnostics

  // Get all slides' code block diagnostics
  const diagnostics = await page.evaluate((targetSlides) => {
    const sections = document.querySelectorAll('section');
    const results = [];

    for (const slideIdx of targetSlides) {
      const section = sections[slideIdx - 1]; // 0-indexed
      if (!section) {
        results.push({ slide: slideIdx, error: 'section not found' });
        continue;
      }

      const slideRect = section.getBoundingClientRect();
      const pres = section.querySelectorAll('pre');

      for (let i = 0; i < pres.length; i++) {
        const pre = pres[i];
        const preStyle = getComputedStyle(pre);
        const preRect = pre.getBoundingClientRect();

        const code = pre.querySelector('code');
        const codeTarget = code || pre;
        const codeStyle = getComputedStyle(codeTarget);

        // Check for transforms on ancestors
        let transformInfo = [];
        let el = pre;
        while (el && el !== document.documentElement) {
          const s = getComputedStyle(el);
          if (s.transform && s.transform !== 'none') {
            transformInfo.push({
              tag: el.tagName,
              class: el.className?.substring(0, 50),
              transform: s.transform,
            });
          }
          el = el.parentElement;
        }

        // Count lines
        const text = codeTarget.textContent || '';
        const lines = text.split('\n');
        const numLines = lines.length;

        // First few lines to check indentation
        const sampleLines = lines.slice(0, 8).map(l => ({
          text: l.substring(0, 60),
          leadingSpaces: l.match(/^(\s*)/)[1].length,
        }));

        // Compute what dom-walker would extract
        const fontSize = parseFloat(preStyle.fontSize);
        const lineHeight = parseFloat(preStyle.lineHeight);
        const paddingTop = parseFloat(preStyle.paddingTop) || 0;
        const paddingBottom = parseFloat(preStyle.paddingBottom) || 0;
        const contentHeight = preRect.height - paddingTop - paddingBottom;
        const naturalHeight = numLines * lineHeight;
        const wouldScale = naturalHeight > contentHeight * 1.05;
        const scaleFactor = wouldScale ? contentHeight / naturalHeight : 1.0;

        results.push({
          slide: slideIdx,
          preIndex: i,
          preRect: {
            x: Math.round((preRect.left - slideRect.left) * 100) / 100,
            y: Math.round((preRect.top - slideRect.top) * 100) / 100,
            width: Math.round(preRect.width * 100) / 100,
            height: Math.round(preRect.height * 100) / 100,
          },
          slideRect: {
            width: Math.round(slideRect.width * 100) / 100,
            height: Math.round(slideRect.height * 100) / 100,
          },
          preFontSize: fontSize,
          preLineHeight: lineHeight,
          codeFontSize: parseFloat(codeStyle.fontSize),
          paddingTop,
          paddingBottom,
          contentHeight: Math.round(contentHeight * 100) / 100,
          naturalHeight: Math.round(naturalHeight * 100) / 100,
          numLines,
          wouldScale,
          scaleFactor: Math.round(scaleFactor * 10000) / 10000,
          transforms: transformInfo,
          sampleLines,
          // What PPTX would get
          pptx: {
            shapeX_inches: Math.round((preRect.left - slideRect.left) / 96 * 1000) / 1000,
            shapeY_inches: Math.round((preRect.top - slideRect.top) / 96 * 1000) / 1000,
            shapeW_inches: Math.round(preRect.width / 96 * 1000) / 1000,
            shapeH_inches: Math.round(preRect.height / 96 * 1000) / 1000,
            shapeH_pt: Math.round(preRect.height / 96 * 72 * 100) / 100,
            fontSize_pt: Math.round(fontSize * 0.75 * 100) / 100,
            maxFontSizePt: Math.round((preRect.height / 96 * 72) / (numLines * 1.2) * 100) / 100,
          },
        });
      }
    }
    return results;
  }, TARGET_SLIDES);

  console.log(JSON.stringify(diagnostics, null, 2));
  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
