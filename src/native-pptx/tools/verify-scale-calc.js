/**
 * Verify computeAutoScaleFactor would work for marp-pre elements.
 */
const puppeteer = require('puppeteer-core');
const path = require('path');

const htmlPath = path.resolve(__dirname, '../test-fixtures/slides-ci.html');
const htmlUrl = 'file:///' + htmlPath.replace(/\\/g, '/');

async function main() {
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    headless: 'new',
  });

  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 720 });
  await page.goto(htmlUrl, { waitUntil: 'networkidle0' });

  const results = await page.evaluate(() => {
    const sectionIds = ['87', '91', '92'];
    return sectionIds.map(id => {
      const section = document.querySelector(`section[id="${id}"]`);
      if (!section) return { id, error: 'section not found' };
      
      const marpPre = section.querySelector('marp-pre');
      if (!marpPre) return { id, error: 'no marp-pre' };
      
      const style = getComputedStyle(marpPre);
      const rect = marpPre.getBoundingClientRect();
      const code = marpPre.querySelector('code');
      if (!code) return { id, error: 'no code' };
      
      const codeRect = code.getBoundingClientRect();
      const text = code.textContent || '';
      const lines = text.split('\n');
      
      // Simulate extractCodeRuns line count
      const numLines = lines.length;
      
      // Simulate computeAutoScaleFactor
      const lineHeight = parseFloat(style.lineHeight) || 0;
      const paddingTop = parseFloat(style.paddingTop) || 0;
      const paddingBottom = parseFloat(style.paddingBottom) || 0;
      const contentHeight = rect.height - paddingTop - paddingBottom;
      const naturalHeight = numLines * lineHeight;
      const wouldScale = naturalHeight > contentHeight * 1.05;
      const scaleFactor = wouldScale ? contentHeight / naturalHeight : 1.0;
      
      // What the result would be
      const originalFontSize = parseFloat(style.fontSize);
      const adjustedFontSize = originalFontSize * scaleFactor;
      const adjustedFontPt = adjustedFontSize * 0.75;
      
      // What the ideal target is (from actual visual rendering)
      const idealScale = codeRect.height / naturalHeight;
      const idealFontSize = originalFontSize * idealScale;
      const idealFontPt = idealFontSize * 0.75;
      
      return {
        id,
        numLines,
        lineHeight: Math.round(lineHeight * 100) / 100,
        rectHeight: Math.round(rect.height * 100) / 100,
        paddingTop: Math.round(paddingTop * 100) / 100,
        paddingBottom: Math.round(paddingBottom * 100) / 100,
        contentHeight: Math.round(contentHeight * 100) / 100,
        naturalHeight: Math.round(naturalHeight * 100) / 100,
        wouldScale,
        scaleFactor: Math.round(scaleFactor * 10000) / 10000,
        originalFontSize,
        adjustedFontSize: Math.round(adjustedFontSize * 100) / 100,
        adjustedFontPt: Math.round(adjustedFontPt * 100) / 100,
        idealScale: Math.round(idealScale * 10000) / 10000,
        idealFontSize: Math.round(idealFontSize * 100) / 100,
        idealFontPt: Math.round(idealFontPt * 100) / 100,
        error_percent: Math.round(Math.abs(scaleFactor - idealScale) / idealScale * 10000) / 100,
      };
    });
  });

  console.log(JSON.stringify(results, null, 2));
  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
