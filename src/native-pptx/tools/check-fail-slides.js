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

  // Load the bundle to use extractSlides
  const bundlePath = path.resolve(__dirname, '../../../lib/native-pptx.cjs');
  const { extractSlides } = require(bundlePath);

  // Extract slide data for specific slides
  const allSlides = await extractSlides(page);
  
  // Slide 22 (index 21)
  const s22 = allSlides[21];
  console.log('=== Slide 22 (index 21) ===');
  console.log('Element count:', s22.elements.length);
  for (const el of s22.elements) {
    console.log(`  type=${el.type} y=${el.y?.toFixed(1)} h=${el.height?.toFixed(1)} font=${el.style?.fontFamily?.substring(0,40)} lh=${el.style?.lineHeight?.toFixed(1)} fs=${el.style?.fontSize?.toFixed(1)}`);
    if (el.type === 'text' && el.runs) {
      const textPreview = el.runs.filter(r => !r.breakLine).map(r => r.text).join('').substring(0, 60);
      console.log(`    text: "${textPreview}..."`);
      // Check if there are paragraph breaks
      const breakCount = el.runs.filter(r => r.breakLine).length;
      console.log(`    breakLines: ${breakCount}, totalRuns: ${el.runs.length}`);
    }
  }

  // Slide 33 (index 32)
  const s33 = allSlides[32];
  console.log('\n=== Slide 33 (index 32) ===');
  console.log('Element count:', s33.elements.length);
  for (const el of s33.elements) {
    console.log(`  type=${el.type} y=${el.y?.toFixed(1)} h=${el.height?.toFixed(1)} fs=${el.style?.fontSize?.toFixed(1)} lh=${el.style?.lineHeight?.toFixed(1)}`);
  }

  // Slide 41 (index 40)
  const s41 = allSlides[40];
  console.log('\n=== Slide 41 (index 40) ===');
  console.log('Element count:', s41.elements.length);
  for (const el of s41.elements) {
    console.log(`  type=${el.type} y=${el.y?.toFixed(1)} h=${el.height?.toFixed(1)} fs=${el.style?.fontSize?.toFixed(1)} lh=${el.style?.lineHeight?.toFixed(1)}`);
  }

  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
