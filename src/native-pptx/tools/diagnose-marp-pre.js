/**
 * Diagnostic: Inspect marp-pre elements for auto-scaling behavior.
 */
const puppeteer = require('puppeteer-core');
const path = require('path');

const htmlPath = path.resolve(__dirname, '../test-fixtures/slides-ci.html');
const htmlUrl = 'file:///' + htmlPath.replace(/\\/g, '/');

async function main() {
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    headless: 'new',
    args: ['--no-sandbox'],
  });

  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 720 });
  await page.goto(htmlUrl, { waitUntil: 'networkidle0' });

  const results = await page.evaluate(() => {
    const sections = ['87', '91', '92'].map(id => document.querySelector(`section[id="${id}"]`));
    
    return sections.map((section, i) => {
      if (!section) return { id: ['87', '91', '92'][i], error: 'not found' };
      
      const slideRect = section.getBoundingClientRect();
      const marpPres = section.querySelectorAll('marp-pre');
      const results = [];
      
      for (const marpPre of marpPres) {
        const marpPreRect = marpPre.getBoundingClientRect();
        const marpPreStyle = getComputedStyle(marpPre);
        const code = marpPre.querySelector('code');
        const codeStyle = code ? getComputedStyle(code) : null;
        const codeRect = code ? code.getBoundingClientRect() : null;
        
        // Check for shadow DOM
        const shadowRoot = marpPre.shadowRoot;
        
        // Check transforms on marp-pre and its ancestors up to section
        const transforms = [];
        let el = marpPre;
        while (el && el !== section) {
          const s = getComputedStyle(el);
          if (s.transform && s.transform !== 'none') {
            transforms.push({
              tag: el.tagName,
              transform: s.transform,
              transformOrigin: s.transformOrigin,
            });
          }
          el = el.parentElement;
        }
        
        // Key insight: check if code element's rect is smaller than marp-pre rect
        // That would indicate auto-scaling
        const text = code?.textContent || '';
        const lines = text.split('\n');
        
        results.push({
          sectionId: section.id,
          dataAutoScaling: marpPre.getAttribute('data-auto-scaling'),
          marpPreRect: {
            x: Math.round((marpPreRect.left - slideRect.left) * 100) / 100,
            y: Math.round((marpPreRect.top - slideRect.top) * 100) / 100,
            w: Math.round(marpPreRect.width * 100) / 100,
            h: Math.round(marpPreRect.height * 100) / 100,
          },
          marpPreStyle: {
            fontSize: marpPreStyle.fontSize,
            lineHeight: marpPreStyle.lineHeight,
            display: marpPreStyle.display,
            overflow: marpPreStyle.overflow,
            transform: marpPreStyle.transform,
          },
          codeStyle: codeStyle ? {
            fontSize: codeStyle.fontSize,
            lineHeight: codeStyle.lineHeight,
            display: codeStyle.display,
            whiteSpace: codeStyle.whiteSpace,
            transform: codeStyle.transform,
          } : null,
          codeRect: codeRect ? {
            x: Math.round((codeRect.left - slideRect.left) * 100) / 100,
            y: Math.round((codeRect.top - slideRect.top) * 100) / 100,
            w: Math.round(codeRect.width * 100) / 100,
            h: Math.round(codeRect.height * 100) / 100,
          } : null,
          hasShadowRoot: !!shadowRoot,
          transforms,
          numLines: lines.length,
          slideH: slideRect.height,
          // Ratio of code bounding rect height to marp-pre rect height
          // If < 1.0, content is being scaled down
          scaleRatio: codeRect && marpPreRect.height > 0
            ? Math.round(codeRect.height / marpPreRect.height * 1000) / 1000
            : null,
        });
      }
      return results;
    }).flat();
  });

  console.log(JSON.stringify(results, null, 2));
  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
