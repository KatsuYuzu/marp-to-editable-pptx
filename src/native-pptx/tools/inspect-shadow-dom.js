/**
 * Inspect the shadow DOM of marp-pre to find the transform scale.
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
      
      const sr = marpPre.shadowRoot;
      if (!sr) return { id, error: 'no shadow root' };
      
      // Get shadow DOM style content
      const styleEl = sr.querySelector('style');
      const styleContent = styleEl?.textContent?.substring(0, 500) || '';
      
      // Check marp-auto-scaling element
      const autoScaling = sr.querySelector('marp-auto-scaling');
      if (!autoScaling) return { id, error: 'no marp-auto-scaling' };
      
      const asStyle = getComputedStyle(autoScaling);
      const asRect = autoScaling.getBoundingClientRect();
      
      // Check deeper: does marp-auto-scaling have a shadow root too?
      const asShadow = autoScaling.shadowRoot;
      
      // Get the slot contents
      const slot = sr.querySelector('slot') || autoScaling.querySelector('slot');
      
      // Check if marp-auto-scaling has its own shadow DOM
      let deeperInfo = null;
      if (asShadow) {
        const deepChildren = Array.from(asShadow.children).map(el => ({
          tag: el.tagName,
          style: el.getAttribute('style')?.substring(0, 300) || '',
          computedTransform: getComputedStyle(el).transform,
          computedFontSize: getComputedStyle(el).fontSize,
          computedZoom: getComputedStyle(el).zoom,
          rect: (() => { const r = el.getBoundingClientRect(); return { w: Math.round(r.width*100)/100, h: Math.round(r.height*100)/100 }; })(),
        }));
        const deeperStyle = asShadow.querySelector('style');
        deeperInfo = {
          children: deepChildren,
          styleContent: deeperStyle?.textContent?.substring(0, 500) || '',
          allWithTransform: Array.from(asShadow.querySelectorAll('*')).filter(el => {
            const t = getComputedStyle(el).transform;
            return t && t !== 'none';
          }).map(el => ({
            tag: el.tagName,
            transform: getComputedStyle(el).transform,
            id: el.id,
            cls: el.className,
          })),
        };
      }
      
      return {
        id,
        styleContent,
        marpAutoScaling: {
          fontSize: asStyle.fontSize,
          zoom: asStyle.zoom,
          transform: asStyle.transform,
          display: asStyle.display,
          rect: { w: Math.round(asRect.width*100)/100, h: Math.round(asRect.height*100)/100 },
          hasShadowRoot: !!asShadow,
        },
        deeperInfo,
      };
    });
  });

  console.log(JSON.stringify(results, null, 2));
  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
