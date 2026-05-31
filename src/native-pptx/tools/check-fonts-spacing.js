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

  const r = await page.evaluate(() => {
    const results = {};
    
    // Slide 87 (code block) - check font
    const s87 = document.querySelector('section[id="87"]');
    if (s87) {
      const code = s87.querySelector('code');
      const marpPre = s87.querySelector('marp-pre');
      results.slide87 = {
        codeFontFamily: code ? getComputedStyle(code).fontFamily : null,
        preFontFamily: marpPre ? getComputedStyle(marpPre).fontFamily : null,
      };
    }
    
    // Slide 22 (long paragraph) - check font and spacing
    const s22 = document.querySelector('section[id="22"]');
    if (s22) {
      const h1 = s22.querySelector('h1');
      const paragraphs = s22.querySelectorAll('p');
      results.slide22 = {
        h1Font: h1 ? getComputedStyle(h1).fontFamily : null,
        h1FontSize: h1 ? getComputedStyle(h1).fontSize : null,
        h1LineHeight: h1 ? getComputedStyle(h1).lineHeight : null,
        paragraphs: Array.from(paragraphs).map(p => {
          const st = getComputedStyle(p);
          const r = p.getBoundingClientRect();
          return {
            fontFamily: st.fontFamily,
            fontSize: st.fontSize,
            lineHeight: st.lineHeight,
            marginTop: st.marginTop,
            marginBottom: st.marginBottom,
            top: r.top,
            bottom: r.bottom,
            height: r.height,
            text: p.textContent.substring(0, 30),
          };
        }),
      };
    }
    
    // Slide 33 (emoji) - check font  
    const s33 = document.querySelector('section[id="33"]');
    if (s33) {
      const li = s33.querySelector('li');
      results.slide33 = {
        liFontFamily: li ? getComputedStyle(li).fontFamily : null,
        liFontSize: li ? getComputedStyle(li).fontSize : null,
        liLineHeight: li ? getComputedStyle(li).lineHeight : null,
      };
    }
    
    // Slide 41 (gradient) - check spacing
    const s41 = document.querySelector('section[id="41"]');
    if (s41) {
      const h2 = s41.querySelector('h2');
      const ps = s41.querySelectorAll('p');
      const lis = s41.querySelectorAll('li');
      results.slide41 = {
        h2LineHeight: h2 ? getComputedStyle(h2).lineHeight : null,
        h2MarginBottom: h2 ? getComputedStyle(h2).marginBottom : null,
        pLineHeight: ps[0] ? getComputedStyle(ps[0]).lineHeight : null,
        pMarginTop: ps[0] ? getComputedStyle(ps[0]).marginTop : null,
        pMarginBottom: ps[0] ? getComputedStyle(ps[0]).marginBottom : null,
        liLineHeight: lis[0] ? getComputedStyle(lis[0]).lineHeight : null,
      };
    }
    
    return results;
  });

  console.log(JSON.stringify(r, null, 2));
  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
