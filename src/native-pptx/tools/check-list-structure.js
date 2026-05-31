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
    // Section 90 = "Slide 89: Bullet items with code blocks interleaved"
    const s = document.querySelector('section[id="90"]');
    if (!s) return { err: 'section 90 not found' };
    
    const lis = s.querySelectorAll('li');
    const liInfo = Array.from(lis).map((li, i) => ({
      idx: i,
      directChildren: Array.from(li.children).map(c => c.tagName.toLowerCase()),
      hasMarpPre: li.querySelector('marp-pre') !== null,
      text: li.textContent.substring(0, 50),
    }));
    
    // Also check the overall structure
    const ul = s.querySelector('ul, ol');
    const ulChildren = ul ? Array.from(ul.children).map(c => c.tagName.toLowerCase()) : [];
    
    return { liInfo, ulChildren };
  });

  console.log(JSON.stringify(r, null, 2));
  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
