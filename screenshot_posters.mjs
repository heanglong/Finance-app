import puppeteer from 'puppeteer';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));

const drinks = [
  { id: 'p1', name: 'Passion Cream' },
  { id: 'p2', name: 'Green Tea with Milk' },
  { id: 'p3', name: 'Sweet and Sour Tamarind Tea' },
  { id: 'p4', name: 'Passion Soda' },
  { id: 'p5', name: 'Red Lemon Tea' },
  { id: 'p6', name: 'Taro with Milk' },
  { id: 'p7', name: 'Iced Coffee with Milk' },
];

const htmlPath = resolve(__dirname, 'public/posters.html');
const fileUrl = `file://${htmlPath}`;

const browser = await puppeteer.launch({
  args: ['--no-sandbox', '--disable-setuid-sandbox'],
  headless: true,
});

const page = await browser.newPage();
await page.setViewport({ width: 1800, height: 1200, deviceScaleFactor: 2 });

await page.goto(fileUrl, { waitUntil: 'networkidle0', timeout: 60000 });

// Extra wait for images and fonts
await new Promise(r => setTimeout(r, 3000));

for (const drink of drinks) {
  const el = await page.$(`#${drink.id}`);
  if (!el) { console.log(`Not found: #${drink.id}`); continue; }
  const out = resolve(__dirname, `public/${drink.name}.png`);
  await el.screenshot({ path: out, type: 'png' });
  console.log(`Saved: ${drink.name}.png`);
}

await browser.close();
console.log('Done!');
