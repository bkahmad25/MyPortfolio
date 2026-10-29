const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');
const { PDFDocument } = require('pdf-lib');

const CHROME_PATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';

async function getPdfPageCount(filePath) {
  const bytes = fs.readFileSync(filePath);
  const pdfDoc = await PDFDocument.load(bytes);
  return pdfDoc.getPageCount();
}

async function renderResumePdfs() {
  console.log('Launching headless Chrome for PDF generation...');
  const browser = await puppeteer.launch({
    executablePath: CHROME_PATH,
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--font-render-hinting=none'],
  });

  const page = await browser.newPage();
  const htmlPath = path.resolve('Bilal_Ahmed_Resume.html');
  const fileUrl = 'file:///' + htmlPath.replace(/\\/g, '/');

  console.log('Loading resume HTML from:', fileUrl);
  await page.goto(fileUrl, { waitUntil: 'networkidle0' });

  // 1. Render Strict 1-Page ATS Version
  console.log('\n--- Rendering Strict 1-Page ATS Resume ---');
  await page.evaluate(() => {
    switchMode('1page');
  });
  // Brief delay to allow CSS transitions
  await new Promise(r => setTimeout(r, 400));

  const pdf1Path = path.resolve('Bilal_Ahmed_Resume_1Page.pdf');
  await page.pdf({
    path: pdf1Path,
    format: 'A4',
    printBackground: true,
    preferCSSPageSize: true,
    margin: { top: '8mm', right: '11mm', bottom: '8mm', left: '11mm' },
  });
  const count1 = await getPdfPageCount(pdf1Path);
  const size1 = fs.statSync(pdf1Path).size;
  console.log(`Generated: Bilal_Ahmed_Resume_1Page.pdf | Pages: ${count1} | Size: ${(size1 / 1024).toFixed(1)} KB`);

  // 2. Render Balanced 2-Page FinTech Version
  console.log('\n--- Rendering Balanced 2-Page FinTech Resume ---');
  await page.evaluate(() => {
    switchMode('2page');
  });
  await new Promise(r => setTimeout(r, 400));

  const pdf2Path = path.resolve('Bilal_Ahmed_Resume_2Page.pdf');
  await page.pdf({
    path: pdf2Path,
    format: 'A4',
    printBackground: true,
    preferCSSPageSize: true,
    margin: { top: '10mm', right: '13mm', bottom: '10mm', left: '13mm' },
  });
  const count2 = await getPdfPageCount(pdf2Path);
  const size2 = fs.statSync(pdf2Path).size;
  console.log(`Generated: Bilal_Ahmed_Resume_2Page.pdf | Pages: ${count2} | Size: ${(size2 / 1024).toFixed(1)} KB`);

  await browser.close();

  // 3. Update default Bilal_Ahmed_Resume.pdf to the pristine 1-page ATS standard
  fs.copyFileSync(pdf1Path, path.resolve('Bilal_Ahmed_Resume.pdf'));
  console.log('Updated Bilal_Ahmed_Resume.pdf (copied from 1-Page ATS standard).');

  return { count1, count2 };
}

renderResumePdfs().then(res => {
  console.log('\n=======================================');
  console.log(`VERIFICATION RESULT:`);
  console.log(`1-Page ATS: ${res.count1} Page(s) [Target: 1] - ${res.count1 === 1 ? 'SUCCESS' : 'WARNING'}`);
  console.log(`2-Page FinTech: ${res.count2} Page(s) [Target: 2] - ${res.count2 === 2 ? 'SUCCESS' : 'WARNING'}`);
  console.log('=======================================');
}).catch(err => {
  console.error('Error rendering PDFs:', err);
  process.exit(1);
});
