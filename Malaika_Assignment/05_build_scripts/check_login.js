require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  await page.goto(process.env.UWI_BASE_URL + '/login/index.php');
  await page.waitForLoadState('networkidle');
  await page.screenshot({ path: 'login_page.png' });

  const inputs = await page.$$eval('input', els =>
    els.map(e => ({ id: e.id, name: e.name, type: e.type, placeholder: e.placeholder }))
  );
  console.log('Inputs:', JSON.stringify(inputs, null, 2));

  const buttons = await page.$$eval('button, input[type=submit]', els =>
    els.map(e => ({ id: e.id, text: e.innerText || e.value, type: e.type }))
  );
  console.log('Buttons:', JSON.stringify(buttons));

  await browser.close();
})();
