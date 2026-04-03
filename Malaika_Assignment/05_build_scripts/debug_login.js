require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 500 });
  const page = await browser.newPage();

  await page.goto(process.env.UWI_BASE_URL + '/login/index.php');
  await page.waitForLoadState('networkidle');
  await page.screenshot({ path: 'step1_login.png' });
  console.log('Step1 URL:', page.url());

  await page.fill('#usernameUserInput', process.env.UWI_USERNAME);
  await page.screenshot({ path: 'step2_username.png' });

  await page.fill('#password', process.env.UWI_PASSWORD);
  await page.screenshot({ path: 'step3_password.png' });

  await page.click('button[type=submit]');
  await page.waitForLoadState('networkidle');
  await page.screenshot({ path: 'step4_after_submit.png' });
  console.log('Step4 URL:', page.url());

  // Check for errors
  const error = await page.$('.alert-danger, .error, #errorMessage, [class*="error"]');
  if (error) {
    const errText = await error.innerText();
    console.log('Error on page:', errText);
  }

  // Wait a bit to see if there's a redirect
  await page.waitForTimeout(3000);
  await page.screenshot({ path: 'step5_final.png' });
  console.log('Final URL:', page.url());

  await browser.close();
})();
