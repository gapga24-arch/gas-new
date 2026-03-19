/**
 * 短縮URL (c.gle) をブラウザで開き、リダイレクト先の日本郵便URLから reqCodeNo1（追跡番号）を抽出する
 * 環境変数: SHORT_URL (必須), TRADE_IN_ID (任意), GAS_WEB_APP_URL (任意・GASに結果を送る)
 */
const { chromium } = require('playwright');

async function main() {
  const shortUrl = (process.env.SHORT_URL || '').trim();
  const tradeInId = (process.env.TRADE_IN_ID || '').trim();
  const gasWebAppUrl = (process.env.GAS_WEB_APP_URL || '').trim();

  if (!shortUrl || shortUrl.length < 10) {
    console.error('SHORT_URL を設定してください（例: https://c.gle/...）');
    process.exit(1);
  }

  let trackingNumber = '';
  const browser = await chromium.launch({ headless: true });
  try {
    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; rv:109.0) Gecko/20100101 Firefox/115.0'
    });
    const page = await context.newPage();

    await page.goto(shortUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(2000);

    const finalUrl = page.url();
    const match = finalUrl.match(/reqCodeNo1=(\d+)/) || finalUrl.match(/jp&reqCodeNo1=(\d+)/);
    if (match) {
      trackingNumber = match[1];
    }
    if (!trackingNumber) {
      const body = await page.content();
      const m = body.match(/reqCodeNo1=(\d{12,14})/) || body.match(/jp&reqCodeNo1=(\d+)/);
      if (m) trackingNumber = m[1];
    }

    await browser.close();
  } catch (e) {
    console.error('Playwright エラー:', e.message);
    await browser.close();
    process.exit(1);
  }

  if (trackingNumber) {
    console.log('TRACKING_NUMBER=' + trackingNumber);
    if (tradeInId) console.log('TRADE_IN_ID=' + tradeInId);
    if (process.env.GITHUB_OUTPUT) {
      const fs = require('fs');
      const delim = (s) => s.replace(/%/g, '%25').replace(/\r/g, '%0D').replace(/\n/g, '%0A');
      fs.appendFileSync(process.env.GITHUB_OUTPUT, `tracking_number<<OUTPUT\n${trackingNumber}\nOUTPUT\n`);
      if (tradeInId) fs.appendFileSync(process.env.GITHUB_OUTPUT, `trade_in_id<<OUTPUT\n${tradeInId}\nOUTPUT\n`);
    }

    if (gasWebAppUrl && tradeInId) {
      const url = gasWebAppUrl.replace(/\?.*$/, '') +
        '?tracking=' + encodeURIComponent(trackingNumber) +
        '&tradeInId=' + encodeURIComponent(tradeInId);
      try {
        const res = await fetch(url);
        console.log('GAS Web App 更新: ' + res.status);
      } catch (e) {
        console.warn('GAS Web App 呼び出し失敗:', e.message);
      }
    }
  } else {
    console.error('追跡番号を取得できませんでした。最終URLを確認してください。');
    process.exit(1);
  }
}

main();
