/**
 * 短縮URL (c.gle) をブラウザで開き、リダイレクト先の日本郵便URLから reqCodeNo1（追跡番号）を抽出する
 * 環境変数: SHORT_URL (必須), TRADE_IN_ID (任意), GAS_WEB_APP_URL (任意・GASに結果を送る)
 */
const { chromium, request } = require('playwright');

function normalizeShortUrl_(url) {
  if (!url) return '';
  let s = String(url).trim();
  // メールHTML断片が混ざるケース（...style=, target= など）を切り落とす
  s = s.replace(/\s+/g, '');
  s = s.replace(/&(amp;)?/g, '&');
  s = s.replace(/["'>].*$/, '');
  s = s.replace(/(?:style|target|class|rel|aria-[a-z-]+)=.*$/i, '');
  // c.gle のURL本体だけを再抽出
  const m = s.match(/https?:\/\/c\.gle\/[A-Za-z0-9._~!$&'()*+,;=:@%\/?-]+/i);
  return m ? m[0] : s;
}

function extractTrackingNumber_(text) {
  if (!text) return '';
  let m = String(text).match(/reqCodeNo1=(\d{10,14})/);
  if (m) return m[1];
  m = String(text).match(/お問い合わせ番号[\s\S]{0,120}?(\d{4}-\d{4}-\d{4})/);
  if (m) return m[1].replace(/-/g, '');
  return '';
}

async function resolveByHttp_(shortUrl) {
  const api = await request.newContext({
    ignoreHTTPSErrors: true,
    extraHTTPHeaders: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
  });
  let url = shortUrl;
  for (let i = 0; i < 10; i++) {
    const res = await api.get(url, { failOnStatusCode: false, maxRedirects: 0 });
    const code = res.status();
    if (code >= 300 && code < 400) {
      const loc = res.headers()['location'];
      if (!loc) break;
      url = loc.startsWith('http') ? loc : new URL(loc, url).toString();
      const n = extractTrackingNumber_(url);
      if (n) {
        await api.dispose();
        return { trackingNumber: n, finalUrl: url };
      }
      continue;
    }
    const body = await res.text();
    let n = extractTrackingNumber_(body) || extractTrackingNumber_(url);
    if (n) {
      await api.dispose();
      return { trackingNumber: n, finalUrl: url };
    }
    const jp = body.match(/https?:\/\/trackings\.post\.japanpost\.jp[^\s"'<>]+/i);
    if (jp) {
      n = extractTrackingNumber_(jp[0]);
      if (n) {
        await api.dispose();
        return { trackingNumber: n, finalUrl: jp[0] };
      }
    }
    break;
  }
  await api.dispose();
  return { trackingNumber: '', finalUrl: url };
}

async function main() {
  const shortUrl = normalizeShortUrl_(process.env.SHORT_URL || '');
  const tradeInId = (process.env.TRADE_IN_ID || '').trim();
  const gasWebAppUrl = (process.env.GAS_WEB_APP_URL || '').trim();

  if (!shortUrl || shortUrl.length < 10) {
    console.error('SHORT_URL を設定してください（例: https://c.gle/...）');
    process.exit(1);
  }
  console.log('入力URL長: ' + shortUrl.length + ', tail=' + shortUrl.slice(-30));

  let trackingNumber = '';
  let finalUrlForLog = shortUrl;

  // まずHTTPでリダイレクト追跡（最速で安定）
  try {
    const byHttp = await resolveByHttp_(shortUrl);
    if (byHttp.finalUrl) finalUrlForLog = byHttp.finalUrl;
    if (byHttp.trackingNumber) trackingNumber = byHttp.trackingNumber;
  } catch (e) {
    console.warn('HTTP追跡失敗:', e.message);
  }

  // 取れなかったときだけブラウザで再試行
  const browser = await chromium.launch({ headless: true });
  try {
    if (!trackingNumber) {
      const context = await browser.newContext({
        userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        viewport: { width: 1280, height: 720 },
        ignoreHTTPSErrors: true
      });
      const page = await context.newPage();
      await page.goto(shortUrl, { waitUntil: 'load', timeout: 30000 });
      try {
        await page.waitForURL(/reqCodeNo1=|japanpost/, { timeout: 15000 });
      } catch (_) {
        await page.waitForTimeout(3000);
      }
      finalUrlForLog = page.url() || finalUrlForLog;
      trackingNumber = extractTrackingNumber_(finalUrlForLog);
      if (!trackingNumber) {
        const body = await page.content();
        trackingNumber = extractTrackingNumber_(body);
        if (!trackingNumber) {
          const href = body.match(/https?:\/\/trackings\.post\.japanpost\.jp[^\s"'<>]+/i);
          if (href) trackingNumber = extractTrackingNumber_(href[0]);
        }
      }
    }
    await browser.close();
  } catch (e) {
    console.error('Playwright エラー:', e.message);
    await browser.close();
    process.exit(1);
  }

  console.log('最終URL: ' + finalUrlForLog.substring(0, 220) + (finalUrlForLog.length > 220 ? '...' : ''));

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
