# 下取り追跡番号の抽出（GitHub Actions + Playwright）

短縮URL（c.gle）をブラウザで開き、リダイレクト先の日本郵便URLから `reqCodeNo1`（追跡番号）を取得します。GAS の UrlFetchApp では取れない場合に使います。

## 使い方

### 1. リポジトリを GitHub に push

- このフォルダ（`.github/workflows` と `tracking-resolver`）を GitHub のリポジトリに push する。

### 2. 手動でワークフローを実行

1. GitHub のリポジトリで **Actions** タブを開く。
2. 左から **「Resolve tracking URL」** を選ぶ。
3. **「Run workflow」** をクリック。
4. **short_url** に、メールに含まれる c.gle のURLを貼る（例: `https://c.gle/ANiao5qCX3...`）。
5. （任意）**trade_in_id** に下取りID（例: `702843461200`）を入れる。
6. **「Run workflow」** で実行。

### 3. 結果の見方

- 実行が終わったら、実行したワークフローを開く。
- **「Resolve URL and extract tracking number」** のログに  
  `TRACKING_NUMBER=679514591593` のように出ます。
- **「Output result」** のステップでも追跡番号が表示されます。

その番号を tracking2 シートの D 列に手動で入れてもよいです。

## ローカルで試す

```bash
cd tracking-resolver
npm install
npx playwright install chromium
SHORT_URL="https://c.gle/あなたのURL" node resolve.js
```

## GAS と連携する場合（任意）

1. GAS で「Web アプリとしてデプロイ」し、URL を取得する。
2. その URL を GitHub の **Settings → Secrets and variables → Actions** で `GAS_WEB_APP_URL` として登録する。
3. ワークフロー実行時に **trade_in_id** を入力すると、結果をその Web アプリに送り、スプレッドシートを更新できるようにする（要：code.gs に doGet の処理を追加）。
