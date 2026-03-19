# GitHub と連携して下取りの追跡番号を自動でスプシに反映する手順

今までと同じように「1時間おきの GAS 実行」の流れのまま、追跡番号が取れないときだけ **GitHub Actions（Playwright）** に任せて、取得した番号をスプシに書き戻します。

---

## 流れのイメージ

1. **GAS** が1時間おきに動く（今まで通り）
2. 下取り発送メールを処理するとき、**まず GAS 内で短縮URLを解決**（UrlFetchApp）を試す
3. **取れなかった場合** → GAS が **GitHub Actions のワークフローを1回だけ起動**
4. **GitHub Actions** が Playwright で短縮URLを開き、追跡番号を取得
5. GitHub が **GAS の Web アプリ** を呼び、**tracking2 の D列・L列を更新**

---

## 手順一覧

| 順番 | やること |
|------|----------|
| 1 | GitHub にリポジトリを作り、このプロジェクトを push する |
| 2 | GitHub で「Personal Access Token」を作り、GAS の Script Properties に登録する |
| 3 | GAS を「Web アプリ」としてデプロイし、URL を GitHub の Secrets に登録する |
| 4 | GAS の Script Properties に GitHub 用の設定（リポジトリ名・ワークフローID）を入れる |
| 5 | 動作確認する |

---

## 1. GitHub にリポジトリを作って push する

### 1-1. リポジトリを用意する

1. [GitHub](https://github.com) にログインする
2. 右上の **+** → **New repository**
3. 名前は何でもよい（例: `gas-new`）
4. **Create repository** で作成

### 1-2. 手元のフォルダを Git で管理して push する

**PowerShell** で次を実行（パスは自分の環境に合わせてください）。

```powershell
cd C:\Users\81709\gas-new

# まだ Git 初期化していなければ
git init
git add .
git commit -m "Initial: GAS + GitHub Actions 追跡番号取得"

# 自分のリポジトリURLに書き換え（例: https://github.com/あなたのユーザー名/gas-new.git）
git remote add origin https://github.com/あなたのユーザー名/gas-new.git
git branch -M main
git push -u origin main
```

- すでに `git init` や `remote` をしている場合は、`git add .` と `git commit`、`git push` だけ実行すればよいです。
- ブランチ名が `main` でない場合は、`GITHUB_WORKFLOW_ID` のところで「どのブランチで動かすか」を後で合わせます（多くの場合は `main`）。

---

## 2. GitHub の Personal Access Token を作る

GAS から「ワークフローを起動する」ために、GitHub のトークンが必要です。

1. GitHub で **右上のアイコン** → **Settings**
2. 左の一番下 **Developer settings**
3. **Personal access tokens** → **Tokens (classic)** または **Fine-grained tokens**
4. **Generate new token** を押す
   - **Classic** の場合: 名前を付けて、**repo** にチェック、**workflow** にチェック → Generate
   - **Fine-grained** の場合: このリポジトリだけ許可し、Permissions で **Actions: Read and write** を付ける
5. 表示された **トークン（ghp_... など）** をコピーして、どこかに控えておく（あとで GAS に入れます）

---

## 3. GAS の Script Properties にトークンとリポジトリ情報を入れる

1. [Google Apps Script](https://script.google.com) で、このプロジェクト（gas-new）を開く
2. 左の **プロジェクトの設定**（歯車アイコン）
3. **スクリプト プロパティ** で **スクリプト プロパティを追加** を押し、次の3つを追加する

| プロパティ | 値 | 説明 |
|------------|-----|------|
| `GITHUB_TOKEN` | （さっきコピーしたトークン） | GitHub の Personal Access Token |
| `GITHUB_REPO` | `あなたのユーザー名/gas-new` | リポジトリの「所有者/リポジトリ名」 |
| `GITHUB_WORKFLOW_ID` | `resolve-tracking.yml` | ワークフローファイル名（そのまま） |

- すでに `SPREADSHEET_ID` などがある場合は、それに加えて上記を追加します。

---

## 4. GAS を Web アプリとしてデプロイし、URL を GitHub に登録する

GitHub が「追跡番号を取得したあと、スプシを更新する」ために、GAS の **Web アプリのURL** を GitHub に教えます。

### 4-1. GAS で Web アプリをデプロイする

1. Apps Script のエディタで **デプロイ** → **新しいデプロイ**
2. 種類で **ウェブアプリ** を選ぶ
3. 説明は任意（例: 「追跡番号更新用」）
4. **次のユーザーとして実行**: 自分のアカウント
5. **アクセスできるユーザー**: 「全員」でも「自分のみ」でもよい（GitHub から呼ぶだけなら「全員」の方が楽）
6. **デプロイ** を押す
7. **ウェブアプリの URL**（`https://script.google.com/macros/s/.../exec`）をコピーする

### 4-2. GitHub の Secrets に URL を登録する

1. GitHub の **リポジトリのページ** を開く
2. **Settings** → 左の **Secrets and variables** → **Actions**
3. **New repository secret**
4. 名前: `GAS_WEB_APP_URL`
5. 値: さっきコピーした **ウェブアプリの URL そのまま**（`https://script.google.com/.../exec`）
6. **Add secret** で保存

これで、「GitHub Actions が追跡番号を取ったあと、この URL に `?tracking=番号&tradeInId=下取りID` を付けて呼ぶ」と、GAS の `doGet` が tracking2 の D列・L列を更新します。

---

## 5. 動作確認

### 5-1. GAS のメイン処理を1回実行

1. Apps Script で `processOrderEmails` を選んで **実行**
2. **表示** → **ログ** を開く
3. 下取り発送で「追跡番号が取れなかった」ケースがあると、  
   `[GitHub] ワークフロー起動しました tradeInId=...` と出る
4. GitHub の **Actions** タブで、**Resolve tracking URL** が実行されているか確認

### 5-2. GitHub 側で番号が取れているか確認

1. GitHub の **Actions** で、いちばん新しい「Resolve tracking URL」の実行を開く
2. **Resolve URL and extract tracking number** のログに  
   `TRACKING_NUMBER=...` が出ていれば成功
3. **GAS_WEB_APP_URL** を設定していれば、同じ実行の中で GAS の Web アプリが呼ばれ、**tracking2 の D列に追跡番号が入っている**はずです

### 5-3. スプシを確認

- **tracking2** シートの該当行の **D列** に追跡番号、**L列** に「発送」が入っていれば完了です。

---

## まとめ（今までと同じように動かすには）

- **トリガー**: 今まで通り「1時間おき」で `processOrderEmails` を実行すればよいです。
- **流れ**:  
  - GAS が下取り発送メールを処理  
  → 短縮URLはあるが GAS では番号が取れない  
  → そのときだけ GitHub のワークフローが1回起動  
  → GitHub が Playwright で番号を取得  
  → GAS の Web アプリを呼んでスプシに反映  

設定は **Script Properties（GITHUB_*）** と **GitHub の Secrets（GAS_WEB_APP_URL）** と **Web アプリのデプロイ** の3つだけです。

---

## トラブル時

- **「[GitHub] 未設定のためスキップ」**  
  → GITHUB_TOKEN, GITHUB_REPO, GITHUB_WORKFLOW_ID のどれかが未設定。Script Properties を確認。
- **「[GitHub] 起動失敗: 404」**  
  → GITHUB_REPO が `所有者/リポジトリ名` になっているか、ワークフローが `main` ブランチに存在するか確認。  
  → デフォルトブランチが `master` の場合は、code.gs の `triggerGitHubResolve_` 内の `ref: 'main'` を `ref: 'master'` に変更する。
- **GitHub の実行は成功するがスプシが更新されない**  
  → GitHub の Secrets の `GAS_WEB_APP_URL` が「ウェブアプリの URL」そのままか、GAS の `doGet` がデプロイされたウェブアプリに含まれているか確認。
