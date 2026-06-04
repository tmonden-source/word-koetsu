# 日本語校閲アドイン（Word + Netlify + Claude API）

Word文書の誤字脱字を Claude API で校閲し、**変更履歴（Track Changes）** として反映する
Officeアドインです。中継サーバーには **Netlify Functions** を使い、APIキーを安全に隠します。

```
[Word アドイン] ──→ [Netlify Functions] ──→ [Claude API]
   変更履歴に反映      APIキーを安全に保管
```

---

## フォルダ構成

```
word-koetsu/
├─ netlify.toml                  … Netlifyの設定
├─ netlify/functions/
│   └─ proofread.js              … 中継サーバー（APIキーはここで使う）
└─ src/                          … 公開される静的ファイル
    ├─ manifest.xml              … ★配布するファイル
    ├─ taskpane.html             … アドインの画面
    ├─ taskpane.js               … アドインの動作
    └─ assets/icon-32.png, icon-80.png
```

---

## セットアップ手順

### 0. 事前に必要なもの
- Netlify の無料アカウント（https://netlify.com）
- Claude API キー（https://console.anthropic.com で取得）
- GitHub アカウント（任意。なくてもドラッグ&ドロップで可）

### 1. GitHub にアップロード

このフォルダの中身をそのままリポジトリにします。

```bash
cd word-koetsu
git init
git add .
git commit -m "初回コミット: 日本語校閲アドイン"
git branch -M main
git remote add origin https://github.com/あなたのユーザー名/word-koetsu.git
git push -u origin main
```

> `.gitignore` により `.env`（APIキー）や `node_modules` は除外されます。
> APIキーがコミットされないことを `git status` で必ず確認してください。

GitHub上で「New repository」を先に作成し、発行されたURLを `origin` に指定します。

### 2. Netlify に GitHub 連携でデプロイ
1. Netlify にログイン →「Add new site」→「Import an existing project」
2. 「GitHub」を選び、先ほどのリポジトリを選択
3. ビルド設定はそのまま（`netlify.toml` が自動で読まれる）→「Deploy」
4. デプロイ完了後、`https://○○○.netlify.app` のようなURLが発行される

以降はGitHubに `git push` するたびに、Netlifyが自動で再デプロイします。

### 2. APIキーを環境変数に登録（重要）
Netlifyのサイト管理画面で:
`Site configuration → Environment variables → Add a variable`
- Key: `ANTHROPIC_API_KEY`
- Value: あなたのClaude APIキー
- Scope: **Functions** にチェック

登録後、`Deploys → Trigger deploy → Deploy site` で再デプロイして反映。

> APIキーはサーバー側にのみ保存され、アドイン（利用者のPC）には一切渡りません。

### 3. URLを書き換える
発行されたNetlifyのURL（例: `word-koetsu.netlify.app`）に合わせて、
以下2ファイルの `YOUR-SITE-NAME.netlify.app` をすべて置換します。
- `src/taskpane.js` の `PROOFREAD_ENDPOINT`
- `src/manifest.xml`（複数箇所）

### 4. manifest の GUID を発行
`manifest.xml` の `<Id>` を一意のGUIDに置き換えます。生成方法:
- Mac/Linux: ターミナルで `uuidgen`
- Windows: PowerShellで `[guid]::NewGuid()`
- もしくは https://guidgenerator.com

書き換えたら `git push` すると、Netlifyが自動で再デプロイします。

---

## 動作確認（自分のPCでテスト＝サイドロード）

### Windows
1. 任意の共有フォルダ（例 `\\PC名\共有`）に `manifest.xml` を置く
2. Word →「ファイル」→「オプション」→「セキュリティセンター」
   →「セキュリティセンターの設定」→「信頼できるアドインカタログ」
3. 上記フォルダのパスを追加し「メニューに表示する」にチェック → Word再起動
4. 「ホーム」タブ →「アドイン」→「個人用アドイン」→ 共有フォルダ → 「日本語校閲」を選択

### Mac
1. 下記フォルダに `manifest.xml` をコピー:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
   （`wef` フォルダが無ければ作成）
2. Word再起動 →「ホーム」タブに「日本語校閲」ボタンが出る

ボタンを押すとパネルが開き、「校閲を実行」で文書全体が校閲され、
修正が**赤字の変更履歴**として入ります。「校閲」タブで承諾/元に戻すが可能です。

---

## チーム配布の方法

配布とは「manifest.xml を全員のWordに届けること」です。3つの方法があります。

### ① Microsoft 365 管理センターで集中配置（推奨・大人数向け）
会社がMicrosoft 365(Business/Enterprise)で、あなたが管理者権限を持つ場合に最適。
1. https://admin.microsoft.com にログイン
2. 「設定」→「統合アプリ」→「アプリのアップロード」→「カスタムアプリ」
3. `manifest.xml` をアップロード
4. 配布対象（全員 / 特定グループ）を選択して展開
5. 数時間〜最大24時間で対象者のWordに自動で表示される

利点: 各自の作業ゼロ。追加・削除・更新を管理者が一括管理。

### ② 共有フォルダ / SharePoint 配置（管理者権限なし・小規模向け）
1. ネットワーク共有フォルダに `manifest.xml` を置く
2. 各メンバーが「動作確認（Windows）」の手順2〜4を一度だけ実施
   （＝共有フォルダを「信頼できるカタログ」に登録）
3. 以降は「個人用アドイン」からボタン一つで読み込める

利点: 管理者権限不要。欠点: 各自で初回設定が必要。〜10人程度向け。

### ③ サイドロード（テスト専用）
上記「動作確認」の手順。配布には使いません。

---

## カスタマイズ

- **校閲の方針を変える**: `netlify/functions/proofread.js` の `systemPrompt` を編集
  （例: 「ですます調に統一」「専門用語は変更しない」など指示を追加）
- **モデルを変える**: 同ファイルの `model` を変更（精度重視なら上位モデル）
- **同じ語を全部直す**: `taskpane.js` の `applyCorrections` で
  `searchResults.items[0]` のループを全件に変更
- **OpenAIに変えたい**: `proofread.js` のfetch先とリクエスト形式を差し替え

---

## トラブルシューティング

| 症状 | 対処 |
|---|---|
| 「APIキーが設定されていません」 | 環境変数 `ANTHROPIC_API_KEY` の登録漏れ。Scopeに Functions を含めて再デプロイ |
| ボタンを押しても何も起きない | `taskpane.js` の `PROOFREAD_ENDPOINT` のURL誤り |
| アドインが一覧に出ない | manifestのURLが全てhttps、GUIDが一意か確認。Word再起動 |
| 変更履歴にならず直接書き換わる | 既に履歴オンなら正常。Wordの「校閲」→「変更履歴の記録」状態を確認 |
| 校閲結果が出ない/解析失敗 | 文書が長すぎる可能性。20000字制限を `proofread.js` で調整 |
```
