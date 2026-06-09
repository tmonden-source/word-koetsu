# Officeアドイン集（Word校閲 + Excel予測家計表）

Claude API を使った2つのOfficeアドインです。中継サーバーには **Netlify Functions** を使い、
APIキーを安全に隠します。1つのNetlifyサイト・1つのAPIキーで両方とも動きます。

1. **Word 日本語校閲** … 誤字脱字を Claude で校閲し、変更履歴（赤字）として反映
2. **Excel 予測家計表** … 実績の家計表から、小規模個人再生の認可後予測家計表を新シートに作成

配布するファイルは用途ごとに別です。
- Word版を使う人 … `src/manifest.xml`
- Excel版を使う人 … `src/manifest-excel.xml`
（登録手順はどちらも同じ。共有フォルダに置いて各自Wordまたはエクセルで登録）

---

## Word 日本語校閲アドイン

Word文書の誤字脱字を Claude API で校閲し、**変更履歴（Track Changes）** として反映します。

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

## Excel 予測家計表アドイン

実績の家計表シートから、小規模個人再生の「認可後1か月分の予測家計表」を新しいシートに作成します。

### 使い方
1. Excelで実績の家計表ブックを開く（収入＝A列費目/B列金額、支出＝C列費目/D列金額の体裁）
2. 「ホーム」タブ →「予測家計表」ボタン → 右にパネルが開く
3. 予測の元にする実績月のシートにチェック（複数月を選ぶと精度が上がる）
4. 「再生計画に基づく月々の返済額」を入力（不明なら空欄可）
5. 「予測表を作成」を押す → 新しいシート「予測家計表(AI作成)」が追加される

### 特徴
- 各項目は実績の傾向（平均・季節変動）から概算を算出。弁護士費用など認可後に発生しない一時費用は除外
- 入力した返済額を「再生計画に基づく返済」として支出に計上
- **計上漏れの定期支出を自動補完**：ガソリン代/駐車場があれば自動車税・車検・自動車保険、
  住宅ローンがあれば固定資産税・火災保険などを概算で仮入力し、項目名に「（要確認：…）」と注記
- 金額はすべて概算。「要確認」項目は必ず実額を確認のこと

### 注意
- 家計表の体裁が大きく異なる場合（列の並びが違うなど）は読み取れないことがあります。
  その場合は `src/predict.js` の `readSheet`（列の対応）を調整してください。
- AIの算出はあくまで下書きです。提出前に必ず内容を精査してください。

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

---

## 証拠説明書の作成アドイン（Word）

訴状（Word）を開いた状態でアドインのボタンを押すと、訴状中で引用されている証拠（甲号証・乙号証）を抽出し、証拠説明書の下書きをExcelで生成します。

### 仕組み
1. 訴状（Word）を開く → リボン「証拠説明書」→「証拠説明書の作成」ボタン
2. アドインが本文を読み取り、中継関数 `evidence-list` 経由でClaudeが証拠を抽出
3. 証拠番号・標目・原本／写し・作成日・作成名義人・立証趣旨を画面に一覧表示
4. 「証拠説明書（Excel）をダウンロード」ボタンで、既存フォーマットに差し込んだ .xlsx をダウンロード

### 構成ファイル
- `netlify/functions/evidence-list.js` … 訴状解析の中継関数
- `src/evidence.html` / `src/evidence.js` … タスクパネル（証拠説明書テンプレートを内蔵）
- `src/manifest-evidence.xml` … Word用マニフェスト（GUIDは校閲アドインと別）

### 配布
校閲アドインと同じ手順で、`manifest-evidence.xml` をWordの信頼できるアドインカタログに登録します。

### 注意
出力は下書きです。標目・作成名義人・作成日・原本／写しの別は、実際の証拠物を確認のうえ確定してください。「要確認」の注記がある項目は特にご注意ください。
