# Timesheet Tool

Google カレンダーから指定月の予定を抽出し、勤務表フォーマットを **Google スプレッドシート**へ自動転記する Python スクリプトです。
**macOS launchd によるバッチ自動実行に対応**しています。

- 休憩時間自動控除・合計行・土日色分け・ヘッダー装飾
- 終日イベント (`date`) / 時間指定イベント (`dateTime`) 両対応
- **ページネーション** 完全対応（`nextPageToken`）
- Sheets API 60 req/min 制限を **65 秒待機**で回避
- 実働時間は跨日対応式 `IF(C<B,(C+1)-B,C-B)*24-D`
- 取得件数 / API 呼び出し回数をログ出力（例 : `📆 西武 : 358 events / 2 request(s)`）
- **バッチ実行対応**: 毎日8:00に当月分を自動処理

---

## 1. 仮想環境のセットアップ

### 仮想環境の作成と有効化

<pre><code>```
# 仮想環境を作成
python -m venv venv

# 仮想環境を有効化（macOS/Linux）
source venv/bin/activate

# 仮想環境を有効化（Windows）
venv\Scripts\activate
```</code></pre>

### 仮想環境の無効化

<pre><code>```
# 仮想環境を無効化
deactivate
```</code></pre>

---

## 2. インストール

仮想環境を有効化した状態で、必要なパッケージをインストールします：

<pre><code>
pip install
pandas pytz gspread gspread_dataframe gspread-formatting
google-api-python-client google-auth google-auth-oauthlib
</code></pre>

---

## 3. Google Cloud 設定

1. 新規プロジェクトを作成  
2. Calendar API と Google Sheets API を有効化  
3. OAuth 同意画面 ▸ 外部 ▸ テスト公開  
4. OAuth 2.0 クライアント ID（デスクトップ）を作成し、`credentials.json` をスクリプトと同じフォルダに配置  
5. 初回実行時にブラウザ認証が走り、`token.json` が自動保存されます。

---

## 4. config.csv の書式

<pre><code>
キーワード1,キーワード2,...,スプレッドシートID 
例）Hoge,ホゲ,12w3e4r5t6y7u8i9o0drftgyhuji
</code></pre>

- 行末がスプレッドシート ID（URL の `/d/` と `/edit` の間）  
- それ以前はイベントタイトルに含めるキーワード（大文字小文字無視・複数可）

### 設定例

現在の設定例（`config.csv`）:
```
VIS,vis,ヴィス,1l0-zLC4cOJmcMiZaB4BFp5LdOkVuNc6yglUlcIt_XSE
```

この設定では：
- カレンダーのイベントタイトルに「VIS」「vis」「ヴィス」のいずれかが含まれる予定を抽出
- 指定されたスプレッドシートIDのシートに勤務表として出力

---

## 5. 使い方

### 手動実行

<pre><code>```
# 当月を自動処理（引数なし）
python Timesheet_Tool.py

# 指定月を処理
python Timesheet_Tool.py 202504

# 対話モード（従来の動作）
python Timesheet_Tool.py --interactive

# ドライラン（スプレッドシートへの書き込みなし）
DRY_RUN=1 python Timesheet_Tool.py
```</code></pre>

### バッチ実行（launchd）

毎日 8:00 に当月分を自動処理するよう設定されています。

<pre><code>```
# バッチをインストール
./scripts/install-launchd.sh

# バッチをアンインストール
./scripts/uninstall-launchd.sh

# 確認
launchctl list | grep timesheettool
```</code></pre>

- 月初 0 時 UTC 〜 翌月 1 日 0 時 UTC の予定を全件取得
- キーワードごとにシートへ出力
- シート完了ごとに 65 秒待機
- ログは `logs/timesheet.log` に出力

例:
📆 hoge会社 : 358 events / 2 request(s) 

---

## 6. オプション

### 環境変数

| 変数名         | 既定値 | 説明                                               |
|----------------|--------|----------------------------------------------------|
| `DEBUG_EVENTS` | 1      | 0 にすると各イベントのデバッグ出力を無効化         |
| `DRY_RUN`      | 0      | 1 にするとスプレッドシートへの書き込みをスキップ   |

### スクリプト内定数

| 変数名     | 既定値 | 説明                           |
|------------|--------|--------------------------------|
| `WAIT_SEC` | 65     | シート書き込み後の待機秒数     |

---

## 7. よくある質問

| 症状                       | 対処方法                                                         |
|----------------------------|------------------------------------------------------------------|
| Quota exceeded (429)       | `WAIT_SEC` を延ばす / 同時に処理するシート数を減らす              |
| Permission denied (403)    | 対象シートに認証アカウントを「編集者」として共有する             |
| 終日予定が抜ける           | 本ツールは `date` 対応済み。タイトルにキーワードが含まれているか確認 |
| 月末が欠ける               | ページネーション実装済み。スクリプトを最新版に更新               |
| バッチが動かない           | `launchctl list \| grep timesheettool` で登録確認。ログを確認     |

---

## 8. ディレクトリ構成

```
TimesheetTool/
├── Timesheet_Tool.py      # メインスクリプト
├── config.csv             # キーワード＆スプレッドシートID設定
├── credentials.json       # Google OAuth認証情報
├── token.json             # OAuthトークン（自動生成）
├── .env.example           # 環境変数テンプレート
├── readme.md              # このファイル
├── launchd/
│   └── com.timesheettool.batch.plist  # launchd設定
├── scripts/
│   ├── install-launchd.sh    # バッチインストール
│   └── uninstall-launchd.sh  # バッチアンインストール
└── logs/
    ├── timesheet.log         # 実行ログ
    └── timesheet.error.log   # エラーログ
```
