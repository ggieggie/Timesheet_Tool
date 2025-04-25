# Timesheet Tool

Google カレンダーから指定月の予定を抽出し、勤務表フォーマットを **Google スプレッドシート**へ自動転記する Python スクリプトです。

- 休憩時間自動控除・合計行・土日色分け・ヘッダー装飾  
- 終日イベント (`date`) / 時間指定イベント (`dateTime`) 両対応  
- **ページネーション** 完全対応（`nextPageToken`）  
- Sheets API 60 req/min 制限を **65 秒待機**で回避  
- 実働時間は跨日対応式 `IF(C<B,(C+1)-B,C-B)*24-D`  
- 取得件数 / API 呼び出し回数をデバッグ出力（例 : `📆 西武 : 358 events / 2 request(s)`）

---

## 1. インストール

<pre><code>```
pip install
pandas pytz gspread gspread_dataframe gspread-formatting
google-api-python-client google-auth google-auth-oauthlib
```</code></pre>

---

## 2. Google Cloud 設定

1. 新規プロジェクトを作成  
2. Calendar API と Google Sheets API を有効化  
3. OAuth 同意画面 ▸ 外部 ▸ テスト公開  
4. OAuth 2.0 クライアント ID（デスクトップ）を作成し、`credentials.json` をスクリプトと同じフォルダに配置  
5. 初回実行時にブラウザ認証が走り、`token.json` が自動保存されます。

---

## 3. config.csv の書式

<pre><code>```
キーワード1,キーワード2,...,スプレッドシートID 
```</code></pre>

- 行末がスプレッドシート ID（URL の `/d/` と `/edit` の間）  
- それ以前はイベントタイトルに含めるキーワード（大文字小文字無視・複数可）

---

## 4. 使い方

<pre><code>```
python Timesheet_Tool.py
対象年月 (YYYYMM): 202504
```</code></pre>

- 月初 0 時 UTC 〜 翌月 1 日 0 時 UTC の予定を全件取得  
- キーワードごとにシートへ出力  
- シート完了ごとに 65 秒カウントダウンを表示  
- 取得件数 / API コール回数を集計表示  

例:
📆 hoge会社 : 358 events / 2 request(s) 
📆 huga会社 : 402 events / 3 request(s)

---

## 5. オプション

| 変数名         | 定義場所   | 既定値 | 説明                                        |
|----------------|------------|--------|---------------------------------------------|
| `WAIT_SEC`     | ソース先頭 | 65     | シート書き込み後の待機秒数                  |
| `DEBUG_EVENTS` | ソース先頭 | False  | True にすると各イベントを逐次表示する       |

---

## 6. よくある質問

| 症状                       | 対処方法                                                         |
|----------------------------|------------------------------------------------------------------|
| Quota exceeded (429)       | `WAIT_SEC` を延ばす / 同時に処理するシート数を減らす              |
| Permission denied (403)    | 対象シートに認証アカウントを「編集者」として共有する             |
| 終日予定が抜ける           | 本ツールは `date` 対応済み。タイトルにキーワードが含まれているか確認 |
| 月末が欠ける               | ページネーション実装済み。スクリプトを最新版に更新               |
