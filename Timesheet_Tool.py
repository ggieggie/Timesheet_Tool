"""
Timesheet_Tool.py   •   2025-04-24
────────────────────────────────────────────────────────
  ● 対象年月を YYYYMM で入力（例 202504）または自動で当月を処理
  ● config.csv の各行「キーワード…,スプレッドシートID」を順に処理
  ● 終日(date) / 通常(dateTime) イベント両対応
  ● ページネーション対応（nextPageToken）
  ● 取得件数・API 呼び出し回数を集計表示
  ● E列実働 = IF(C<B,(C+1)-B,C-B)*24-D   ← 跨日対応
  ● シートごとに 65 秒待機で Sheets 60 req/min 制限を回避

  バッチ実行:
    python Timesheet_Tool.py               # 当月を自動処理
    python Timesheet_Tool.py 202504        # 指定月を処理
    python Timesheet_Tool.py --interactive # 対話モード
    DRY_RUN=1 python Timesheet_Tool.py     # ドライラン（書き込みなし）
"""

import csv, os, sys, re, logging, time, argparse
from datetime import datetime, timedelta
from collections import defaultdict
from pathlib import Path
import pandas as pd
import pytz

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import googleapiclient.discovery
import gspread
from gspread_dataframe import set_with_dataframe
from gspread_formatting import (
    CellFormat, Color, TextFormat,
    format_cell_range, set_frozen, set_column_width
)

# ───────── Config ─────────
DEBUG_EVENTS = os.environ.get('DEBUG_EVENTS', '1') == '1'
DRY_RUN      = os.environ.get('DRY_RUN', '0') == '1'
SCOPES       = ['https://www.googleapis.com/auth/calendar.readonly',
                'https://www.googleapis.com/auth/spreadsheets']
SCRIPT_DIR   = Path(__file__).resolve().parent
TOKEN_FILE   = str(SCRIPT_DIR / 'token.json')
CONFIG_FILE  = str(SCRIPT_DIR / 'config.csv')
CREDS_FILE   = str(SCRIPT_DIR / 'credentials.json')
LOG_DIR      = SCRIPT_DIR / 'logs'
JST          = pytz.timezone('Asia/Tokyo')
WAIT_SEC     = 65
DATE_FMT     = '%Y/%m/%d'

# ───────── Logging Setup ─────────
LOG_DIR.mkdir(exist_ok=True)
log_file = LOG_DIR / 'timesheet.log'
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('%(message)s'))

logger = logging.getLogger('timesheet')
logger.setLevel(logging.INFO)
logger.addHandler(file_handler)
logger.addHandler(console_handler)

SAT_FONT = CellFormat(textFormat=TextFormat(foregroundColor=Color(0.18,0.50,1.0)))
SUN_FONT = CellFormat(textFormat=TextFormat(foregroundColor=Color(1.00,0.25,0.25)))
HEAD_FMT = CellFormat(backgroundColor=Color(0.86,0.90,0.96),
                      horizontalAlignment='CENTER',
                      textFormat=TextFormat(bold=True,foregroundColor=Color(0.14,0.14,0.35)))
NUM_FMT  = CellFormat(numberFormat={'type':'NUMBER','pattern':'0.00'})
TAG_RE   = re.compile(r'\[[^\]]+\]\s*')

googleapiclient.discovery.logger.setLevel(logging.WARNING)

# ───────── Utils ─────────
def to_rfc3339_z(dt): return dt.astimezone(pytz.UTC).strftime('%Y-%m-%dT%H:%M:%SZ')

def authenticate():
    """OAuth認証（トークン自動リフレッシュ対応）"""
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # トークンが無効または期限切れの場合、リフレッシュを試みる
    if creds and creds.expired and creds.refresh_token:
        try:
            logger.info('🔄 トークンをリフレッシュ中...')
            creds.refresh(Request())
            with open(TOKEN_FILE, 'w') as f:
                f.write(creds.to_json())
            logger.info('✅ トークンをリフレッシュしました')
        except Exception as e:
            logger.error('❌ トークンのリフレッシュに失敗: %s', e)
            creds = None

    # トークンがない場合は新規認証（対話が必要）
    if not creds or not creds.valid:
        if not sys.stdin.isatty():
            logger.error('❌ トークンが無効です。対話モードで再認証してください: python Timesheet_Tool.py --interactive')
            sys.exit(1)
        flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as f:
            f.write(creds.to_json())

    return creds

def get_current_month():
    """当月のYYYYMMを返す"""
    today = datetime.now(JST)
    return today.strftime('%Y%m')

# ───────── Calendar fetch ─────────
def get_events(svc, start, end, keywords):
    rec = defaultdict(list)
    page_token = None
    total_cnt  = 0
    req_cnt    = 0

    while True:
        resp = svc.events().list(
            calendarId='primary',
            timeMin=to_rfc3339_z(start),
            timeMax=to_rfc3339_z(end),
            singleEvents=True,
            orderBy='startTime',
            pageToken=page_token
        ).execute()

        req_cnt += 1
        for ev in resp.get('items', []):
            title = ev.get('summary', '')
            if not any(k.lower() in title.lower() for k in keywords):
                continue

            # dateTime / date 両対応
            if 'dateTime' in ev['start']:
                s = datetime.fromisoformat(ev['start']['dateTime'])
                e = datetime.fromisoformat(ev['end'  ]['dateTime'])
            else:  # 終日
                s = datetime.fromisoformat(ev['start']['date'] + 'T00:00:00+00:00')
                e = datetime.fromisoformat(ev['end'  ]['date'] + 'T00:00:00+00:00')

            if DEBUG_EVENTS:
                logger.debug('%s %s', s.astimezone(JST).strftime('%Y-%m-%dT%H:%M'), title)

            rec[s.astimezone(JST).strftime(DATE_FMT)].append(
                {'start': s, 'end': e, 'title': title}
            )
            total_cnt += 1

        page_token = resp.get('nextPageToken')
        if not page_token:
            break

    logger.info('📆 %s : %d events / %d request(s)', keywords[0], total_cnt, req_cnt)
    return rec

# ───────── Summarize ─────────
def summarize(rec, y, m):
    cols=['日付','in (time)','out (time)','休憩（h）','時間（h）','稼働内容']
    rows=[]
    for d,evs in sorted(rec.items()):
        s_min=min(e['start'] for e in evs).astimezone(JST)
        e_max=max(e['end'  ] for e in evs).astimezone(JST)
        span=(e_max-s_min).total_seconds()/3600
        work=sum((e['end']-e['start']).total_seconds()/3600 for e in evs)
        out_disp='24:00' if (e_max.hour==0 and e_max.minute==0 and e_max.date()!=s_min.date()) \
                         else e_max.strftime('%H:%M')
        rows.append([d, s_min.strftime('%H:%M'), out_disp,
                     round(max(0,span-work),2), '', TAG_RE.sub('',', '.join(e['title'] for e in evs))])

    df_evt = pd.DataFrame(rows, columns=cols)
    last   = (datetime(y,m,28)+timedelta(days=4)).replace(day=1)-timedelta(days=1)
    base   = pd.date_range(f'{y}-{m:02}-01', last)
    df     = (pd.DataFrame({'日付':base.strftime(DATE_FMT)})
              .merge(df_evt, on='日付', how='left')
              .fillna({'in (time)':'','out (time)':'','休憩（h）':0,
                       '時間（h）':'','稼働内容':''}))
    df.loc[len(df)] = ['合計','','','', '', '']
    return df

# ───────── Sheets output ─────────
def write_sheet(creds, df, sid, name):
    gc=gspread.authorize(creds)
    sh=gc.open_by_key(sid)
    try: ws=sh.worksheet(name); ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws=sh.add_worksheet(title=name, rows='400', cols='20')

    set_with_dataframe(ws,df)

    n=len(df)-1
    formulas=[[f'=IF(C{r}<B{r},(C{r}+1)-B{r},C{r}-B{r})*24-D{r}'] for r in range(2,2+n)]
    ws.update(formulas, f'E2:E{n+1}', value_input_option='USER_ENTERED')
    ws.update_acell(f'E{n+2}', f'=SUM(E2:E{n+1})')

    format_cell_range(ws, f'E2:E{n+2}', NUM_FMT)
    format_cell_range(ws, 'A1:F1', HEAD_FMT); set_frozen(ws, rows=1)

    for r,d in enumerate(df['日付'][:-1], start=2):
        wd=datetime.strptime(d,DATE_FMT).weekday()
        fmt=SAT_FONT if wd==5 else SUN_FONT if wd==6 else None
        if fmt: format_cell_range(ws,f'A{r}',fmt)

    set_column_width(ws,'A',105); set_column_width(ws,'F',320)
    logger.info('✅ %s → %s', name, sid)

# ───────── Countdown ─────────
def countdown(sec, interactive=False):
    if interactive:
        for s in range(sec,0,-1):
            print(f'\r⏳ 次のシートまで {s:02d} 秒待機…', end='', flush=True)
            time.sleep(1)
        print('\r⏳ 待機終了。次を処理します！   ')
    else:
        logger.info('⏳ 次のシートまで %d 秒待機…', sec)
        time.sleep(sec)

# ───────── main ─────────
def main():
    parser = argparse.ArgumentParser(description='Google Calendar → タイムシート自動生成')
    parser.add_argument('month', nargs='?', help='対象年月 YYYYMM（省略時は当月）')
    parser.add_argument('--interactive', '-i', action='store_true', help='対話モード')
    args = parser.parse_args()

    # 対象年月の決定
    if args.interactive:
        ym = input('対象年月 (YYYYMM): ').strip()
    elif args.month:
        ym = args.month
    else:
        ym = get_current_month()
        logger.info('📅 対象年月（自動）: %s', ym)

    if len(ym) != 6 or not ym.isdigit():
        logger.error('YYYYMM 6桁で入力してください: %s', ym)
        sys.exit(1)

    y, m = int(ym[:4]), int(ym[4:])
    start = datetime(y, m, 1, tzinfo=pytz.UTC)
    end = (start.replace(day=28) + timedelta(days=4)).replace(day=1, tzinfo=pytz.UTC)

    logger.info('🚀 タイムシート処理開始: %s', ym)
    if DRY_RUN:
        logger.info('⚠️  DRY_RUN モード: スプレッドシートへの書き込みはスキップ')

    creds = authenticate()
    cal_svc = build('calendar', 'v3', credentials=creds)

    rows = [r for r in csv.reader(open(CONFIG_FILE, encoding='utf-8')) if r]
    total = len(rows)
    success_count = 0
    error_count = 0

    for i, (*kw, sid) in enumerate(rows, 1):
        kw = [k.strip() for k in kw if k.strip()]
        try:
            df = summarize(get_events(cal_svc, start, end, kw), y, m)
            if DRY_RUN:
                logger.info('🔍 [DRY_RUN] %s: %d 日分のデータ', kw[0], len(df) - 1)
            else:
                write_sheet(creds, df, sid, ym)
            success_count += 1
            if i < total:
                countdown(WAIT_SEC, interactive=args.interactive)
        except Exception as e:
            logger.error('❌ 行%d でエラー: %s\n   ↳ %s', i, kw + [sid], e)
            error_count += 1

    logger.info('🏁 処理完了: 成功 %d / 失敗 %d', success_count, error_count)
    sys.exit(0 if error_count == 0 else 1)

if __name__ == '__main__':
    main()
