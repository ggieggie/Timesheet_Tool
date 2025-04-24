"""
Timesheet_Tool.py   •   2025-04-24
────────────────────────────────────────────────────────
  ● 対象年月を YYYYMM で入力（例 202504）
  ● config.csv の各行「キーワード…,スプレッドシートID」を順に処理
  ● 終日(date) / 通常(dateTime) イベント両対応
  ● ページネーション対応（nextPageToken）
  ● 取得件数・API 呼び出し回数を集計表示
      📆 Plan-B : 402 events / 3 request(s)
  ● E列実働 = IF(C<B,(C+1)-B,C-B)*24-D   ← 跨日対応
  ● シートごとに 65 秒待機で Sheets 60 req/min 制限を回避
"""

import csv, os, sys, re, logging, time
from datetime import datetime, timedelta
from collections import defaultdict
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
DEBUG_EVENTS = True                 # True で各予定を逐次表示
SCOPES      = ['https://www.googleapis.com/auth/calendar.readonly',
               'https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE  = 'token.json'
JST         = pytz.timezone('Asia/Tokyo')
WAIT_SEC    = 65
DATE_FMT    = '%Y/%m/%d'

SAT_FONT = CellFormat(textFormat=TextFormat(foregroundColor=Color(0.18,0.50,1.0)))
SUN_FONT = CellFormat(textFormat=TextFormat(foregroundColor=Color(1.00,0.25,0.25)))
HEAD_FMT = CellFormat(backgroundColor=Color(0.86,0.90,0.96),
                      horizontalAlignment='CENTER',
                      textFormat=TextFormat(bold=True,foregroundColor=Color(0.14,0.14,0.35)))
NUM_FMT  = CellFormat(numberFormat={'type':'NUMBER','pattern':'0.00'})
TAG_RE   = re.compile(r'\[[^\]]+\]\s*')

logging.basicConfig(level=logging.INFO)
googleapiclient.discovery.logger.setLevel(logging.DEBUG)

# ───────── Utils ─────────
def to_rfc3339_z(dt): return dt.astimezone(pytz.UTC).strftime('%Y-%m-%dT%H:%M:%SZ')

def authenticate():
    if os.path.exists(TOKEN_FILE):
        from google.oauth2.credentials import Credentials
        return Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    flow=InstalledAppFlow.from_client_secrets_file('credentials.json',SCOPES)
    creds=flow.run_local_server(port=0)
    with open(TOKEN_FILE,'w') as f: f.write(creds.to_json())
    return creds

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
                print(s.astimezone(JST).strftime('%Y-%m-%dT%H:%M'), title)

            rec[s.astimezone(JST).strftime(DATE_FMT)].append(
                {'start': s, 'end': e, 'title': title}
            )
            total_cnt += 1

        page_token = resp.get('nextPageToken')
        if not page_token:
            break

    print(f'📆 {keywords[0]} : {total_cnt} events / {req_cnt} request(s)')
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
    print(f'✅ {name} → {sid}')

# ───────── Countdown ─────────
def countdown(sec):
    for s in range(sec,0,-1):
        print(f'\r⏳ 次のシートまで {s:02d} 秒待機…', end='', flush=True)
        time.sleep(1)
    print('\r⏳ 待機終了。次を処理します！   ')

# ───────── main ─────────
def main():
    ym=input('対象年月 (YYYYMM): ').strip()
    if len(ym)!=6 or not ym.isdigit():
        print('YYYYMM 6桁で入力してください'); sys.exit(1)
    y,m=int(ym[:4]),int(ym[4:])
    start=datetime(y,m,1,tzinfo=pytz.UTC)
    end=(start.replace(day=28)+timedelta(days=4)).replace(day=1,tzinfo=pytz.UTC)

    creds=authenticate()
    cal_svc=build('calendar','v3',credentials=creds)

    rows=[r for r in csv.reader(open('config.csv',encoding='utf-8')) if r]
    total=len(rows)
    for i,( *kw,sid) in enumerate(rows,1):
        kw=[k.strip() for k in kw if k.strip()]
        try:
            df=summarize(get_events(cal_svc,start,end,kw),y,m)
            write_sheet(creds,df,sid,ym)
            if i<total: countdown(WAIT_SEC)
        except Exception as e:
            print(f'❌ 行{i} でエラー: {kw+[sid]}\n   ↳ {e}')

if __name__=='__main__':
    main()
