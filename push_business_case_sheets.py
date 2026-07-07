#!/usr/bin/env python3
"""
Ka-Minions Business Case — Google Sheets Updater

Creates / updates a tab "Business Case Analysis" in the target spreadsheet
with the full financial model (3 value pillars, ROI, phase roadmap, AI costs).

Auth: OAuth Desktop (token cached in token.json after first browser login)

Dependencies:
  pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client

Setup (first run only):
  1. Go to console.cloud.google.com → APIs & Services → Credentials
  2. Create OAuth 2.0 Client ID → Desktop Application → Download JSON
  3. Rename the file to credentials.json and place it next to this script
  4. Run this script → browser opens → authorize with your Google account
  5. token.json is saved for all future runs (no browser needed again)
"""

import os, json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SPREADSHEET_ID = '1KFEm_CjGwdv7yespJ5TrWLETzQ9ZuWISGwkOTHGhVNw'
SHEET_NAME     = 'Business Case Analysis'
SCOPES         = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE     = os.path.join(os.path.dirname(__file__), 'token.json')
CREDS_FILE     = os.path.join(os.path.dirname(__file__), 'credentials.json')

# ── Auth ──────────────────────────────────────────────────────────────────────

def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDS_FILE):
                raise FileNotFoundError(
                    f'\ncredentials.json not found at: {CREDS_FILE}\n'
                    'Download it from Google Cloud Console:\n'
                    '  console.cloud.google.com → APIs & Services → Credentials\n'
                    '  → Create OAuth 2.0 Client ID → Desktop Application → Download JSON\n'
                    '  → rename to credentials.json → place next to this script'
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as f:
            f.write(creds.to_json())
    return creds


# ── Sheet data ────────────────────────────────────────────────────────────────

# Color helpers (RGB 0-1 scale)
def rgb(r, g, b): return {'red': r/255, 'green': g/255, 'blue': b/255}

DARK_GREEN    = rgb(29,  75,   0)   # #1D4B00
LIME          = rgb(181, 233,  65)  # #B5E941
MID_GREEN     = rgb(96,  159,  40)  # #609F28
OFF_WHITE     = rgb(244, 242, 239)  # #F4F2EF
PURPLE        = rgb(134, 89,  228)  # #8659E4
ORANGE        = rgb(255, 167,  63)  # #FFA73F
TEAL          = rgb(44,  124, 135)  # #2C7C87
WARM_GRAY     = rgb(80,   78,  72)  # #504E48
WHITE         = rgb(255, 255, 255)
LIGHT_GRAY    = rgb(221, 219, 213)  # #DDDBD5

ROWS = [
    # ── Section 0: header ─────────────────────────────────────────────────────
    ['KA-MINIONS — BUSINESS CASE ANALYSIS', '', '', '', '', ''],
    ['IDEAS2IMPACT 2026', '', '', '', '', ''],
    ['', '', '', '', '', ''],

    # ── Section 1: 3 Value Pillars ────────────────────────────────────────────
    ['3 VALUE PILLARS', '', '', '', '', ''],
    ['#', 'Pillar', 'Pain point', 'Ka-Minion response', 'Primary KPI', 'Target'],
    ['1', 'Faster Delivery',
         '30–40% sprint capacity on maintenance, not features',
         'Automate ≤1SP bugs & routine tasks end-to-end via Paperclip',
         'Bug resolution time',
         '3 days → <3 hours (6×)'],
    ['2', 'Zombie Task Closure',
         '22% of Jira tickets stale >30 days — backlog noise',
         'Centinell 24/7: auto-triage, enrich, close or re-prioritise stale tickets',
         '% tickets >30 days',
         '22% → <8% (est.)'],
    ['3', 'On-call Support',
         'Manual triage at all hours — engineer fatigue, slow response',
         'Ka-Minions handle Tier-0/1 incidents; escalate edge cases only',
         'Human on-call interruptions',
         '−30% out-of-hours pages'],
    ['', '', '', '', '', ''],

    # ── Section 2: Pilot financial data ───────────────────────────────────────
    ['PILOT DATA (6-MONTH BACKLOG — TEAMS AZ & MO)', '', '', '', '', ''],
    ['Team', 'Automated Tasks', 'Automated Bugs', 'Hours Freed / Year',
     'Hourly Rate (€80K/1760h)', 'Annual Savings (€)'],
    ['Team AZ', 48, 50, 1168.2, 45.45, 53100.00],
    ['Team MO', 56, 37, 1113.6, 45.45, 50618.18],
    ['TOTAL (2 teams)', '=B13+B14', '=C13+C14', '=D13+D14', '', '=F13+F14'],
    ['', '', '', '', '', ''],

    # ── Section 3: Automation rate assumptions ────────────────────────────────
    ['AUTOMATION RATE ASSUMPTIONS', '', '', '', '', ''],
    ['Issue type', 'Total backlog (est.)', 'Automatable %', 'Automatable volume',
     'Avg hours each', 'Hours/yr freed'],
    ['Bugs (≤1SP)', 87, '80%', '=ROUND(B18*0.80,0)', 6, '=D18*E18'],
    ['BAU Tasks (≤1SP)', 104, '25%', '=ROUND(B19*0.25,0)', 6, '=D19*E19'],
    ['TOTAL', '=B18+B19', '', '=D18+D19', '', '=F18+F19'],
    ['', '', '', '', '', ''],

    # ── Section 4: Phase roadmap + KPIs ───────────────────────────────────────
    ['PHASE ROADMAP & KPIS', '', '', '', '', ''],
    ['Phase', 'Focus', 'Timeline', 'Key KPI', 'Target', 'Teams in scope'],
    ['Phase 1: Hypothesis Validation',
     'Automate repetitive ops within 2 pilot teams',
     '1 Quarter',
     'Bug/task resolution time',
     '3 days → <3 hours',
     'Team AZ, Team MO'],
    ['Phase 2: MVPs Creation',
     'Boost impact across 5 selected teams',
     '2 Quarters',
     'Ka-Minion success rate',
     '70% success rate',
     '5 selected teams'],
    ['Phase 3: Growth & Iterations',
     'AI agents integrated within existing headcounts',
     '4 Quarters',
     'Team adoption',
     '40% teams adopted',
     'Wider KA engineering'],
    ['Phase 4: Scale Up',
     'Full AI agentic teams as integrated team members',
     'Always On',
     'Team adoption + feature delivery',
     '70% adoption / +30% features to prod',
     'All KA (80 engineers)'],
    ['', '', '', '', '', ''],

    # ── Section 5: AI cost model ───────────────────────────────────────────────
    ['AI COST MODEL — PAPERCLIP + CLAUDE CODE', '', '', '', '', ''],
    ['Cost item', 'Phase 1 (one-time)', 'Annual — Phase 1 scale',
     'Annual — Phase 4 scale', 'Notes', ''],
    ['Paperclip setup', '€60K – €80K', '—', '—',
     'Integration, guardrails, audit trail, dashboards', ''],
    ['Paperclip maintenance', '—', '€20K – €30K', '€60K – €90K',
     'License + prompt optimisation + domain re-training', ''],
    ['Claude Code / LLM API', '—', '€2K – €3.5K', '€18K – €27K',
     '~50–100K tokens per task · ~€0.80–€1.50/task', ''],
    ['TOTAL annual AI cost', '—', '~€23K – €34K', '~€78K – €117K', '', ''],
    ['', '', '', '', '', ''],

    # ── Section 6: ROI summary ────────────────────────────────────────────────
    ['ROI SUMMARY', '', '', '', '', ''],
    ['Scenario', 'Annual AI Cost (€)', 'Annual Savings / Value (€)',
     'Net Gain (€)', 'ROI ratio', 'Break-even'],
    ['Phase 1 — 2 pilot teams', 28500, 103718, '=C38-B38', '=C38/B38', 'Q1 pilot'],
    ['Phase 2 — 5 teams (est.)', 34000, 260000, '=C39-B39', '=C39/B39', 'Q2'],
    ['Phase 4 — all KA 80 eng', 100000, 1600000, '=C40-B40', '=C40/B40', 'Year 2'],
    ['', '', '', '', '', ''],

    # ── Section 7: Governance & risk ──────────────────────────────────────────
    ['GOVERNANCE & RISK MITIGATION', '', '', '', '', ''],
    ['Risk', 'Mitigation', '', '', '', ''],
    ['AI hallucinations / vulnerable code in production',
     'Paperclip governed orchestrator: no merge to main without mandatory unit test gate', '', '', '', ''],
    ['Developer friction (fear of replacement)',
     'Position Ka-Minions as digital assistants eliminating boring work; shifts devs to high-value design', '', '', '', ''],
    ['Vague / poorly documented Jira tickets',
     'Mandatory Markdown ticket templates; Ka-Minion auto-rejects & flags under-specified tasks', '', '', '', ''],
    ['LLM token budget overrun',
     '10–20% of AI maintenance budget dedicated to prompt optimisation & token budget supervision', '', '', '', ''],
]


# ── Sheet helpers ─────────────────────────────────────────────────────────────

def ensure_sheet(service, spreadsheet_id, sheet_name):
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in meta['sheets']:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    body = {'requests': [{'addSheet': {'properties': {'title': sheet_name}}}]}
    resp = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=body
    ).execute()
    return resp['replies'][0]['addSheet']['properties']['sheetId']


def col_letter(n):
    """0-indexed column number → letter (0→A, 25→Z, 26→AA …)"""
    s = ''
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def a1(row, col, sheet=None):
    ref = f'{col_letter(col)}{row + 1}'
    return f"'{sheet}'!{ref}" if sheet else ref


def section_rows():
    """Return 0-indexed row indices where section headers appear."""
    headers = {
        '3 VALUE PILLARS', 'PILOT DATA', 'AUTOMATION RATE',
        'PHASE ROADMAP', 'AI COST MODEL', 'ROI SUMMARY', 'GOVERNANCE',
    }
    found = []
    for i, row in enumerate(ROWS):
        if row and any(h in str(row[0]) for h in headers):
            found.append(i)
    return found


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    creds   = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    sheet   = service.spreadsheets()

    sheet_id = ensure_sheet(service, SPREADSHEET_ID, SHEET_NAME)

    # 1. Write all values
    range_name = f"'{SHEET_NAME}'!A1"
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body={'values': ROWS},
    ).execute()

    num_rows = len(ROWS)
    num_cols = 6

    requests = []

    # 2. Clear all existing formatting first
    requests.append({'repeatCell': {
        'range': {
            'sheetId': sheet_id,
            'startRowIndex': 0, 'endRowIndex': num_rows,
            'startColumnIndex': 0, 'endColumnIndex': num_cols,
        },
        'cell': {'userEnteredFormat': {
            'backgroundColor': WHITE,
            'textFormat': {'bold': False, 'fontSize': 10,
                           'foregroundColor': WARM_GRAY},
        }},
        'fields': 'userEnteredFormat(backgroundColor,textFormat)',
    }})

    # 3. Row 0: Big title — dark green bg, lime text
    requests.append({'repeatCell': {
        'range': {'sheetId': sheet_id,
                  'startRowIndex': 0, 'endRowIndex': 1,
                  'startColumnIndex': 0, 'endColumnIndex': num_cols},
        'cell': {'userEnteredFormat': {
            'backgroundColor': DARK_GREEN,
            'textFormat': {'bold': True, 'fontSize': 14,
                           'foregroundColor': LIME},
        }},
        'fields': 'userEnteredFormat(backgroundColor,textFormat)',
    }})
    # Merge title across all columns
    requests.append({'mergeCells': {
        'range': {'sheetId': sheet_id,
                  'startRowIndex': 0, 'endRowIndex': 1,
                  'startColumnIndex': 0, 'endColumnIndex': num_cols},
        'mergeType': 'MERGE_ALL',
    }})

    # 4. Row 1: subtitle — lime bg, dark green text
    requests.append({'repeatCell': {
        'range': {'sheetId': sheet_id,
                  'startRowIndex': 1, 'endRowIndex': 2,
                  'startColumnIndex': 0, 'endColumnIndex': num_cols},
        'cell': {'userEnteredFormat': {
            'backgroundColor': LIME,
            'textFormat': {'bold': True, 'fontSize': 10,
                           'foregroundColor': DARK_GREEN},
        }},
        'fields': 'userEnteredFormat(backgroundColor,textFormat)',
    }})
    requests.append({'mergeCells': {
        'range': {'sheetId': sheet_id,
                  'startRowIndex': 1, 'endRowIndex': 2,
                  'startColumnIndex': 0, 'endColumnIndex': num_cols},
        'mergeType': 'MERGE_ALL',
    }})

    # 5. Section header rows — mid-green bg, white bold text
    SECTION_ACCENT = {
        '3 VALUE PILLARS':      MID_GREEN,
        'PILOT DATA':           MID_GREEN,
        'AUTOMATION RATE':      MID_GREEN,
        'PHASE ROADMAP':        MID_GREEN,
        'AI COST MODEL':        TEAL,
        'ROI SUMMARY':          PURPLE,
        'GOVERNANCE':           ORANGE,
    }
    for i, row in enumerate(ROWS):
        if not row or not row[0]:
            continue
        matched = next((v for k, v in SECTION_ACCENT.items() if k in str(row[0])), None)
        if matched:
            requests.append({'repeatCell': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': i, 'endRowIndex': i + 1,
                          'startColumnIndex': 0, 'endColumnIndex': num_cols},
                'cell': {'userEnteredFormat': {
                    'backgroundColor': matched,
                    'textFormat': {'bold': True, 'fontSize': 11,
                                   'foregroundColor': WHITE},
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat)',
            }})
            requests.append({'mergeCells': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': i, 'endRowIndex': i + 1,
                          'startColumnIndex': 0, 'endColumnIndex': num_cols},
                'mergeType': 'MERGE_ALL',
            }})

    # 6. Sub-header rows (row after section header, where row[0] is '#', 'Team', 'Phase'…)
    SUB_HEADERS = {'#', 'Team', 'Phase', 'Issue type', 'Cost item',
                   'Scenario', 'Risk', 'Pillar', 'Cost item'}
    for i, row in enumerate(ROWS):
        if row and str(row[0]) in SUB_HEADERS:
            requests.append({'repeatCell': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': i, 'endRowIndex': i + 1,
                          'startColumnIndex': 0, 'endColumnIndex': num_cols},
                'cell': {'userEnteredFormat': {
                    'backgroundColor': OFF_WHITE,
                    'textFormat': {'bold': True, 'fontSize': 9,
                                   'foregroundColor': DARK_GREEN},
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat)',
            }})

    # 7. TOTAL rows — light bold
    for i, row in enumerate(ROWS):
        if row and str(row[0]).upper().startswith('TOTAL'):
            requests.append({'repeatCell': {
                'range': {'sheetId': sheet_id,
                          'startRowIndex': i, 'endRowIndex': i + 1,
                          'startColumnIndex': 0, 'endColumnIndex': num_cols},
                'cell': {'userEnteredFormat': {
                    'backgroundColor': LIGHT_GRAY,
                    'textFormat': {'bold': True, 'fontSize': 10,
                                   'foregroundColor': DARK_GREEN},
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat)',
            }})

    # 8. Freeze top 2 rows + first column
    requests.append({'updateSheetProperties': {
        'properties': {
            'sheetId': sheet_id,
            'gridProperties': {'frozenRowCount': 2, 'frozenColumnCount': 1},
        },
        'fields': 'gridProperties(frozenRowCount,frozenColumnCount)',
    }})

    # 9. Column widths
    COL_WIDTHS = [220, 130, 260, 240, 240, 160]
    for col_i, width in enumerate(COL_WIDTHS):
        requests.append({'updateDimensionProperties': {
            'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                      'startIndex': col_i, 'endIndex': col_i + 1},
            'properties': {'pixelSize': width},
            'fields': 'pixelSize',
        }})

    # 10. Wrap text across all data cells
    requests.append({'repeatCell': {
        'range': {'sheetId': sheet_id,
                  'startRowIndex': 0, 'endRowIndex': num_rows,
                  'startColumnIndex': 0, 'endColumnIndex': num_cols},
        'cell': {'userEnteredFormat': {'wrapStrategy': 'WRAP'}},
        'fields': 'userEnteredFormat(wrapStrategy)',
    }})

    # 11. Border around the whole data range
    border = {'style': 'SOLID', 'color': LIGHT_GRAY}
    requests.append({'updateBorders': {
        'range': {'sheetId': sheet_id,
                  'startRowIndex': 0, 'endRowIndex': num_rows,
                  'startColumnIndex': 0, 'endColumnIndex': num_cols},
        'top': border, 'bottom': border, 'left': border, 'right': border,
        'innerHorizontal': {'style': 'SOLID', 'color': LIGHT_GRAY},
        'innerVertical':   {'style': 'SOLID', 'color': LIGHT_GRAY},
    }})

    # Apply all formatting
    sheet.batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': requests},
    ).execute()

    print(f'Done — tab "{SHEET_NAME}" written to:')
    print(f'https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit')


if __name__ == '__main__':
    main()
