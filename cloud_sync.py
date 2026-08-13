#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import re
import glob
import math
import warnings
import logging
from datetime import datetime, timedelta
import requests
import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
from psycopg2 import sql
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore")
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Config Cloud
NEON_CONN_STRING = os.environ.get("NEON_CONN_STRING")

# Config Entreprise (Lue depuis les Secrets GitHub)
MON_ID = os.environ.get("eRec_USERNAME")
MON_MDP = os.environ.get("eRec_PASSWORD")
REPORT_URL = "http://10.253.68.142:801/Reports.aspx"
FAUX_DNA = "0,3.1,0,0,0,1914395348,0,-1,-1,0,-1,-1,0,-1,-1,0,-1,-1,1,0,0,1,1,1,97852684,1,1,0,0,0,1,1920,1080,2,0,145,0,128858925"
FAUX_DEVICE_ID = "128858925"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

LISTE_PROJETS = {
    "1165": "Auchan_Retail",
    "1097": "Basic_Fit_Mada",
    "1474": "Basic_Fit_Polyglote",
    "1073": "BYTEL_DIGITAL",
    "1075": "DEPOT_BINGO",
    "1561": "Driiveme",
    "1294": "GRANEET_Back_Office",
    "871": "HABITAT_PRESTO",
    "1260": "Juritravail",
    "1230": "LegalPlace",
    "2903": "LINXEA_KYC",
    "2933": "Shiseido_CPB_SL_et_Gallinee",
    "1568": "Smartbox",
    "1064": "Stych",
    "1431": "TRUSTPAIR",
    "1292": "Back_Market",
    "1464": "BackMarket_Polyglot",
    "1183": "Zalando_SC"
}

# Dossiers temporaires
DOSSIER_DESTINATION = "/tmp/erec_data"
DOSSIER_NICE = "/tmp/nice_data"

def get_asp_tokens(html):
    def extract(pattern):
        m = re.search(pattern, html)
        return m.group(1) if m else ""
    return {
        "VS": extract(r'id="__VIEWSTATE"\s+value="([^"]*)"'),
        "VSG": extract(r'id="__VIEWSTATEGENERATOR"\s+value="([^"]*)"'),
        "EV": extract(r'id="__EVENTVALIDATION"\s+value="([^"]*)"'),
        "IDUser": extract(r'id="main_content_hf_ID_User"\s+value="([^"]*)"'),
        "Login": extract(r'id="main_content_hf_Login"\s+value="([^"]*)"'),
        "Name": extract(r'id="main_content_hf_Name"\s+value="([^"]*)"')
    }

def build_report_body(tokens, event_target, btn_name, btn_val, date_str, team, project, mon_id):
    body = {
        'ToolkitScriptManager1_HiddenField': ';;AjaxControlToolkit...', '__EVENTTARGET': event_target, '__EVENTARGUMENT': '', '__LASTFOCUS': '',
        '__VIEWSTATE': tokens['VS'], '__VIEWSTATEGENERATOR': tokens['VSG'], '__EVENTVALIDATION': tokens['EV'],
        'ctl00$main_content$ddl_team': team, 'ctl00$main_content$ddl_agents': '0', 'ctl00$main_content$ddl_project': project,
        'ctl00$main_content$hf_language': 'en-EN', 'ctl00$main_content$hf_ID_User': tokens['IDUser'] if tokens['IDUser'] else '27601',
        'ctl00$main_content$hf_Login': tokens['Login'] if tokens['Login'] else mon_id, 'ctl00$main_content$hf_Name': tokens['Name']
    }
    if date_str: body['ctl00$main_content$txt_date'] = date_str
    if btn_name: body[btn_name] = btn_val
    return body

def telecharger_donnees():
    os.makedirs(DOSSIER_DESTINATION, exist_ok=True)
    os.makedirs(DOSSIER_NICE, exist_ok=True)
    session = requests.Session()
    session.headers.update({'User-Agent': USER_AGENT})
    
    res = session.get(REPORT_URL)
    tokens = get_asp_tokens(res.text)

    if "Login.aspx" in res.url:
        logger.info("Authentification Erec requise...")
        login_body = {
            '__EVENTTARGET': '', '__EVENTARGUMENT': '', '__LASTFOCUS': '', '__VIEWSTATE': tokens['VS'],
            '__VIEWSTATEGENERATOR': tokens['VSG'], '__EVENTVALIDATION': tokens['EV'], 'bt_language': 'en-EN',
            'txt_userName': MON_ID, 'txt_password': MON_MDP, 'bt_login': 'Login', 'hf_DNA': FAUX_DNA, 'hf_deviceID': FAUX_DEVICE_ID
        }
        session.post(res.url, data=login_body)
        res = session.get(REPORT_URL)
        tokens = get_asp_tokens(res.text)

    date_jour = datetime.now().strftime("%d/%m/%Y")
    date_fich = datetime.now().strftime("%d-%m-%Y")

    tokens = get_asp_tokens(session.post(REPORT_URL, data=build_report_body(tokens, '', 'ctl00$main_content$bt_custom', 'Custom', None, '0', '0', MON_ID)).text)
    tokens = get_asp_tokens(session.post(REPORT_URL, data=build_report_body(tokens, 'ctl00$main_content$txt_date', '', '', date_jour, '0', '0', MON_ID)).text)
    tokens = get_asp_tokens(session.post(REPORT_URL, data=build_report_body(tokens, 'ctl00$main_content$ddl_team', '', '', date_jour, '0', '0', MON_ID)).text)

    for pid, nom in LISTE_PROJETS.items():
        logger.info(f"Téléchargement {nom}...")
        tokens = get_asp_tokens(session.post(REPORT_URL, data=build_report_body(tokens, 'ctl00$main_content$ddl_project', '', '', date_jour, '0', pid, MON_ID)).text)
        res_exp = session.post(REPORT_URL, data=build_report_body(tokens, '', 'ctl00$main_content$bt_activityDetailsReport', 'Activity Details Report', date_jour, '0', pid, MON_ID))
        if res_exp.status_code == 200:
            file_path = os.path.join(DOSSIER_DESTINATION, f"{nom}_{date_fich}.xls")
            with open(file_path, 'wb') as f: f.write(res_exp.content)
        else: 
            logger.warning(f"Échec téléchargement pour {nom}")

def parse_excel_date(v):
    if v is None or pd.isna(v) or v == '': return None
    if isinstance(v, float) and math.isnan(v): return None
    if isinstance(v, (int, float)):
        try: return (datetime(1899, 12, 30) + timedelta(days=v)).strftime('%Y-%m-%d')
        except: return None
    if isinstance(v, str):
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%Y/%m/%d"):
            try: return datetime.strptime(v.strip(), fmt).strftime('%Y-%m-%d')
            except: pass
    return None

def parse_time(v):
    if v is None or pd.isna(v) or v == '': return None
    if isinstance(v, float) and math.isnan(v): return None
    if isinstance(v, (int, float)):
        try:
            s = int(v)
            if s >= 86400: s = s % 86400
            return f"{s // 3600:02d}:{(s % 3600) // 60:02d}:{s % 60:02d}"
        except: return None
    if isinstance(v, str):
        c = v.strip().lower()
        if c in ('nan', 'none', 'null', ''): return None
        m = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$', c)
        if m: return f"{int(m.group(1)):02d}:{int(m.group(2)):02d}:{int(m.group(3)) if m.group(3) else 0:02d}"
    return None

def preparer_data_erec():
    fichiers = glob.glob(os.path.join(DOSSIER_DESTINATION, "*.xls"))
    if not fichiers: return None
    dfs = []
    for f in fichiers:
        try:
            with open(f, 'r', encoding='utf-8', errors='ignore') as file: soup = BeautifulSoup(file, 'xml')
            t = [[c.find('Data').text.strip() if c.find('Data') else "" for c in r.find_all('Cell')] for r in soup.find_all('Row')]
            if t: dfs.append(pd.DataFrame(t))
        except Exception as e: logger.error(f"Err Erec {f}: {e}")
    if not dfs: return None
    df = pd.concat(dfs, ignore_index=True).replace("", pd.NA).dropna(subset=[0])
    df.columns = df.iloc[0].astype(str).str.strip()
    df = df[1:]
    if 'Date' in df.columns:
        df = df[df['Date'] != "Date"]
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce').dt.date.astype(str).replace('NaT', None)
    if 'IRIS ID' in df.columns:
        df['IRIS ID'] = pd.to_numeric(df['IRIS ID'], errors='coerce').astype('Int64')
        df.rename(columns={'IRIS ID': 'Workday ID'}, inplace=True)
    m = {'Date':'Date', 'Workday ID':'Workday ID', 'Activity':'Activity', 'Start time':'Start time', 'Stop time':'Stop time', 'Project':'Project'}
    df = df[[c for c in m if c in df.columns]].rename(columns=m)
    for c in ['Start time', 'Stop time']:
        if c in df.columns: df[c] = df[c].apply(parse_time).astype(object).where(pd.notna(df[c]), None)
    df['is_fim'] = False
    df = df.dropna(subset=['Date'])
    logger.info(f"[EREC] {len(df)} lignes préparées.")
    return df

def preparer_data_nice():
    fichiers = [f for f in glob.glob(os.path.join(DOSSIER_NICE, "*.*")) if not os.path.basename(f).startswith("~$")]
    if not fichiers: return None
    dfs = []
    for f in fichiers:
        try: dfs.append(pd.read_excel(f, header=None))
        except: pass
    if not dfs: return None
    df = pd.concat(dfs, ignore_index=True).rename(columns={1:"Column2", 2:"Date", 3:"Debut_shift", 4:"Fin_shift", 6:"Activite", 7:"Debut_activite", 10:"Fin_Activite"})
    for c in ["Date", "Debut_shift", "Fin_shift", "Debut_activite", "Fin_Activite"]:
        if c in df.columns: df[c] = df[c].replace(["Date", "Début", "Fin"], pd.NA)
    df['Remarque'] = df['Column2'].apply(lambda x: x if isinstance(x, str) and "Remarque:" in x else pd.NA)
    df[['Date', 'Column2', 'Remarque']] = df[['Date', 'Column2', 'Remarque']].ffill()
    df['Projet'] = df['Remarque'].str[14:]
    df.dropna(subset=['Date'], inplace=True)
    df = df[~df['Column2'].astype(str).str.contains("Administration", na=False)]
    df[['Nice_ID', 'Agent']] = df['Column2'].astype(str).str.extract(r'(?P<Nice_ID>\d+)(?P<Agent>.*)')
    df.dropna(subset=['Debut_shift', 'Fin_shift', 'Activite', 'Debut_activite', 'Fin_Activite'], how='all', inplace=True)
    df['Nice_ID'] = pd.to_numeric(df['Nice_ID'], errors='coerce').astype('Int64')
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce').dt.date.astype(str)
    for c in ["Debut_shift", "Fin_shift", "Debut_activite", "Fin_Activite"]:
        if c in df.columns: df[c] = df[c].apply(parse_time).astype(object).where(pd.notna(df[c]), None)
    df = df[[c for c in ['Projet','Nice_ID','Agent','Date','Debut_shift','Fin_shift','Activite','Debut_activite','Fin_Activite'] if c in df.columns]].dropna(subset=['Date'])
    logger.info(f"[NICE] {len(df)} lignes préparées.")
    return df

def inject_dataframe(conn, df, table_name, conflict_columns=None, replace_keys=None):
    if df is None or df.empty: return 0
    with conn.cursor() as cur:
        cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = %s AND column_name NOT IN ('rowid','created_at','updated_at')", (table_name,))
        cols = [r[0] for r in cur.fetchall()]
    cols = [c for c in df.columns if c in cols]
    if not cols: return 0
    df_c = df[cols].copy()
    if conflict_columns:
        cc = [c for c in conflict_columns if c in df_c.columns]
        if cc: df_c = df_c.drop_duplicates(subset=cc, keep='last').sort_values(by=cc)
    if replace_keys:
        rk = [k for k in replace_keys if k in df_c.columns]
        if rk:
            p = df_c[rk].astype(object).dropna(how='all').drop_duplicates()
            pairs = [tuple(None if pd.isna(v) else str(v) for v in r) for r in p.itertuples(index=False, name=None)]
            if pairs:
                vcn = [f"k{i}" for i in range(len(rk))]
                ds = sql.SQL("DELETE FROM {} WHERE ({}) IN (SELECT {} FROM (VALUES %s) AS v ({}))").format(
                    sql.Identifier(table_name), sql.SQL(', ').join(sql.SQL("{}::text").format(sql.Identifier(k)) for k in rk),
                    sql.SQL(', ').join(sql.Identifier(c) for c in vcn), sql.SQL(', ').join(sql.Identifier(c) for c in vcn))
                with conn.cursor() as cur: execute_values(cur, ds, pairs, page_size=1000)
    df_c = df_c.astype(object)
    data = [tuple(None if pd.isna(v) else v for v in r) for r in df_c.to_records(index=False)]
    qc = [sql.Identifier(c) for c in cols]
    is_ = sql.SQL("INSERT INTO {} ({}) VALUES %s").format(sql.Identifier(table_name), sql.SQL(', ').join(qc))
    if conflict_columns:
        cf = [sql.Identifier(c) for c in conflict_columns]
        uc = [sql.SQL("{} = EXCLUDED.{}").format(sql.Identifier(c), sql.Identifier(c)) for c in cols if c not in conflict_columns]
        if uc:
            ts = datetime.now().strftime('%Y-%m-%d')
            wc = [sql.SQL("{} IS DISTINCT FROM EXCLUDED.{}").format(sql.Identifier(table_name, c), sql.Identifier(c)) for c in cols if c not in conflict_columns]
            dc = sql.SQL("{}::text = {}").format(sql.Identifier(table_name, "Date"), sql.Literal(ts))
            wh = sql.SQL("({} OR {})").format(dc, sql.SQL(' OR ').join(wc))
            cs = sql.SQL("ON CONFLICT ({}) DO UPDATE SET {} WHERE {}").format(sql.SQL(', ').join(cf), sql.SQL(', ').join(uc), wh)
        else: cs = sql.SQL("ON CONFLICT ({}) DO NOTHING").format(sql.SQL(', ').join(cf))
    else: cs = sql.SQL("")
    with conn.cursor() as cur: execute_values(cur, sql.SQL("{} {}").format(is_, cs), data, page_size=1000)
    conn.commit()
    return len(data)

if __name__ == "__main__":
    logger.info("=== Démarrage du script Cloud (VPN) ===")
    telecharger_donnees()
    
    df_erec = preparer_data_erec()
    df_nice = preparer_data_nice()
    
    conn = psycopg2.connect(NEON_CONN_STRING)
    conn.autocommit = False
    if df_erec is not None: inject_dataframe(conn, df_erec, 'erec', conflict_columns=['Date', 'Workday ID', 'Start time', 'Activity'])
    if df_nice is not None: inject_dataframe(conn, df_nice, 'nice', conflict_columns=['Date', 'Nice_ID', 'Debut_activite'], replace_keys=['Date', 'Nice_ID'])
    with conn.cursor() as cur: cur.execute("REFRESH MATERIALIZED VIEW CONCURRENTLY mv_daily_agent_stats")
    conn.commit()
    conn.close()
    logger.info("✅ Injection Cloud terminée avec succès !")
