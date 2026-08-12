#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import re
import glob
import math
import warnings
import logging
from datetime import datetime, timedelta
import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
from psycopg2 import sql
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore")
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration Cloud
NEON_CONN_STRING = os.environ.get("NEON_CONN_STRING")
DOSSIER_DESTINATION = "erec_data"
DOSSIER_NICE = "nice_data"

def parse_excel_date(valeur):
    if valeur is None or pd.isna(valeur) or valeur == '': return None
    if isinstance(valeur, float) and math.isnan(valeur): return None
    if isinstance(valeur, (int, float)):
        try: return (datetime(1899, 12, 30) + timedelta(days=valeur)).strftime('%Y-%m-%d')
        except: return None
    if isinstance(valeur, str):
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%Y/%m/%d"):
            try: return datetime.strptime(valeur.strip(), fmt).strftime('%Y-%m-%d')
            except ValueError: continue
    return None

def parse_time(valeur):
    if valeur is None or pd.isna(valeur) or valeur == '': return None
    if isinstance(valeur, float) and math.isnan(valeur): return None
    if isinstance(valeur, (int, float)):
        try:
            secs = int(valeur)
            if secs >= 86400: secs = secs % 86400
            return f"{secs // 3600:02d}:{(secs % 3600) // 60:02d}:{secs % 60:02d}"
        except: return None
    if isinstance(valeur, str):
        cleaned = valeur.strip().lower()
        if cleaned in ('nan', 'none', 'null', ''): return None
        match = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$', cleaned)
        if match:
            h, m, s = int(match.group(1)), int(match.group(2)), int(match.group(3)) if match.group(3) else 0
            return f"{h:02d}:{m:02d}:{s:02d}"
    return None

def preparer_data_erec():
    fichiers = glob.glob(os.path.join(DOSSIER_DESTINATION, "*.xls"))
    if not fichiers: return None
    dfs = []
    for fichier in fichiers:
        try:
            with open(fichier, 'r', encoding='utf-8', errors='ignore') as f: soup = BeautifulSoup(f, 'xml')
            tableau_data = []
            for ligne in soup.find_all('Row'):
                ligne_data = [cell.find('Data').text.strip() if cell.find('Data') else "" for cell in ligne.find_all('Cell')]
                if any(ligne_data): tableau_data.append(ligne_data)
            if tableau_data: dfs.append(pd.DataFrame(tableau_data))
        except Exception as e: logger.error(f"[EREC] Erreur sur {os.path.basename(fichier)} : {e}")
    if not dfs: return None
    df_final = pd.concat(dfs, ignore_index=True).replace("", pd.NA).dropna(subset=[0])
    df_final.columns = df_final.iloc[0].astype(str).str.strip()
    df_final = df_final[1:]
    if 'Date' in df_final.columns:
        df_final = df_final[df_final['Date'] != "Date"]
        df_final['Date'] = pd.to_datetime(df_final['Date'], dayfirst=True, errors='coerce').dt.date.astype(str).replace('NaT', None)
    if 'IRIS ID' in df_final.columns:
        df_final['IRIS ID'] = pd.to_numeric(df_final['IRIS ID'], errors='coerce').astype('Int64')
        df_final.rename(columns={'IRIS ID': 'Workday ID'}, inplace=True)
    mapping = {'Date': 'Date', 'Workday ID': 'Workday ID', 'Activity': 'Activity', 'Start time': 'Start time', 'Stop time': 'Stop time', 'Project': 'Project'}
    df_final = df_final[[c for c in mapping.keys() if c in df_final.columns]].rename(columns=mapping)
    for col in ['Start time', 'Stop time']:
        if col in df_final.columns:
            df_final[col] = df_final[col].apply(parse_time)
            df_final[col] = df_final[col].astype(object)
            df_final[col] = df_final[col].where(pd.notna(df_final[col]), None)
    df_final['is_fim'] = False
    df_final = df_final.dropna(subset=['Date'])
    logger.info(f"[EREC] Nombre de lignes préparées : {len(df_final)}")
    return df_final

def preparer_data_nice():
    fichiers = [f for f in glob.glob(os.path.join(DOSSIER_NICE, "*.*")) if not os.path.basename(f).startswith("~$")]
    if not fichiers: return None
    dfs = []
    for fichier in fichiers:
        try: dfs.append(pd.read_excel(fichier, header=None))
        except Exception as e: logger.error(f"[NICE] Erreur lecture {os.path.basename(fichier)}: {e}")
    if not dfs: return None
    df_final = pd.concat(dfs, ignore_index=True).rename(columns={1: "Column2", 2: "Date", 3: "Debut_shift", 4: "Fin_shift", 6: "Activite", 7: "Debut_activite", 10: "Fin_Activite"})
    for col in ["Date", "Debut_shift", "Fin_shift", "Debut_activite", "Fin_Activite"]:
        if col in df_final.columns: df_final[col] = df_final[col].replace(["Date", "Début", "Fin"], pd.NA)
    df_final['Remarque'] = df_final['Column2'].apply(lambda x: x if isinstance(x, str) and "Remarque:" in x else pd.NA)
    df_final[['Date', 'Column2', 'Remarque']] = df_final[['Date', 'Column2', 'Remarque']].ffill()
    df_final['Projet'] = df_final['Remarque'].str[14:]
    df_final.dropna(subset=['Date'], inplace=True)
    df_final = df_final[~df_final['Column2'].astype(str).str.contains("Administration, IEX Platform", na=False)]
    df_final[['Nice_ID', 'Agent']] = df_final['Column2'].astype(str).str.extract(r'(?P<Nice_ID>\d+)(?P<Agent>.*)')
    df_final.dropna(subset=['Debut_shift', 'Fin_shift', 'Activite', 'Debut_activite', 'Fin_Activite'], how='all', inplace=True)
    df_final['Nice_ID'] = pd.to_numeric(df_final['Nice_ID'], errors='coerce').astype('Int64')
    df_final['Date'] = pd.to_datetime(df_final['Date'], dayfirst=True, errors='coerce').dt.date.astype(str)
    for col in ["Debut_shift", "Fin_shift", "Debut_activite", "Fin_Activite"]:
        if col in df_final.columns:
            df_final[col] = df_final[col].apply(parse_time)
            df_final[col] = df_final[col].astype(object)
            df_final[col] = df_final[col].where(pd.notna(df_final[col]), None)
    df_final = df_final[[c for c in ['Projet', 'Nice_ID', 'Agent', 'Date', 'Debut_shift', 'Fin_shift', 'Activite', 'Debut_activite', 'Fin_Activite'] if c in df_final.columns]].dropna(subset=['Date'])
    logger.info(f"[NICE] Nombre de lignes préparées : {len(df_final)}")
    return df_final

def inject_dataframe(conn, df, table_name, conflict_columns=None, replace_keys=None):
    if df is None or df.empty: return 0
    with conn.cursor() as cur:
        cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = %s AND column_name NOT IN ('rowid', 'created_at', 'updated_at')", (table_name,))
        existing_cols = [row[0] for row in cur.fetchall()]
    cols = [c for c in df.columns if c in existing_cols]
    if not cols: return 0
    df_clean = df[cols].copy()
    
    if conflict_columns:
        conflict_cols_exist = [col for col in conflict_columns if col in df_clean.columns]
        if conflict_cols_exist:
            df_clean = df_clean.drop_duplicates(subset=conflict_cols_exist, keep='last')
            # Trier les données pour éviter les deadlocks PostgreSQL
            df_clean = df_clean.sort_values(by=conflict_cols_exist)
            
    if replace_keys:
        replace_keys_exist = [k for k in replace_keys if k in df_clean.columns]
        if replace_keys_exist:
            pairs_df = df_clean[replace_keys_exist].astype(object).dropna(how='all').drop_duplicates()
            pairs = [tuple(None if pd.isna(v) else str(v) for v in row) for row in pairs_df.itertuples(index=False, name=None)]
            if pairs:
                value_col_names = [f"k{i}" for i in range(len(replace_keys_exist))]
                delete_stmt = sql.SQL("DELETE FROM {table} WHERE ({target_cols}) IN (SELECT {value_cols} FROM (VALUES %s) AS v ({value_cols_def}))").format(
                    table=sql.Identifier(table_name), target_cols=sql.SQL(', ').join(sql.SQL("{}::text").format(sql.Identifier(k)) for k in replace_keys_exist),
                    value_cols=sql.SQL(', ').join(sql.Identifier(c) for c in value_col_names), value_cols_def=sql.SQL(', ').join(sql.Identifier(c) for c in value_col_names))
                with conn.cursor() as cur:
                    execute_values(cur, delete_stmt, pairs, page_size=1000)
                    logger.info(f"Table {table_name}: suppression effectuée pour {len(pairs)} couple(s) avant réinjection.")
                    
    df_clean = df_clean.astype(object)
    data = [tuple(None if pd.isna(v) else v for v in row) for row in df_clean.to_records(index=False)]
    quoted_cols = [sql.Identifier(col) for col in cols]
    insert_stmt = sql.SQL("INSERT INTO {} ({}) VALUES %s").format(sql.Identifier(table_name), sql.SQL(', ').join(quoted_cols))
    if conflict_columns:
        conflict_cols = [sql.Identifier(col) for col in conflict_columns]
        update_cols = [sql.SQL("{} = EXCLUDED.{}").format(sql.Identifier(col), sql.Identifier(col)) for col in cols if col not in conflict_columns]
        if update_cols:
            today_str = datetime.now().strftime('%Y-%m-%d')
            where_conditions = [sql.SQL("{} IS DISTINCT FROM EXCLUDED.{}").format(sql.Identifier(table_name, col), sql.Identifier(col)) for col in cols if col not in conflict_columns]
            date_condition = sql.SQL("{}::text = {}").format(sql.Identifier(table_name, "Date"), sql.Literal(today_str))
            where_clause = sql.SQL("({} OR {})").format(date_condition, sql.SQL(' OR ').join(where_conditions))
            conflict_stmt = sql.SQL("ON CONFLICT ({}) DO UPDATE SET {} WHERE {}").format(sql.SQL(', ').join(conflict_cols), sql.SQL(', ').join(update_cols), where_clause)
        else: conflict_stmt = sql.SQL("ON CONFLICT ({}) DO NOTHING").format(sql.SQL(', ').join(conflict_cols))
    else: conflict_stmt = sql.SQL("")
    final_sql = sql.SQL("{} {}").format(insert_stmt, conflict_stmt)
    with conn.cursor() as cur: execute_values(cur, final_sql, data, page_size=1000)
    conn.commit()
    logger.info(f"Table {table_name}: Scan de {len(data)} lignes via Smart Upsert terminé.")
    return len(data)

if __name__ == "__main__":
    logger.info("=== Démarrage de l'injection Cloud (GitHub Actions) ===")
    try:
        conn = psycopg2.connect(NEON_CONN_STRING)
        conn.autocommit = False
        
        df_erec = preparer_data_erec()
        df_nice = preparer_data_nice()
        
        if df_erec is not None:
            inject_dataframe(conn, df_erec, 'erec', conflict_columns=['Date', 'Workday ID', 'Start time', 'Activity'])
        if df_nice is not None:
            inject_dataframe(conn, df_nice, 'nice', conflict_columns=['Date', 'Nice_ID', 'Debut_activite'], replace_keys=['Date', 'Nice_ID'])
            
        with conn.cursor() as cur:
            cur.execute("REFRESH MATERIALIZED VIEW CONCURRENTLY mv_daily_agent_stats")
        conn.commit()
        logger.info("✅ Vue matérialisée PostgreSQL rafraîchie.")
        conn.close()
        logger.info("=== Injection Cloud terminée avec succès ===")
    except Exception as e:
        logger.error(f"❌ Erreur fatale: {e}", exc_info=True)
        exit(1)
