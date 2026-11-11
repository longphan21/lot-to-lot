import psycopg2
import pandas as pd
from datetime import datetime, timedelta
import streamlit as st

# Lấy chuỗi kết nối từ Secrets của Streamlit
DB_URL = st.secrets["SUPABASE_URL"]


def get_connection():
    return psycopg2.connect(DB_URL)


def init_db():
    with get_connection() as conn:
        cursor = conn.cursor()
        # Dùng SERIAL thay cho AUTOINCREMENT của PostgreSQL
        cursor.execute("""
                       CREATE TABLE IF NOT EXISTS stage1_setup
                       (
                           id
                           SERIAL
                           PRIMARY
                           KEY,
                           test_name
                           TEXT
                           NOT
                           NULL,
                           machine
                           TEXT
                           NOT
                           NULL,
                           machine_id
                           TEXT,
                           is_selected
                           INTEGER
                           DEFAULT
                           1,
                           target_conc
                           REAL
                           NOT
                           NULL,
                           unit
                           TEXT
                           DEFAULT
                           'mmol/L',
                           tea_pct
                           REAL
                           NOT
                           NULL,
                           swrl_pct
                           REAL
                           NOT
                           NULL,
                           sr_pct
                           REAL
                           NOT
                           NULL,
                           cd_pct
                           REAL,
                           cd_swrl
                           REAL,
                           sr_swrl
                           REAL,
                           power_achieve
                           REAL,
                           n_samples
                           INTEGER,
                           rejection_limit
                           REAL,
                           k_factor
                           REAL,
                           note
                           TEXT,
                           updated_at
                           TIMESTAMP
                           DEFAULT
                           CURRENT_TIMESTAMP
                       )
                       """)
        cursor.execute("""
                       CREATE TABLE IF NOT EXISTS stage2_evaluations
                       (
                           id
                           SERIAL
                           PRIMARY
                           KEY,
                           eval_date
                           TEXT
                           NOT
                           NULL,
                           test_name
                           TEXT
                           NOT
                           NULL,
                           machine
                           TEXT
                           NOT
                           NULL,
                           machine_id
                           TEXT,
                           unit
                           TEXT,
                           lot_old
                           TEXT
                           NOT
                           NULL,
                           lot_new
                           TEXT
                           NOT
                           NULL,
                           level_name
                           TEXT,
                           target_conc
                           REAL,
                           sid
                           TEXT,
                           val_old
                           REAL,
                           val_new
                           REAL,
                           diff_d
                           REAL,
                           mean_d
                           REAL,
                           mean_d_pct
                           REAL,
                           rl_pct
                           REAL,
                           k_factor
                           REAL,
                           status
                           TEXT,
                           reagent_name
                           TEXT
                           DEFAULT
                           '',
                           created_at
                           TIMESTAMP
                           DEFAULT
                           CURRENT_TIMESTAMP
                       )
                       """)
        conn.commit()


def save_test_setup(test_name: str, machine: str, machine_id: str, df_rows: pd.DataFrame):
    init_db()
    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()
    mach_id_clean = machine_id.strip()
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')

    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM stage1_setup WHERE UPPER(test_name) = %s AND machine = %s",
            (test_clean, machine_clean)
        )
        for _, row in df_rows.iterrows():
            cursor.execute("""
                           INSERT INTO stage1_setup (test_name, machine, machine_id, is_selected, target_conc, unit,
                                                     tea_pct, swrl_pct, sr_pct, cd_pct, cd_swrl, sr_swrl,
                                                     power_achieve, n_samples, rejection_limit, k_factor, note,
                                                     updated_at)
                           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                           """, (
                               test_clean, machine_clean, mach_id_clean,
                               1 if bool(row.get("Chọn chạy (Vàng)", False)) else 0,
                               float(row.get("Target concentration", 0)),
                               str(row.get("unit", "mmol/L")),
                               float(row.get("TEa%", 0)),
                               float(row.get("Swrl = CVwrl %", 0)),
                               float(row.get("Sr = CVr %", 0)),
                               0.0,
                               float(row.get("cd_swrl", 0)),
                               float(row.get("sr_swrl", 0)),
                               float(row.get("power_achieve", 0)),
                               int(row.get("n_samples", 0)),
                               float(row.get("rejection_limit", 0)),
                               float(row.get("k_factor", 0)),
                               str(row.get("Ghi chú", "")),
                               vn_now
                           ))
        conn.commit()


def load_all_setups() -> pd.DataFrame:
    init_db()
    with get_connection() as conn:
        return pd.read_sql_query("SELECT * FROM stage1_setup ORDER BY test_name, machine, target_conc", conn)


def load_specific_setup(test_name: str, machine: str) -> pd.DataFrame:
    init_db()
    with get_connection() as conn:
        return pd.read_sql_query(
            "SELECT * FROM stage1_setup WHERE UPPER(test_name) = %s AND machine = %s ORDER BY target_conc",
            conn, params=(test_name.strip().upper(), machine.strip())
        )


def delete_specific_setup(test_name: str, machine: str):
    init_db()
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM stage1_setup WHERE UPPER(test_name) = %s AND machine = %s",
                       (test_name.strip().upper(), machine.strip()))
        conn.commit()


def save_stage2_evaluation(records: list):
    init_db()
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')

    with get_connection() as conn:
        cursor = conn.cursor()
        for r in records:
            record_time = r.get('created_at') if r.get('created_at') else vn_now
            cursor.execute("""
                           INSERT INTO stage2_evaluations (created_at, eval_date, test_name, machine, machine_id, unit,
                                                           lot_old, lot_new,
                                                           level_name, target_conc, sid, val_old, val_new, diff_d,
                                                           mean_d,
                                                           mean_d_pct, rl_pct, k_factor, status, reagent_name)
                           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                           """, (
                               record_time,
                               r.get('eval_date'),
                               r.get('test_name'),
                               r.get('machine'),
                               r.get('machine_id', ''),
                               r.get('unit', ''),
                               r.get('lot_old'),
                               r.get('lot_new'),
                               r.get('level_name'),
                               r.get('target_conc'),
                               r.get('sid'),
                               r.get('val_old'),
                               r.get('val_new'),
                               r.get('diff_d'),
                               r.get('mean_d'),
                               r.get('mean_d_pct'),
                               r.get('rl_pct'),
                               r.get('k_factor'),
                               r.get('status'),
                               r.get('reagent_name', '')
                           ))
        conn.commit()


def load_stage2_history() -> pd.DataFrame:
    init_db()
    with get_connection() as conn:
        return pd.read_sql_query("SELECT * FROM stage2_evaluations ORDER BY eval_date DESC, id DESC", conn)


def delete_stage2_evaluation(created_at: str):
    init_db()
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM stage2_evaluations WHERE created_at = %s", (created_at,))
        conn.commit()


def import_stage1_from_dataframe(df: pd.DataFrame) -> int:
    init_db()
    imported_count = 0
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')

    with get_connection() as conn:
        cursor = conn.cursor()
        unique_groups = df[['test_name', 'machine']].drop_duplicates()
        for _, group in unique_groups.iterrows():
            t_name = str(group['test_name']).strip().upper() if not pd.isna(group['test_name']) else ""
            m_name = str(group['machine']).strip() if not pd.isna(group['machine']) else ""
            if t_name and m_name:
                cursor.execute("DELETE FROM stage1_setup WHERE UPPER(test_name) = %s AND machine = %s",
                               (t_name, m_name))

        for _, row in df.iterrows():
            test_clean = str(row['test_name']).strip().upper() if not pd.isna(row.get('test_name')) else ""
            machine_clean = str(row['machine']).strip() if not pd.isna(row.get('machine')) else ""

            if not test_clean or not machine_clean:
                continue

            machine_id = str(row['machine_id']).strip() if 'machine_id' in df.columns and not pd.isna(
                row['machine_id']) else ""
            unit_str = str(row['unit']).strip() if 'unit' in df.columns and not pd.isna(row['unit']) else "mmol/L"
            note_str = str(row['note']).strip() if 'note' in df.columns and not pd.isna(row['note']) else ""

            cursor.execute("""
                           INSERT INTO stage1_setup (test_name, machine, machine_id, is_selected, target_conc, unit,
                                                     tea_pct, swrl_pct, sr_pct, cd_pct, cd_swrl, sr_swrl,
                                                     power_achieve, n_samples, rejection_limit, k_factor, note,
                                                     updated_at)
                           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                           """, (
                               test_clean, machine_clean, machine_id,
                               int(row.get('is_selected', 1)),
                               float(row.get('target_conc', 0)),
                               unit_str,
                               float(row.get('tea_pct', 0)),
                               float(row.get('swrl_pct', 0)),
                               float(row.get('sr_pct', 0)),
                               float(row.get('cd_pct', 0)),
                               float(row.get('cd_swrl', 0)),
                               float(row.get('sr_swrl', 0)),
                               float(row.get('power_achieve', 0)),
                               int(row.get('n_samples', 0)),
                               float(row.get('rejection_limit', 0)),
                               float(row.get('k_factor', 0)),
                               note_str,
                               vn_now
                           ))
            imported_count += 1
        conn.commit()
    return imported_count