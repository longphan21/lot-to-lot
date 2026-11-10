import sqlite3
import pandas as pd
from pathlib import Path

DB_PATH = Path(__file__).parent / "lot_to_lot_data.db"

def get_connection():
    return sqlite3.connect(DB_PATH)

def init_db():
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stage1_setup (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_name TEXT NOT NULL,
                machine TEXT NOT NULL,
                machine_id TEXT,
                is_selected INTEGER DEFAULT 1,
                target_conc REAL NOT NULL,
                unit TEXT DEFAULT 'mmol/L',
                tea_pct REAL NOT NULL,
                swrl_pct REAL NOT NULL,
                sr_pct REAL NOT NULL,
                cd_pct REAL,
                cd_swrl REAL,
                sr_swrl REAL,
                power_achieve REAL,
                n_samples INTEGER,
                rejection_limit REAL,
                k_factor REAL,
                note TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stage2_evaluations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                eval_date TEXT NOT NULL,
                test_name TEXT NOT NULL,
                machine TEXT NOT NULL,
                machine_id TEXT,
                unit TEXT,
                lot_old TEXT NOT NULL,
                lot_new TEXT NOT NULL,
                level_name TEXT,
                target_conc REAL,
                sid TEXT,
                val_old REAL,
                val_new REAL,
                diff_d REAL,
                mean_d REAL,
                mean_d_pct REAL,
                rl_pct REAL,
                k_factor REAL,
                status TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

def save_test_setup(test_name: str, machine: str, machine_id: str, df_rows: pd.DataFrame):
    init_db()
    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()
    mach_id_clean = machine_id.strip()

    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM stage1_setup WHERE UPPER(test_name) = ? AND machine = ?",
            (test_clean, machine_clean)
        )
        for _, row in df_rows.iterrows():
            cursor.execute("""
                INSERT INTO stage1_setup (
                    test_name, machine, machine_id, is_selected, target_conc, unit,
                    tea_pct, swrl_pct, sr_pct, cd_pct, cd_swrl, sr_swrl,
                    power_achieve, n_samples, rejection_limit, k_factor, note
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                test_clean, machine_clean, mach_id_clean,
                1 if bool(row.get("Chọn chạy (Vàng)", False)) else 0,
                float(row["Target concentration"]), str(row.get("Đơn vị", "mmol/L")),
                float(row["TEa%"]), float(row["Swrl = CVwrl %"]), float(row["Sr = CVr %"]),
                float(row["CD% (Critical Difference)"]), float(row["CD/Swrl"]),
                float(row["Sr/Swrl"]), float(row["Power achieve"]),
                int(row["Number of sample each concentration"]),
                float(row["Rejection limit"]), float(row["k"]), str(row.get("Ghi chú", ""))
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
            "SELECT * FROM stage1_setup WHERE UPPER(test_name) = ? AND machine = ? ORDER BY target_conc",
            conn, params=(test_name.strip().upper(), machine.strip())
        )

def delete_specific_setup(test_name: str, machine: str):
    init_db()
    with get_connection() as conn:
        conn.execute("DELETE FROM stage1_setup WHERE UPPER(test_name) = ? AND machine = ?", (test_name.strip().upper(), machine.strip()))
        conn.commit()

def save_stage2_evaluation(records: list):
    init_db()
    with get_connection() as conn:
        cursor = conn.cursor()
        for r in records:
            cursor.execute("""
                INSERT INTO stage2_evaluations (
                    eval_date, test_name, machine, machine_id, unit, lot_old, lot_new,
                    level_name, target_conc, sid, val_old, val_new, diff_d,
                    mean_d, mean_d_pct, rl_pct, k_factor, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                r["eval_date"], r["test_name"], r["machine"], r["machine_id"], r["unit"],
                r["lot_old"], r["lot_new"], r["level_name"], r["target_conc"],
                r["sid"], r["val_old"], r["val_new"], r["diff_d"],
                r["mean_d"], r["mean_d_pct"], r["rl_pct"], r["k_factor"], r["status"]
            ))
        conn.commit()

def load_stage2_history() -> pd.DataFrame:
    init_db()
    with get_connection() as conn:
        return pd.read_sql_query("SELECT * FROM stage2_evaluations ORDER BY eval_date DESC, id DESC", conn)

def delete_stage2_evaluation(created_at: str):
    init_db()
    with get_connection() as conn:
        conn.execute("DELETE FROM stage2_evaluations WHERE created_at = ?", (created_at,))
        conn.commit()

def import_stage1_from_dataframe(df: pd.DataFrame) -> int:
    """Đọc dữ liệu từ file CSV/Excel và đẩy vào Database (Xử lý được cả file từ phiên bản cũ)"""
    init_db()
    imported_count = 0
    with get_connection() as conn:
        cursor = conn.cursor()

        # 1. Nhóm theo Tên Xét Nghiệm và Dòng Máy để xóa dữ liệu cũ (tránh trùng lặp khi import 2 lần)
        unique_groups = df[['test_name', 'machine']].drop_duplicates()
        for _, group in unique_groups.iterrows():
            t_name = str(group['test_name']).strip().upper() if not pd.isna(group['test_name']) else ""
            m_name = str(group['machine']).strip() if not pd.isna(group['machine']) else ""
            if t_name and m_name:
                cursor.execute("DELETE FROM stage1_setup WHERE UPPER(test_name) = ? AND machine = ?", (t_name, m_name))

        # 2. Quét từng dòng của file CSV để nạp vào DB
        for _, row in df.iterrows():
            test_clean = str(row['test_name']).strip().upper() if not pd.isna(row.get('test_name')) else ""
            machine_clean = str(row['machine']).strip() if not pd.isna(row.get('machine')) else ""

            if not test_clean or not machine_clean:
                continue

            # Xử lý tương thích ngược: Nếu CSV cũ chưa có cột machine_id thì để trống
            machine_id = str(row['machine_id']).strip() if 'machine_id' in df.columns and not pd.isna(
                row['machine_id']) else ""
            unit_str = str(row['unit']).strip() if 'unit' in df.columns and not pd.isna(row['unit']) else "mmol/L"
            note_str = str(row['note']).strip() if 'note' in df.columns and not pd.isna(row['note']) else ""

            cursor.execute("""
                           INSERT INTO stage1_setup (test_name, machine, machine_id, is_selected, target_conc, unit,
                                                     tea_pct, swrl_pct, sr_pct, cd_pct, cd_swrl, sr_swrl,
                                                     power_achieve, n_samples, rejection_limit, k_factor, note)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                               note_str
                           ))
            imported_count += 1
        conn.commit()
    return imported_count