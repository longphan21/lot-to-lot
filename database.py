import sqlite3
import pandas as pd
from pathlib import Path

DB_PATH = Path(__file__).parent / "lot_to_lot_data.db"


def get_connection():
    return sqlite3.connect(DB_PATH)


def init_db():
    """Khởi tạo bảng lưu thông số Stage 1 Setup nếu chưa có."""
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stage1_setup (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_name TEXT NOT NULL,
                machine TEXT NOT NULL,
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
        conn.commit()


def save_test_setup(test_name: str, machine: str, df_rows: pd.DataFrame):
    """Lưu đè bộ thông số của 1 xét nghiệm trên 1 dòng máy."""
    init_db()
    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()

    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM stage1_setup WHERE UPPER(test_name) = ? AND machine = ?",
            (test_clean, machine_clean)
        )
        for _, row in df_rows.iterrows():
            cursor.execute("""
                INSERT INTO stage1_setup (
                    test_name, machine, is_selected, target_conc, unit,
                    tea_pct, swrl_pct, sr_pct, cd_pct, cd_swrl, sr_swrl,
                    power_achieve, n_samples, rejection_limit, k_factor, note
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                test_clean,
                machine_clean,
                1 if bool(row.get("Chọn chạy (Vàng)", False)) else 0,
                float(row["Target concentration"]),
                str(row.get("Đơn vị", "mmol/L")),
                float(row["TEa%"]),
                float(row["Swrl = CVwrl %"]),
                float(row["Sr = CVr %"]),
                float(row["CD% (Critical Difference)"]),
                float(row["CD/Swrl"]),
                float(row["Sr/Swrl"]),
                float(row["Power achieve"]),
                int(row["Number of sample each concentration"]),
                float(row["Rejection limit"]),
                float(row["k"]),
                str(row.get("Ghi chú", ""))
            ))
        conn.commit()


def load_all_setups() -> pd.DataFrame:
    """Đọc toàn bộ dữ liệu Setup đã lưu."""
    init_db()
    with get_connection() as conn:
        df = pd.read_sql_query(
            "SELECT * FROM stage1_setup ORDER BY test_name, machine, target_conc",
            conn
        )
    return df


def load_specific_setup(test_name: str, machine: str) -> pd.DataFrame:
    """Đọc thông số của riêng 1 xét nghiệm và 1 máy."""
    init_db()
    with get_connection() as conn:
        df = pd.read_sql_query(
            "SELECT * FROM stage1_setup WHERE UPPER(test_name) = ? AND machine = ? ORDER BY target_conc",
            conn,
            params=(test_name.strip().upper(), machine.strip())
        )
    return df


def delete_specific_setup(test_name: str, machine: str):
    """Xóa cấu hình của 1 xét nghiệm trên 1 máy."""
    init_db()
    with get_connection() as conn:
        conn.execute(
            "DELETE FROM stage1_setup WHERE UPPER(test_name) = ? AND machine = ?",
            (test_name.strip().upper(), machine.strip())
        )
        conn.commit()