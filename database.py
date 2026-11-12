import streamlit as st
import gspread
import pandas as pd
from datetime import datetime, timedelta
from connection import get_spreadsheet

STAGE1_HEADERS = [
    "id", "test_name", "machine", "machine_id", "is_selected",
    "target_conc", "unit", "tea_pct", "swrl_pct", "sr_pct",
    "cd_pct", "cd_swrl", "sr_swrl", "power_achieve",
    "n_samples", "rejection_limit", "k_factor", "note", "updated_at"
]

STAGE2_HEADERS = [
    "id", "eval_date", "test_name", "machine", "machine_id", "unit",
    "lot_old", "lot_new", "level_name", "target_conc", "sid",
    "val_old", "val_new", "diff_d", "mean_d", "mean_d_pct",
    "rl_pct", "k_factor", "status", "reagent_name", "created_at"
]

@st.cache_resource(show_spinner=False)
def init_db():
    sh = get_spreadsheet()

    try:
        ws1 = sh.worksheet("stage1_setup")
        if not ws1.row_values(1):
            ws1.append_row(STAGE1_HEADERS)
    except gspread.WorksheetNotFound:
        ws1 = sh.add_worksheet(title="stage1_setup", rows=1000, cols=20)
        ws1.append_row(STAGE1_HEADERS)

    try:
        ws2 = sh.worksheet("stage2_evaluations")
        if not ws2.row_values(1):
            ws2.append_row(STAGE2_HEADERS)
    except gspread.WorksheetNotFound:
        ws2 = sh.add_worksheet(title="stage2_evaluations", rows=1000, cols=25)
        ws2.append_row(STAGE2_HEADERS)

    # KHỞI TẠO THÊM SHEET THÙNG RÁC
    try:
        ws3 = sh.worksheet("stage2_trash")
        if not ws3.row_values(1):
            ws3.append_row(STAGE2_HEADERS)
    except gspread.WorksheetNotFound:
        ws3 = sh.add_worksheet(title="stage2_trash", rows=1000, cols=25)
        ws3.append_row(STAGE2_HEADERS)

    return True

# ==========================================
# 1. CÁC HÀM QUẢN LÝ SETUP (STAGE 1)
# ==========================================
@st.cache_data(ttl=60, show_spinner=False)
def load_all_setups() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = sh.worksheet("stage1_setup")
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=STAGE1_HEADERS)

    df = pd.DataFrame(data)
    num_cols = [
        "is_selected", "target_conc", "tea_pct", "swrl_pct", "sr_pct",
        "cd_pct", "cd_swrl", "sr_swrl", "power_achieve", "n_samples",
        "rejection_limit", "k_factor"
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "is_selected" in df.columns:
        df["is_selected"] = df["is_selected"].astype(int)

    sort_cols = [c for c in ["test_name", "machine", "target_conc"] if c in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols).reset_index(drop=True)
    return df

def load_specific_setup(test_name: str, machine: str) -> pd.DataFrame:
    df = load_all_setups()
    if df.empty:
        return df
    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()
    mask = (
        (df["test_name"].astype(str).str.strip().str.upper() == test_clean) &
        (df["machine"].astype(str).str.strip() == machine_clean)
    )
    return df[mask].copy()

def save_test_setup(test_name: str, machine: str, machine_id: str, df_rows: pd.DataFrame):
    sh = get_spreadsheet()
    ws = sh.worksheet("stage1_setup")
    data = ws.get_all_records()
    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()
    mach_id_clean = machine_id.strip()
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')
    base_id = datetime.utcnow().strftime('%Y%m%d%H%M%S')

    filtered_data = [
        row for row in data
        if not (str(row.get("test_name", "")).strip().upper() == test_clean and str(row.get("machine", "")).strip() == machine_clean)
    ]

    for idx, (_, row) in enumerate(df_rows.iterrows()):
        new_row = {
            "id": f"{base_id}_{idx}",
            "test_name": test_clean,
            "machine": machine_clean,
            "machine_id": mach_id_clean,
            "is_selected": 1 if bool(row.get("Chọn chạy (Vàng)", False)) else 0,
            "target_conc": float(row.get("Target concentration", 0)),
            "unit": str(row.get("unit", "mmol/L")),
            "tea_pct": float(row.get("TEa%", 0)),
            "swrl_pct": float(row.get("Swrl = CVwrl %", 0)),
            "sr_pct": float(row.get("Sr = CVr %", 0)),
            "cd_pct": 0.0,
            "cd_swrl": float(row.get("cd_swrl", 0)),
            "sr_swrl": float(row.get("sr_swrl", 0)),
            "power_achieve": float(row.get("power_achieve", 0)),
            "n_samples": int(row.get("n_samples", 0)),
            "rejection_limit": float(row.get("rejection_limit", 0)),
            "k_factor": float(row.get("k_factor", 0)),
            "note": str(row.get("Ghi chú", "")),
            "updated_at": vn_now
        }
        filtered_data.append(new_row)

    ws.clear()
    if filtered_data:
        df_to_save = pd.DataFrame(filtered_data).reindex(columns=STAGE1_HEADERS).fillna("")
        ws.update([STAGE1_HEADERS] + df_to_save.values.tolist())
    else:
        ws.append_row(STAGE1_HEADERS)
    load_all_setups.clear()

def delete_specific_setup(test_name: str, machine: str):
    sh = get_spreadsheet()
    ws = sh.worksheet("stage1_setup")
    data = ws.get_all_records()
    if not data: return
    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()
    filtered_data = [
        row for row in data
        if not (str(row.get("test_name", "")).strip().upper() == test_clean and str(row.get("machine", "")).strip() == machine_clean)
    ]
    ws.clear()
    if filtered_data:
        df_to_save = pd.DataFrame(filtered_data).reindex(columns=STAGE1_HEADERS).fillna("")
        ws.update([STAGE1_HEADERS] + df_to_save.values.tolist())
    else:
        ws.append_row(STAGE1_HEADERS)
    load_all_setups.clear()

# ==========================================
# 2. CÁC HÀM QUẢN LÝ LỊCH SỬ VÀ THÙNG RÁC (STAGE 2)
# ==========================================
@st.cache_data(ttl=60, show_spinner=False)
def load_stage2_history() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = sh.worksheet("stage2_evaluations")
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=STAGE2_HEADERS)

    df = pd.DataFrame(data)
    num_cols = ["target_conc", "val_old", "val_new", "diff_d", "mean_d", "mean_d_pct", "rl_pct", "k_factor"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    if not df.empty and "eval_date" in df.columns and "id" in df.columns:
        df = df.sort_values(by=["eval_date", "id"], ascending=[False, False]).reset_index(drop=True)
    return df

@st.cache_data(ttl=60, show_spinner=False)
def load_trash_history() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = sh.worksheet("stage2_trash")
    data = ws.get_all_records()
    if not data: return pd.DataFrame(columns=STAGE2_HEADERS)
    df = pd.DataFrame(data)
    return df

def save_stage2_evaluation(records: list):
    sh = get_spreadsheet()
    ws = sh.worksheet("stage2_evaluations")
    data = ws.get_all_records()
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')
    base_id = datetime.utcnow().strftime('%Y%m%d%H%M%S')

    for idx, r in enumerate(records):
        record_time = r.get("created_at") if r.get("created_at") else vn_now
        new_row = {
            "id": f"{base_id}_{idx}",
            "eval_date": r.get("eval_date", ""),
            "test_name": r.get("test_name", ""),
            "machine": r.get("machine", ""),
            "machine_id": r.get("machine_id", ""),
            "unit": r.get("unit", ""),
            "lot_old": r.get("lot_old", ""),
            "lot_new": r.get("lot_new", ""),
            "level_name": r.get("level_name", ""),
            "target_conc": float(r.get("target_conc", 0)),
            "sid": r.get("sid", ""),
            "val_old": float(r.get("val_old", 0)),
            "val_new": float(r.get("val_new", 0)),
            "diff_d": float(r.get("diff_d", 0)),
            "mean_d": float(r.get("mean_d", 0)),
            "mean_d_pct": float(r.get("mean_d_pct", 0)),
            "rl_pct": float(r.get("rl_pct", 0)),
            "k_factor": float(r.get("k_factor", 0)),
            "status": r.get("status", ""),
            "reagent_name": r.get("reagent_name", ""),
            "created_at": record_time
        }
        data.append(new_row)

    ws.clear()
    if data:
        df_to_save = pd.DataFrame(data).reindex(columns=STAGE2_HEADERS).fillna("")
        ws.update([STAGE2_HEADERS] + df_to_save.values.tolist())
    else:
        ws.append_row(STAGE2_HEADERS)
    load_stage2_history.clear()

def delete_stage2_evaluation(created_at: str):
    """Xóa từ bảng chính và lưu trữ vào bảng thùng rác (stage2_trash)"""
    sh = get_spreadsheet()
    ws = sh.worksheet("stage2_evaluations")
    ws_trash = sh.worksheet("stage2_trash")
    data = ws.get_all_records()
    if not data: return

    to_keep = []
    to_trash = []
    for row in data:
        if str(row.get("created_at")) == str(created_at):
            to_trash.append(row)
        else:
            to_keep.append(row)

    if to_trash:
        trash_data = ws_trash.get_all_records()
        trash_data.extend(to_trash)
        df_trash = pd.DataFrame(trash_data).reindex(columns=STAGE2_HEADERS).fillna("")
        ws_trash.clear()
        ws_trash.update([STAGE2_HEADERS] + df_trash.values.tolist())

    ws.clear()
    if to_keep:
        df_to_save = pd.DataFrame(to_keep).reindex(columns=STAGE2_HEADERS).fillna("")
        ws.update([STAGE2_HEADERS] + df_to_save.values.tolist())
    else:
        ws.append_row(STAGE2_HEADERS)

    load_stage2_history.clear()
    load_trash_history.clear()

def restore_stage2_evaluation(created_at: str):
    """Khôi phục dữ liệu từ thùng rác về bảng chính"""
    sh = get_spreadsheet()
    ws = sh.worksheet("stage2_evaluations")
    ws_trash = sh.worksheet("stage2_trash")
    trash_data = ws_trash.get_all_records()
    if not trash_data: return

    to_keep_in_trash = []
    to_restore = []
    for row in trash_data:
        if str(row.get("created_at")) == str(created_at):
            to_restore.append(row)
        else:
            to_keep_in_trash.append(row)

    if to_restore:
        main_data = ws.get_all_records()
        main_data.extend(to_restore)
        df_main = pd.DataFrame(main_data).reindex(columns=STAGE2_HEADERS).fillna("")
        ws.clear()
        ws.update([STAGE2_HEADERS] + df_main.values.tolist())

    ws_trash.clear()
    if to_keep_in_trash:
        df_trash = pd.DataFrame(to_keep_in_trash).reindex(columns=STAGE2_HEADERS).fillna("")
        ws_trash.update([STAGE2_HEADERS] + df_trash.values.tolist())
    else:
        ws_trash.append_row(STAGE2_HEADERS)

    load_stage2_history.clear()
    load_trash_history.clear()

# ==========================================
# 3. HÀM IMPORT DỮ LIỆU TỪ EXCEL VÀO DB
# ==========================================
def import_stage1_from_dataframe(df: pd.DataFrame) -> int:
    sh = get_spreadsheet()
    ws = sh.worksheet("stage1_setup")
    data = ws.get_all_records()
    imported_count = 0
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')
    base_id = datetime.utcnow().strftime('%Y%m%d%H%M%S')

    unique_groups = df[["test_name", "machine"]].drop_duplicates()
    groups_to_delete = set()
    for _, group in unique_groups.iterrows():
        t_name = str(group["test_name"]).strip().upper() if pd.notna(group["test_name"]) else ""
        m_name = str(group["machine"]).strip() if pd.notna(group["machine"]) else ""
        if t_name and m_name: groups_to_delete.add((t_name, m_name))

    filtered_data = [
        row for row in data
        if (str(row.get("test_name", "")).strip().upper(), str(row.get("machine", "")).strip()) not in groups_to_delete
    ]

    for idx, (_, row) in enumerate(df.iterrows()):
        test_clean = str(row["test_name"]).strip().upper() if pd.notna(row.get("test_name")) else ""
        machine_clean = str(row["machine"]).strip() if pd.notna(row.get("machine")) else ""
        if not test_clean or not machine_clean: continue
        new_row = {
            "id": f"{base_id}_{idx}",
            "test_name": test_clean,
            "machine": machine_clean,
            "machine_id": str(row.get("machine_id", "")).strip() if pd.notna(row.get("machine_id")) else "",
            "is_selected": int(row.get("is_selected", 1)),
            "target_conc": float(row.get("target_conc", 0)),
            "unit": str(row.get("unit", "")).strip() if pd.notna(row.get("unit")) else "mmol/L",
            "tea_pct": float(row.get("tea_pct", 0)),
            "swrl_pct": float(row.get("swrl_pct", 0)),
            "sr_pct": float(row.get("sr_pct", 0)),
            "cd_pct": float(row.get("cd_pct", 0)),
            "cd_swrl": float(row.get("cd_swrl", 0)),
            "sr_swrl": float(row.get("sr_swrl", 0)),
            "power_achieve": float(row.get("power_achieve", 0)),
            "n_samples": int(row.get("n_samples", 0)),
            "rejection_limit": float(row.get("rejection_limit", 0)),
            "k_factor": float(row.get("k_factor", 0)),
            "note": str(row.get("note", "")).strip() if pd.notna(row.get("note")) else "",
            "updated_at": vn_now
        }
        filtered_data.append(new_row)
        imported_count += 1

    ws.clear()
    if filtered_data:
        df_to_save = pd.DataFrame(filtered_data).reindex(columns=STAGE1_HEADERS).fillna("")
        ws.update([STAGE1_HEADERS] + df_to_save.values.tolist())
    else:
        ws.append_row(STAGE1_HEADERS)

    load_all_setups.clear()
    return imported_count