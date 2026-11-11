import streamlit as st
import gspread
import pandas as pd
from datetime import datetime, timedelta


# ==========================================
# 1. HÀM KẾT NỐI GOOGLE SHEETS
# ==========================================
ddef get_gsheet_client():
    if "gcp_service_account" in st.secrets:
        # Chuyển secrets thành dictionary để có thể chỉnh sửa
        creds_dict = dict(st.secrets["gcp_service_account"])

        # SỬA LỖI VALUE_ERROR Ở ĐÂY: Ép các ký tự '\n' thành xuống dòng chuẩn
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        client = gspread.service_account_from_dict(creds_dict)
    else:
        # Chạy local trên PyCharm
        client = gspread.service_account(filename="credentials.json")

    # Tên file Google Sheet
    spreadsheet = client.open("Lot_To_Lot_DB")
    return spreadsheet


def init_db():
    try:
        sh = get_gsheet_client()

        # 1. Kiểm tra / Tạo sheet stage1_setup
        try:
            sh.worksheet("stage1_setup")
        except gspread.WorksheetNotFound:
            ws1 = sh.add_worksheet(title="stage1_setup", rows=1000, cols=20)
            ws1.append_row([
                "id", "test_name", "machine", "machine_id", "is_selected",
                "target_conc", "unit", "tea_pct", "swrl_pct", "sr_pct",
                "cd_pct", "cd_swrl", "sr_swrl", "power_achieve",
                "n_samples", "rejection_limit", "k_factor", "note", "updated_at"
            ])

        # 2. Kiểm tra / Tạo sheet stage2_evaluations
        try:
            sh.worksheet("stage2_evaluations")
        except gspread.WorksheetNotFound:
            ws2 = sh.add_worksheet(title="stage2_evaluations", rows=1000, cols=25)
            ws2.append_row([
                "id", "eval_date", "test_name", "machine", "machine_id", "unit",
                "lot_old", "lot_new", "level_name", "target_conc", "sid",
                "val_old", "val_new", "diff_d", "mean_d", "mean_d_pct",
                "rl_pct", "k_factor", "status", "reagent_name", "created_at"
            ])
    except Exception as e:
        print(f"Lỗi khởi tạo Google Sheets: {e}")


# ==========================================
# 2. CÁC HÀM QUẢN LÝ SETUP (STAGE 1)
# ==========================================
def save_test_setup(test_name: str, machine: str, machine_id: str, df_rows: pd.DataFrame):
    init_db()
    sh = get_gsheet_client()
    ws = sh.worksheet("stage1_setup")
    data = ws.get_all_records()

    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()
    mach_id_clean = machine_id.strip()
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')

    # Lọc bỏ dữ liệu cũ của test & machine này (Giống hệt câu DELETE FROM)
    filtered_data = [
        row for row in data
        if not (str(row.get("test_name", "")).strip().upper() == test_clean and str(
            row.get("machine", "")).strip() == machine_clean)
    ]

    # Thêm dữ liệu mới từ df_rows
    for _, row in df_rows.iterrows():
        new_row = {
            "id": datetime.utcnow().strftime('%Y%m%d%H%M%S%f'),  # Tạo ID tạm thay cho SERIAL
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

    # Ghi toàn bộ data mới lên Sheets
    ws.clear()
    if filtered_data:
        df_to_save = pd.DataFrame(filtered_data).fillna("")  # fillna("") để tránh lỗi NaN khi dump JSON
        ws.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())


def load_all_setups() -> pd.DataFrame:
    init_db()
    sh = get_gsheet_client()
    data = sh.worksheet("stage1_setup").get_all_records()
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data)
    # Sắp xếp như SQL: ORDER BY test_name, machine, target_conc
    sort_cols = [c for c in ['test_name', 'machine', 'target_conc'] if c in df.columns]
    if sort_cols:
        df = df.sort_values(by=sort_cols)
    return df


def load_specific_setup(test_name: str, machine: str) -> pd.DataFrame:
    df = load_all_setups()
    if df.empty:
        return df

    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()

    df['test_name'] = df['test_name'].astype(str).str.strip().str.upper()
    df['machine'] = df['machine'].astype(str).str.strip()

    filtered = df[(df['test_name'] == test_clean) & (df['machine'] == machine_clean)]
    return filtered


def delete_specific_setup(test_name: str, machine: str):
    init_db()
    sh = get_gsheet_client()
    ws = sh.worksheet("stage1_setup")
    data = ws.get_all_records()
    if not data:
        return

    test_clean = test_name.strip().upper()
    machine_clean = machine.strip()

    filtered_data = [
        row for row in data
        if not (str(row.get("test_name", "")).strip().upper() == test_clean and str(
            row.get("machine", "")).strip() == machine_clean)
    ]

    ws.clear()
    if filtered_data:
        df_to_save = pd.DataFrame(filtered_data).fillna("")
        ws.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
    else:
        # Nếu xóa hết thì tạo lại header
        ws.append_row([
            "id", "test_name", "machine", "machine_id", "is_selected", "target_conc", "unit",
            "tea_pct", "swrl_pct", "sr_pct", "cd_pct", "cd_swrl", "sr_swrl", "power_achieve",
            "n_samples", "rejection_limit", "k_factor", "note", "updated_at"
        ])


# ==========================================
# 3. CÁC HÀM QUẢN LÝ LỊCH SỬ (STAGE 2)
# ==========================================
def save_stage2_evaluation(records: list):
    init_db()
    sh = get_gsheet_client()
    ws = sh.worksheet("stage2_evaluations")
    data = ws.get_all_records()
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')

    for r in records:
        record_time = r.get('created_at') if r.get('created_at') else vn_now
        new_row = {
            "id": datetime.utcnow().strftime('%Y%m%d%H%M%S%f'),
            "eval_date": r.get('eval_date', ''),
            "test_name": r.get('test_name', ''),
            "machine": r.get('machine', ''),
            "machine_id": r.get('machine_id', ''),
            "unit": r.get('unit', ''),
            "lot_old": r.get('lot_old', ''),
            "lot_new": r.get('lot_new', ''),
            "level_name": r.get('level_name', ''),
            "target_conc": r.get('target_conc', 0),
            "sid": r.get('sid', ''),
            "val_old": r.get('val_old', 0),
            "val_new": r.get('val_new', 0),
            "diff_d": r.get('diff_d', 0),
            "mean_d": r.get('mean_d', 0),
            "mean_d_pct": r.get('mean_d_pct', 0),
            "rl_pct": r.get('rl_pct', 0),
            "k_factor": r.get('k_factor', 0),
            "status": r.get('status', ''),
            "reagent_name": r.get('reagent_name', ''),
            "created_at": record_time
        }
        data.append(new_row)

    ws.clear()
    if data:
        df_to_save = pd.DataFrame(data).fillna("")
        ws.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())


def load_stage2_history() -> pd.DataFrame:
    init_db()
    sh = get_gsheet_client()
    data = sh.worksheet("stage2_evaluations").get_all_records()
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data)

    # ORDER BY eval_date DESC, id DESC
    if not df.empty and 'eval_date' in df.columns and 'id' in df.columns:
        df = df.sort_values(by=['eval_date', 'id'], ascending=[False, False])
    return df


def delete_stage2_evaluation(created_at: str):
    init_db()
    sh = get_gsheet_client()
    ws = sh.worksheet("stage2_evaluations")
    data = ws.get_all_records()
    if not data:
        return

    filtered_data = [row for row in data if str(row.get("created_at")) != str(created_at)]

    ws.clear()
    if filtered_data:
        df_to_save = pd.DataFrame(filtered_data).fillna("")
        ws.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
    else:
        ws.append_row([
            "id", "eval_date", "test_name", "machine", "machine_id", "unit",
            "lot_old", "lot_new", "level_name", "target_conc", "sid",
            "val_old", "val_new", "diff_d", "mean_d", "mean_d_pct",
            "rl_pct", "k_factor", "status", "reagent_name", "created_at"
        ])


# ==========================================
# 4. HÀM IMPORT DỮ LIỆU TỪ EXCEL VÀO DB
# ==========================================
def import_stage1_from_dataframe(df: pd.DataFrame) -> int:
    init_db()
    sh = get_gsheet_client()
    ws = sh.worksheet("stage1_setup")
    data = ws.get_all_records()

    imported_count = 0
    vn_now = (datetime.utcnow() + timedelta(hours=7)).strftime('%Y-%m-%d %H:%M:%S')

    # Tìm các nhóm (test_name, machine) để xóa khỏi data cũ
    unique_groups = df[['test_name', 'machine']].drop_duplicates()
    groups_to_delete = []
    for _, group in unique_groups.iterrows():
        t_name = str(group['test_name']).strip().upper() if pd.notna(group['test_name']) else ""
        m_name = str(group['machine']).strip() if pd.notna(group['machine']) else ""
        if t_name and m_name:
            groups_to_delete.append((t_name, m_name))

    # Lọc bỏ data cũ
    filtered_data = []
    for row in data:
        r_test = str(row.get("test_name", "")).strip().upper()
        r_mach = str(row.get("machine", "")).strip()
        if (r_test, r_mach) not in groups_to_delete:
            filtered_data.append(row)

    # Nhồi data mới vào
    for _, row in df.iterrows():
        test_clean = str(row['test_name']).strip().upper() if pd.notna(row.get('test_name')) else ""
        machine_clean = str(row['machine']).strip() if pd.notna(row.get('machine')) else ""

        if not test_clean or not machine_clean:
            continue

        machine_id = str(row.get('machine_id', '')).strip() if pd.notna(row.get('machine_id')) else ""
        unit_str = str(row.get('unit', '')).strip() if pd.notna(row.get('unit')) else "mmol/L"
        note_str = str(row.get('note', '')).strip() if pd.notna(row.get('note')) else ""

        new_row = {
            "id": datetime.utcnow().strftime('%Y%m%d%H%M%S%f'),
            "test_name": test_clean,
            "machine": machine_clean,
            "machine_id": machine_id,
            "is_selected": int(row.get('is_selected', 1)),
            "target_conc": float(row.get('target_conc', 0)),
            "unit": unit_str,
            "tea_pct": float(row.get('tea_pct', 0)),
            "swrl_pct": float(row.get('swrl_pct', 0)),
            "sr_pct": float(row.get('sr_pct', 0)),
            "cd_pct": float(row.get('cd_pct', 0)),
            "cd_swrl": float(row.get('cd_swrl', 0)),
            "sr_swrl": float(row.get('sr_swrl', 0)),
            "power_achieve": float(row.get('power_achieve', 0)),
            "n_samples": int(row.get('n_samples', 0)),
            "rejection_limit": float(row.get('rejection_limit', 0)),
            "k_factor": float(row.get('k_factor', 0)),
            "note": note_str,
            "updated_at": vn_now
        }
        filtered_data.append(new_row)
        imported_count += 1

    # Ghi đè lại lên sheet
    ws.clear()
    if filtered_data:
        df_to_save = pd.DataFrame(filtered_data).fillna("")
        ws.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
    else:
        ws.append_row([
            "id", "test_name", "machine", "machine_id", "is_selected", "target_conc", "unit",
            "tea_pct", "swrl_pct", "sr_pct", "cd_pct", "cd_swrl", "sr_swrl", "power_achieve",
            "n_samples", "rejection_limit", "k_factor", "note", "updated_at"
        ])
    return imported_count