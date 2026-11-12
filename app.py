import subprocess
import tempfile
import os
import time
from datetime import date
import streamlit as st
import pandas as pd
from clsi_engine import get_candidate_options_for_row, evaluate_stage2_level
from report_generator import generate_report_docx
from database import (
    init_db, save_test_setup, load_all_setups, load_specific_setup,
    delete_specific_setup, save_stage2_evaluation, load_stage2_history,
    delete_stage2_evaluation, import_stage1_from_dataframe,
    load_trash_history, restore_stage2_evaluation, permanently_delete_stage2_evaluation
)

st.set_page_config(
    page_title="CLSI EP26 - Lot-to-Lot",
    layout="wide",
    initial_sidebar_state="expanded"
)

DEFAULT_SAMPLES = {
    ("CHOLESTEROL", "Alinity C"): pd.DataFrame([
        {
            "Chọn chạy (Vàng)": True,
            "Target concentration": 6.37,
            "TEa%": 12.9,
            "Swrl = CVwrl %": 0.5,
            "Sr = CVr %": 0.3
        },
    ])
}


def inject_custom_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        /* 1. TỔNG THỂ NỀN & FONT CHỮ */
        html, body, [class*="css"], .stApp {
            background-color: #F1F5F9 !important;
            color: #0F172A !important;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }

        .block-container {
            padding-top: 1.5rem !important;
            padding-bottom: 2.5rem !important;
            max-width: 1440px !important;
        }

        /* Định dạng số liệu thẳng hàng dọc (tabular-nums) */
        .tabular-nums, [data-testid="stMetricValue"], [data-testid="stDataFrame"], table td {
            font-variant-numeric: tabular-nums !important;
        }

        /* 2. PHÂN CẤP TIÊU ĐỀ (Bỏ in hoa toàn bộ, độ đậm vừa phải) */
        h1, h2, h3, h4, h5, h6 {
            color: #0F172A !important;
            font-family: 'Inter', sans-serif !important;
            letter-spacing: -0.01em;
            margin-top: 0 !important;
        }
        h3 {
            font-size: 1.125rem !important;
            font-weight: 600 !important;
            color: #0F172A !important;
            margin-bottom: 0.75rem !important;
        }
        h4, h5 {
            font-size: 1.0rem !important;
            font-weight: 600 !important;
            color: #0F172A !important;
            margin-bottom: 0.5rem !important;
        }

        /* 3. SIDEBAR MENU DỌC (Nền xanh Navy đậm #0F172A) */
        [data-testid="stSidebar"] {
            background-color: #0F172A !important;
            border-right: 1px solid #1E293B !important;
        }
        [data-testid="stSidebar"] > div:first-child {
            background-color: #0F172A !important;
            padding-top: 1.25rem !important;
        }
        .sidebar-brand {
            padding: 0 0.5rem 1.25rem 0.5rem;
            border-bottom: 1px solid #1E293B;
            margin-bottom: 1rem;
        }
        .sidebar-brand-title {
            color: #FFFFFF;
            font-size: 1.15rem;
            font-weight: 700;
            letter-spacing: -0.01em;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .sidebar-brand-sub {
            color: #94A3B8;
            font-size: 0.8rem;
            font-weight: 400;
            margin-top: 4px;
        }
        .sidebar-section-label {
            color: #64748B;
            font-size: 0.72rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            padding: 0 0.5rem;
            margin-bottom: 0.5rem;
        }

        /* Biến st.radio trong Sidebar thành thanh Menu dọc chuyên nghiệp */
        [data-testid="stSidebar"] .stRadio > label {
            display: none !important;
        }
        [data-testid="stSidebar"] [role="radiogroup"] {
            gap: 6px !important;
        }
        [data-testid="stSidebar"] [role="radiogroup"] label[data-baseweb="radio"] {
            background-color: transparent !important;
            padding: 10px 12px !important;
            border-radius: 8px !important;
            border-left: 3px solid transparent !important;
            transition: background-color 0.15s ease, color 0.15s ease !important;
            cursor: pointer !important;
            margin-bottom: 0 !important;
            width: 100% !important;
        }
        /* Ẩn nút tròn mặc định của radio trong sidebar */
        [data-testid="stSidebar"] [role="radiogroup"] label[data-baseweb="radio"] > div:first-child {
            display: none !important;
        }
        [data-testid="stSidebar"] [role="radiogroup"] label[data-baseweb="radio"] p {
            color: #94A3B8 !important;
            font-size: 0.9rem !important;
            font-weight: 500 !important;
            margin: 0 !important;
        }
        [data-testid="stSidebar"] [role="radiogroup"] label[data-baseweb="radio"]:hover {
            background-color: #1E293B !important;
        }
        [data-testid="stSidebar"] [role="radiogroup"] label[data-baseweb="radio"]:hover p {
            color: #F8FAFC !important;
        }
        /* Mục đang chọn trên Sidebar */
        [data-testid="stSidebar"] [role="radiogroup"] label[data-baseweb="radio"]:has(input:checked) {
            background-color: #1E293B !important;
            border-left: 3px solid #38BDF8 !important;
        }
        [data-testid="stSidebar"] [role="radiogroup"] label[data-baseweb="radio"]:has(input:checked) p {
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }

        /* 4. BANNER ĐẦU TRANG */
        .page-banner {
            background-color: #E0F2FE;
            border: 1px solid #BAE6FD;
            border-left: 4px solid #0284C7;
            border-radius: 10px;
            padding: 14px 20px;
            margin-bottom: 20px;
        }
        .page-banner-title {
            color: #0F172A;
            font-size: 1.25rem;
            font-weight: 700;
            margin: 0 0 4px 0;
        }
        .page-banner-desc {
            color: #475569;
            font-size: 0.85rem;
            font-weight: 400;
            margin: 0;
        }

        /* 5. KHỐI CARD TRẮNG & EXPANDER */
        [data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF !important;
            border: 1px solid #E2E8F0 !important;
            border-radius: 10px !important;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03) !important;
            padding: 1rem !important;
            margin-bottom: 1rem !important;
        }

        [data-testid="stExpander"] {
            background-color: #FFFFFF !important;
            border: 1px solid #E2E8F0 !important;
            border-radius: 8px !important;
            margin-bottom: 10px !important;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.02) !important;
            overflow: hidden;
        }
        [data-testid="stExpander"] summary {
            background-color: #F8FAFC !important;
            padding: 10px 14px !important;
            border-radius: 8px !important;
            transition: background-color 0.15s ease;
        }
        [data-testid="stExpander"] summary:hover {
            background-color: #F1F5F9 !important;
        }
        [data-testid="stExpander"] summary p {
            color: #1E293B !important;
            font-weight: 600 !important;
            font-size: 0.9rem !important;
        }

        /* 6. Ô NHẬP LIỆU & SELECTBOX */
        .stSelectbox label p, .stTextInput label p, .stDateInput label p,
        .stNumberInput label p, .stFileUploader label p {
            color: #1E293B !important;
            font-weight: 600 !important;
            font-size: 0.875rem !important;
        }
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        div[data-baseweb="base-input"] {
            background-color: #FFFFFF !important;
            border-color: #E2E8F0 !important;
            border-radius: 8px !important;
            color: #0F172A !important;
        }
        div[data-baseweb="select"] > div:focus-within,
        div[data-baseweb="input"] > div:focus-within {
            border-color: #0284C7 !important;
            box-shadow: 0 0 0 1px #0284C7 !important;
        }

        /* Ức chế bàn phím ảo trên điện thoại khi bấm vào Selectbox */
        @media (max-width: 768px) {
            div[data-baseweb="select"] input {
                caret-color: transparent !important;
                pointer-events: none !important;
            }
        }

        /* 7. NÚT BẤM (BUTTONS) */
        .stButton > button[kind="primary"] {
            background-color: #0284C7 !important;
            color: #FFFFFF !important;
            font-weight: 600 !important;
            font-size: 0.875rem !important;
            border: 1px solid #0284C7 !important;
            border-radius: 8px !important;
            padding: 0.5rem 1.1rem !important;
            box-shadow: 0 1px 2px rgba(2, 132, 199, 0.15) !important;
            transition: background-color 0.15s ease !important;
        }
        .stButton > button[kind="primary"] p {
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }
        .stButton > button[kind="primary"]:hover {
            background-color: #0369A1 !important;
            border-color: #0369A1 !important;
        }

        .stButton > button[kind="secondary"], .stDownloadButton > button {
            background-color: #FFFFFF !important;
            color: #1E293B !important;
            font-weight: 600 !important;
            font-size: 0.875rem !important;
            border: 1px solid #CBD5E1 !important;
            border-radius: 8px !important;
            padding: 0.5rem 1rem !important;
            transition: all 0.15s ease !important;
        }
        .stButton > button[kind="secondary"] p, .stDownloadButton > button p {
            color: #1E293B !important;
            font-weight: 600 !important;
        }
        .stButton > button[kind="secondary"]:hover, .stDownloadButton > button:hover {
            background-color: #F8FAFC !important;
            border-color: #0284C7 !important;
            color: #0284C7 !important;
        }
        .stButton > button[kind="secondary"]:hover p, .stDownloadButton > button:hover p {
            color: #0284C7 !important;
        }

        /* 8. METRICS HIỂN THỊ SỐ LIỆU */
        [data-testid="stMetric"] {
            background-color: #F8FAFC !important;
            border: 1px solid #E2E8F0 !important;
            border-radius: 8px !important;
            padding: 10px 14px !important;
        }
        [data-testid="stMetricLabel"] p {
            color: #475569 !important;
            font-weight: 500 !important;
            font-size: 0.8rem !important;
        }
        [data-testid="stMetricValue"] {
            color: #0F172A !important;
            font-weight: 700 !important;
            font-size: 1.25rem !important;
            font-variant-numeric: tabular-nums !important;
        }

        /* 9. NHÃN TRẠNG THÁI (BADGES) & KHỐI KẾT QUẢ */
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 2px 10px;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 600;
            line-height: 1.4;
            white-space: nowrap;
        }
        .badge-pass {
            background-color: #DCFCE7;
            color: #15803D;
            border: 1px solid #86EFAC;
        }
        .badge-fail {
            background-color: #FEE2E2;
            color: #B91C1C;
            border: 1px solid #FCA5A5;
        }
        .badge-warn {
            background-color: #FEF3C7;
            color: #B45309;
            border: 1px solid #FDE68A;
        }

        .eval-result-box {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 10px 14px;
            margin-top: 10px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            font-size: 0.875rem;
        }
        .eval-result-pass {
            border-left: 4px solid #15803D;
        }
        .eval-result-fail {
            border-left: 4px solid #B91C1C;
        }

        /* 10. LỊCH SỬ & BẢNG DỮ LIỆU CHUẨN Y KHOA */
        .month-header-banner {
            background-color: #FFFFFF;
            color: #0F172A;
            padding: 8px 14px;
            border: 1px solid #E2E8F0;
            border-left: 4px solid #0284C7;
            border-radius: 8px;
            font-weight: 600;
            font-size: 0.95rem;
            margin-top: 18px;
            margin-bottom: 10px;
        }
        .history-meta-bar {
            background-color: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 10px 14px;
            margin-bottom: 12px;
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            align-items: center;
            color: #0F172A;
            font-size: 0.875rem;
        }
        .date-pill {
            background-color: #E0F2FE;
            color: #0369A1;
            border: 1px solid #BAE6FD;
            padding: 2px 10px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.85rem;
            font-variant-numeric: tabular-nums;
        }
        .lot-pill {
            background-color: #F1F5F9;
            color: #0F172A;
            border: 1px solid #CBD5E1;
            padding: 2px 10px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.85rem;
            font-variant-numeric: tabular-nums;
        }

        .lab-table-wrapper {
            max-height: 360px;
            overflow-y: auto;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            background-color: #FFFFFF;
            margin-bottom: 8px;
        }
        .lab-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.875rem;
            color: #0F172A;
        }
        .lab-table thead th {
            position: sticky;
            top: 0;
            background-color: #F8FAFC;
            color: #475569;
            font-weight: 600;
            padding: 10px 12px;
            border-bottom: 1px solid #E2E8F0;
            z-index: 1;
        }
        .lab-table tbody td {
            padding: 9px 12px;
            border-bottom: 1px solid #F1F5F9;
            font-variant-numeric: tabular-nums;
        }
        .lab-table tbody tr:nth-child(even) {
            background-color: #F8FAFC;
        }
        .lab-table tbody tr:hover {
            background-color: #E0F2FE;
        }
        .text-left { text-align: left !important; }
        .text-right { text-align: right !important; }
        .text-center { text-align: center !important; }
    </style>
    """, unsafe_allow_html=True)


def render_page_banner(title: str, desc: str):
    st.markdown(f"""
    <div class="page-banner">
        <div class="page-banner-title">{title}</div>
        <div class="page-banner-desc">{desc}</div>
    </div>
    """, unsafe_allow_html=True)


def format_status_badge_html(status_val: str) -> str:
    s = str(status_val).strip().upper()
    if s in ("PASS", "ĐẠT", "✅ ĐẠT"):
        return '<span class="status-badge badge-pass">✓ Đạt</span>'
    if s in ("FAIL", "TRƯỢT", "KHÔNG ĐẠT", "❌ TRƯỢT"):
        return '<span class="status-badge badge-fail">✕ Không đạt</span>'
    return f'<span class="status-badge badge-warn">! {status_val}</span>'


inject_custom_css()
init_db()

all_setups = load_all_setups()
hist_df_master = load_stage2_history()

# ==============================================================================
# THANH MENU DỌC BÊN TRÁI (SIDEBAR)
# ==============================================================================
MENU_ITEMS = [
    "1. Đánh giá Mẫu thực tế (Stage 2)",
    "2. Lịch sử Đánh giá",
    "3. Tổng quan LOT (Dashboard)",
    "4. Cấu hình Thông số (Stage 1)",
    "5. Danh mục XN đã lưu"
]

MENU_ICONS = {
    "1. Đánh giá Mẫu thực tế (Stage 2)": ":material/science:  1. Đánh giá Mẫu thực tế (Stage 2)",
    "2. Lịch sử Đánh giá": ":material/history:  2. Lịch sử Đánh giá",
    "3. Tổng quan LOT (Dashboard)": ":material/dashboard:  3. Tổng quan LOT (Dashboard)",
    "4. Cấu hình Thông số (Stage 1)": ":material/tune:  4. Cấu hình Thông số (Stage 1)",
    "5. Danh mục XN đã lưu": ":material/folder_managed:  5. Danh mục XN đã lưu"
}

with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="sidebar-brand-title">CLSI EP26-Ed2</div>
        <div class="sidebar-brand-sub">Hệ thống đánh giá Lot-to-Lot</div>
    </div>
    <div class="sidebar-section-label">Danh mục chức năng</div>
    """, unsafe_allow_html=True)

    page = st.radio(
        "Điều hướng",
        options=MENU_ITEMS,
        format_func=lambda x: MENU_ICONS.get(x, x),
        label_visibility="collapsed"
    )


def convert_docx_to_pdf_bytes(docx_bytes):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "report.docx")
        pdf_path = os.path.join(tmpdir, "report.pdf")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        try:
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', docx_path, '--outdir', tmpdir],
                           check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except FileNotFoundError:
            try:
                subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', docx_path, '--outdir', tmpdir],
                               check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            except Exception:
                return None
        except Exception:
            return None
        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                return f.read()
    return None


# ==============================================================================
# TRANG 1: GIAI ĐOẠN 2 - ĐÁNH GIÁ MẪU HẰNG NGÀY
# ==============================================================================
if page == "1. Đánh giá Mẫu thực tế (Stage 2)":
    render_page_banner(
        "Đánh giá Mẫu thực tế (Stage 2)",
        "Nhập kết quả chạy mẫu song song trên lô hóa chất cũ và mới để kiểm tra độ tương đồng theo tiêu chuẩn CLSI EP26-Ed2."
    )

    active_setups = all_setups[all_setups["is_selected"] == 1] if not all_setups.empty else pd.DataFrame()
    if active_setups.empty:
        st.warning("Chưa có cấu hình xét nghiệm nào. Vui lòng chọn mục **4. Cấu hình Thông số (Stage 1)** ở thanh menu bên trái để thiết lập thông số ban đầu!", icon=":material/warning:")
    else:
        with st.container(border=True):
            st.markdown("#### :material/assignment: Thông tin đợt đánh giá")
            c1, c2, c3 = st.columns([1.5, 1.5, 1])
            with c1:
                eval_test = st.selectbox("Chọn Xét nghiệm:", active_setups["test_name"].unique().tolist(),
                                         key="s2_test")
            with c2:
                eval_machine = st.selectbox("Dòng máy:", active_setups[active_setups["test_name"] == eval_test][
                    "machine"].unique().tolist(), key="s2_mach")
                eval_machine_id = ""
                mach_df = active_setups[
                    (active_setups["test_name"] == eval_test) & (active_setups["machine"] == eval_machine)]
                if not mach_df.empty and "machine_id" in mach_df.columns and pd.notna(mach_df["machine_id"].iloc[0]):
                    eval_machine_id = str(mach_df["machine_id"].iloc[0])
            with c3:
                eval_date = st.date_input("Ngày thực hiện:", value=date.today(), key="s2_date", format="DD/MM/YYYY")

            c4, c5, c6 = st.columns([2, 1, 1])
            with c4:
                reagent_name = st.text_input("Loại hóa chất thay thế (Huyết học/HbA1c):",
                                             placeholder="VD: CELL-DYN Diluent/Sheath... (Bỏ trống nếu là Sinh hóa)",
                                             key="s2_reagent").strip()
            with c5:
                lot_old = st.text_input("Số LOT cũ:", value="82273UD00", key="s2_lot_old").strip()
            with c6:
                lot_new = st.text_input("Số LOT mới:", value="86280UD00", key="s2_lot_new").strip()

        test_levels_df = mach_df.sort_values("target_conc").reset_index(drop=True)
        unit_s2 = str(test_levels_df["unit"].iloc[0]) if not test_levels_df.empty else "mmol/L"

        st.markdown("### :material/biotech: Nhập kết quả mẫu")
        flat_records_to_save = []
        summary_table_rows = []

        for idx, lv_row in test_levels_df.iterrows():
            level_num = idx + 1
            target_val = float(lv_row["target_conc"])
            rl_pct_db = float(lv_row["rejection_limit"])
            k_db = float(lv_row["k_factor"])
            n_req = max(3, int(lv_row["n_samples"]))

            with st.container(border=True):
                head_col1, head_col2 = st.columns([2.2, 1.8])
                with head_col1:
                    st.markdown(f"#### :material/target: Mức nồng độ: `{target_val:.3f} {unit_s2}`")
                    st.markdown(
                        f"**Yêu cầu:** `{int(lv_row['n_samples'])} mẫu` &nbsp;|&nbsp; **Giới hạn RL:** `{rl_pct_db:.3f}%`")

                default_samples_lv = pd.DataFrame(
                    [{"SID (Mã mẫu)": "", "Kết quả LOT cũ": 0.00, "Kết quả LOT mới": 0.00} for _ in range(n_req)])
                in_col, out_col = st.columns([2.2, 1.8])
                with in_col:
                    edited_samples = st.data_editor(
                        default_samples_lv, num_rows="dynamic", use_container_width=True,
                        column_config={
                            "SID (Mã mẫu)": st.column_config.TextColumn("SID mẫu bệnh nhân"),
                            "Kết quả LOT cũ": st.column_config.NumberColumn(f"LOT cũ ({lot_old})", format="%.3f"),
                            "Kết quả LOT mới": st.column_config.NumberColumn(f"LOT mới ({lot_new})", format="%.3f"),
                        }, key=f"s2_{eval_test}_{level_num}"
                    )
                eval_res = evaluate_stage2_level(samples_df=edited_samples, rl_pct=rl_pct_db)

                with out_col:
                    if not eval_res["has_data"]:
                        st.info("Nhập ít nhất 1 cặp kết quả (> 0) để xem đánh giá.", icon=":material/info:")
                    else:
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Trung bình d", f"{eval_res['mean_d']:.4f}")
                        m2.metric("Trung bình độ chệch (%)", f"{eval_res['mean_d_pct']:.3f}%")
                        m3.metric("Rejection Limit", f"{eval_res['rl_pct']:.3f}%")

                        if eval_res["result"] == "PASS":
                            st.markdown(f"""
                            <div class="eval-result-box eval-result-pass">
                                <div><strong>Kết luận:</strong> <span class="tabular-nums">{eval_res['mean_d_pct']:.3f}% &lt; {eval_res['rl_pct']:.3f}%</span></div>
                                {format_status_badge_html('PASS')}
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.markdown(f"""
                            <div class="eval-result-box eval-result-fail">
                                <div><strong>Kết luận:</strong> <span class="tabular-nums">{eval_res['mean_d_pct']:.3f}% &ge; {eval_res['rl_pct']:.3f}%</span></div>
                                {format_status_badge_html('FAIL')}
                            </div>
                            """, unsafe_allow_html=True)

                        for _, d_row in eval_res["details"].iterrows():
                            flat_records_to_save.append({
                                "eval_date": str(eval_date), "test_name": eval_test, "machine": eval_machine,
                                "machine_id": eval_machine_id,
                                "unit": unit_s2, "lot_old": lot_old, "lot_new": lot_new,
                                "level_name": f"Level {level_num}", "target_conc": target_val,
                                "sid": d_row["SID"], "val_old": d_row["LOT_old_val"], "val_new": d_row["LOT_new_val"],
                                "diff_d": d_row["d"],
                                "mean_d": eval_res["mean_d"], "mean_d_pct": eval_res["mean_d_pct"],
                                "rl_pct": eval_res["rl_pct"], "k_factor": k_db,
                                "status": eval_res["result"], "reagent_name": reagent_name
                            })
                        summary_table_rows.append({"Kết quả": eval_res["result"]})

        if summary_table_rows:
            st.write("")
            if st.button(":material/save: Lưu kết quả vào lịch sử", type="primary", use_container_width=True):
                save_stage2_evaluation(flat_records_to_save)
                st.success("Đã lưu biên bản đánh giá! Vui lòng chọn mục **2. Lịch sử Đánh giá** ở thanh menu bên trái để xuất file Word/PDF.", icon=":material/check_circle:")
                time.sleep(0.5)
                st.rerun()


# ==============================================================================
# TRANG 2: LỊCH SỬ ĐÁNH GIÁ & XÁC NHẬN XÓA & THÙNG RÁC
# ==============================================================================
elif page == "2. Lịch sử Đánh giá":
    render_page_banner(
        "Lịch sử Đánh giá Mẫu thực tế",
        "Tra cứu hồ sơ đánh giá Lot-to-Lot theo từng tháng, xuất biên bản báo cáo Word/PDF và khôi phục các bản ghi đã xóa."
    )

    hist_df = hist_df_master.copy()

    if hist_df.empty:
        st.info("Chưa có đợt đánh giá nào được lưu.", icon=":material/info:")
    else:
        hist_df['eval_date'] = pd.to_datetime(hist_df['eval_date'])
        hist_df = hist_df.sort_values(by=['eval_date', 'created_at'], ascending=[False, False])
        hist_df['YearMonth'] = hist_df['eval_date'].dt.strftime('%m / %Y')
        year_months = hist_df['YearMonth'].unique().tolist()

        for ym in year_months:
            st.markdown(f'<div class="month-header-banner">Tháng {ym}</div>', unsafe_allow_html=True)
            ym_df = hist_df[hist_df['YearMonth'] == ym]
            grouped_hist = ym_df.groupby('created_at', sort=False)

            for created_at, group in grouped_hist:
                first_row = group.iloc[0]
                t_name = first_row['test_name']
                mach = first_row['machine']
                mach_id = first_row.get('machine_id', '')
                l_old = first_row['lot_old']
                l_new = first_row['lot_new']
                reagent = str(first_row.get('reagent_name', '')).strip()
                if reagent.lower() == 'nan':
                    reagent = ''
                e_date_str = first_row['eval_date'].strftime('%Y-%m-%d')
                display_date_str = first_row['eval_date'].strftime('%d/%m/%Y')
                unit = first_row['unit']

                overall_pass = all(str(r['status']).upper() == 'PASS' for _, r in group.iterrows())
                status_text = "✓ Đạt" if overall_pass else "✕ Không đạt"
                reagent_str = f"  •  Hóa chất: {reagent}" if reagent else ""
                expander_label = f":material/event: [{display_date_str}]  —  {t_name} ({mach}){reagent_str}  •  LOT: {l_old} ➔ {l_new}  [{status_text}]"

                with st.expander(expander_label, expanded=False):
                    reagent_meta = f'<div><strong>Hóa chất:</strong> {reagent}</div>' if reagent else ''
                    overall_badge = format_status_badge_html("PASS" if overall_pass else "FAIL")
                    st.markdown(f"""
                    <div class="history-meta-bar">
                        <div><strong>Ngày thực hiện:</strong> <span class="date-pill">{display_date_str}</span></div>
                        <div><strong>Xét nghiệm:</strong> {t_name} ({unit})</div>
                        <div><strong>Thiết bị:</strong> {mach} {f'({mach_id})' if mach_id else ''}</div>
                        {reagent_meta}
                        <div><strong>Chuyển đổi LOT:</strong> <span class="lot-pill">{l_old} ➔ {l_new}</span></div>
                        <div><strong>Trạng thái:</strong> {overall_badge}</div>
                    </div>
                    """, unsafe_allow_html=True)

                    # Bảng HTML chuẩn y khoa: sticky header, zebra rows, số căn phải, badge trạng thái
                    rows_html = []
                    for _, r in group.iterrows():
                        badge_cell = format_status_badge_html(r['status'])
                        rows_html.append(f"""
                        <tr>
                            <td class="text-left">{r['level_name']}</td>
                            <td class="text-right">{float(r['target_conc']):.3f}</td>
                            <td class="text-left">{r['sid']}</td>
                            <td class="text-right">{float(r['val_old']):.3f}</td>
                            <td class="text-right">{float(r['val_new']):.3f}</td>
                            <td class="text-right">{float(r['diff_d']):.3f}</td>
                            <td class="text-right">{float(r['mean_d']):.3f}</td>
                            <td class="text-right">{float(r['mean_d_pct']):.3f}</td>
                            <td class="text-right">{float(r['rl_pct']):.3f}</td>
                            <td class="text-center">{badge_cell}</td>
                        </tr>
                        """)

                    table_html = f"""
                    <div class="lab-table-wrapper">
                        <table class="lab-table">
                            <thead>
                                <tr>
                                    <th class="text-left">Mức (Level)</th>
                                    <th class="text-right">Nồng độ đích</th>
                                    <th class="text-left">Mã mẫu (SID)</th>
                                    <th class="text-right">KQ LOT cũ</th>
                                    <th class="text-right">KQ LOT mới</th>
                                    <th class="text-right">Độ chệch |d|</th>
                                    <th class="text-right">TB |d|</th>
                                    <th class="text-right">TB |d| (%)</th>
                                    <th class="text-right">Giới hạn RL (%)</th>
                                    <th class="text-center">Kết quả</th>
                                </tr>
                            </thead>
                            <tbody>
                                {"".join(rows_html)}
                            </tbody>
                        </table>
                    </div>
                    """
                    st.markdown(table_html, unsafe_allow_html=True)

                    st.write("")
                    col_word, col_pdf, col_del = st.columns([1, 1, 1])

                    with col_word:
                        setup_df = all_setups[(all_setups['test_name'] == t_name) & (all_setups['machine'] == mach)]
                        docx_file = generate_report_docx(e_date_str, t_name, unit, mach, mach_id, reagent, l_old, l_new,
                                                         setup_df, group)
                        st.download_button(
                            label=":material/description: Xuất file Word",
                            data=docx_file,
                            file_name=f"LotToLot_{t_name}_{l_old}_{l_new}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_word_{created_at}",
                            use_container_width=True
                        )

                    with col_pdf:
                        if f"pdf_ready_{created_at}" not in st.session_state:
                            if st.button(":material/picture_as_pdf: Tạo file PDF", key=f"create_pdf_{created_at}", use_container_width=True):
                                with st.spinner("Đang vẽ PDF..."):
                                    pdf_bytes = convert_docx_to_pdf_bytes(docx_file.getvalue())
                                    if pdf_bytes:
                                        st.session_state[f"pdf_bytes_{created_at}"] = pdf_bytes
                                        st.session_state[f"pdf_ready_{created_at}"] = True
                                        st.rerun()
                                    else:
                                        st.error("Chưa cài LibreOffice trên Cloud!", icon=":material/error:")
                        else:
                            st.download_button(
                                label=":material/download: Tải PDF xuống",
                                data=st.session_state[f"pdf_bytes_{created_at}"],
                                file_name=f"LotToLot_{t_name}_{l_old}_{l_new}.pdf",
                                mime="application/pdf",
                                key=f"dl_pdf_{created_at}",
                                use_container_width=True
                            )

                    with col_del:
                        # TÍNH NĂNG XÁC NHẬN TRƯỚC KHI XÓA
                        if f"confirm_del_{created_at}" not in st.session_state:
                            st.session_state[f"confirm_del_{created_at}"] = False

                        if not st.session_state[f"confirm_del_{created_at}"]:
                            if st.button(":material/delete: Xóa đợt này", key=f"del_hist_{created_at}", use_container_width=True):
                                st.session_state[f"confirm_del_{created_at}"] = True
                                st.rerun()
                        else:
                            st.warning("Bạn có chắc chắn muốn xóa? (Dữ liệu sẽ vào Thùng rác)", icon=":material/warning:")
                            c_yes, c_no = st.columns(2)
                            with c_yes:
                                if st.button(":material/check: Xác nhận", key=f"yes_del_{created_at}", type="primary", use_container_width=True):
                                    delete_stage2_evaluation(created_at)
                                    st.session_state[f"confirm_del_{created_at}"] = False
                                    st.rerun()
                            with c_no:
                                if st.button(":material/close: Hủy", key=f"no_del_{created_at}", use_container_width=True):
                                    st.session_state[f"confirm_del_{created_at}"] = False
                                    st.rerun()

    # TÍNH NĂNG THÙNG RÁC - HIỂN THỊ DƯỚI CÙNG CỦA TRANG LỊCH SỬ
    st.write("")
    with st.expander(":material/restore_from_trash: Thùng rác (Khôi phục hoặc xóa vĩnh viễn)", expanded=False):
        trash_df = load_trash_history()
        if trash_df.empty:
            st.info("Thùng rác hiện đang trống.", icon=":material/info:")
        else:
            grouped_trash = trash_df.groupby('created_at', sort=False)
            for c_at, group in grouped_trash:
                first = group.iloc[0]
                disp_date = pd.to_datetime(first['eval_date']).strftime('%d/%m/%Y')
                t_name = first['test_name']
                reagent = str(first.get('reagent_name', '')).strip()
                reagent_str = f" ({reagent})" if reagent and reagent.lower() != 'nan' else ""

                c_trash1, c_restore, c_delete = st.columns([3, 1, 1])
                with c_trash1:
                    st.markdown(f"**{disp_date}** &nbsp;|&nbsp; **{t_name}**{reagent_str} &nbsp;|&nbsp; LOT: `{first['lot_old']} ➔ {first['lot_new']}`")

                with c_restore:
                    if st.button(":material/restore: Khôi phục", key=f"restore_{c_at}", use_container_width=True):
                        restore_stage2_evaluation(c_at)
                        st.success("Đã khôi phục!", icon=":material/check_circle:")
                        time.sleep(0.5)
                        st.rerun()

                with c_delete:
                    if st.button(":material/delete_forever: Xóa hẳn", key=f"perm_del_{c_at}", use_container_width=True):
                        permanently_delete_stage2_evaluation(c_at)
                        st.error("Đã xóa vĩnh viễn!", icon=":material/delete:")
                        time.sleep(0.5)
                        st.rerun()


# ==============================================================================
# TRANG 3: TỔNG QUAN LOT (DASHBOARD)
# ==============================================================================
elif page == "3. Tổng quan LOT (Dashboard)":
    render_page_banner(
        "Tổng quan LOT (Dashboard)",
        "Bảng điều khiển hiện trạng hóa chất và lịch sử thay đổi số LOT trên từng thiết bị phân tích."
    )

    df_hist_dash = hist_df_master.copy()
    if df_hist_dash.empty:
        st.info("Chưa có dữ liệu đánh giá.", icon=":material/info:")
    else:
        df_hist_dash['reagent_name'] = df_hist_dash['reagent_name'].fillna('').astype(str).str.strip()
        df_hist_dash.loc[df_hist_dash['reagent_name'].str.lower() == 'nan', 'reagent_name'] = ''
        df_hist_dash['eval_date_obj'] = pd.to_datetime(df_hist_dash['eval_date'])

        df_unique_evals = df_hist_dash.drop_duplicates(subset=['created_at']).copy()
        df_unique_evals = df_unique_evals.sort_values(by=['eval_date_obj', 'created_at'], ascending=[False, False])

        df_sh = df_unique_evals[df_unique_evals['reagent_name'] == '']
        df_hh = df_unique_evals[df_unique_evals['reagent_name'] != '']

        col_sh, col_hh = st.columns(2)

        def render_dashboard_section(df_section, title, icon, group_by_col):
            st.markdown(f"### {icon} {title}")
            if df_section.empty:
                st.info("Chưa có dữ liệu.", icon=":material/info:")
                return
            for mach in sorted(df_section['machine'].unique()):
                with st.container(border=True):
                    st.markdown(f"#### :material/dns: Máy: **{mach}**")
                    for item_name, group in df_section[df_section['machine'] == mach].groupby(group_by_col, sort=False):
                        latest_eval = group.iloc[0]
                        latest_status_str = "✓ Đạt" if str(latest_eval['status']).upper() == "PASS" else "✕ Không đạt"
                        expander_title = (
                            f":material/vaccines: {item_name}  |  LOT đang dùng: {latest_eval['lot_new']}  "
                            f"(Cập nhật: {latest_eval['eval_date_obj'].strftime('%d/%m/%Y')})  [{latest_status_str}]"
                        )
                        with st.expander(expander_title):
                            dash_rows_html = []
                            for _, r in group.iterrows():
                                date_fmt = r['eval_date_obj'].strftime('%d/%m/%Y')
                                badge_html = format_status_badge_html(r['status'])
                                dash_rows_html.append(f"""
                                <tr>
                                    <td class="text-left">{date_fmt}</td>
                                    <td class="text-right">{r['lot_old']}</td>
                                    <td class="text-right">{r['lot_new']}</td>
                                    <td class="text-center">{badge_html}</td>
                                </tr>
                                """)
                            dash_table_html = f"""
                            <div class="lab-table-wrapper">
                                <table class="lab-table">
                                    <thead>
                                        <tr>
                                            <th class="text-left">Ngày thay LOT</th>
                                            <th class="text-right">LOT cũ</th>
                                            <th class="text-right">LOT mới</th>
                                            <th class="text-center">Kết quả</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {"".join(dash_rows_html)}
                                    </tbody>
                                </table>
                            </div>
                            """
                            st.markdown(dash_table_html, unsafe_allow_html=True)

        with col_sh:
            render_dashboard_section(df_sh, "Sinh hóa - Miễn dịch", ":material/water_drop:", group_by_col="test_name")
        with col_hh:
            render_dashboard_section(df_hh, "Huyết học - HbA1c", ":material/bloodtype:", group_by_col="reagent_name")


# ==============================================================================
# TRANG 4: GIAI ĐOẠN 1 - SETUP THÔNG SỐ
# ==============================================================================
elif page == "4. Cấu hình Thông số (Stage 1)":
    render_page_banner(
        "Cấu hình Thông số (Stage 1)",
        "Thiết lập nồng độ đích, TEa%, Swrl% và Sr% để hệ thống tra bảng tính số lượng mẫu (n), lực thống kê (Power) và giới hạn từ chối (RL)."
    )

    all_setups_for_dropdown = all_setups
    with st.container(border=True):
        st.markdown("#### :material/tune: Thông tin xét nghiệm & thiết bị")
        col1, col2, col3, col4, col5 = st.columns([1.5, 1.5, 1.2, 1, 1.5])
        with col1:
            db_tests = all_setups_for_dropdown[
                "test_name"].unique().tolist() if not all_setups_for_dropdown.empty else []
            test_list = list(dict.fromkeys(
                db_tests + ["CHOLESTEROL", "GLUCOSE", "TSH", "GGT", "AST", "CREATININ", "ALT", "TRIGLYXERIDE", "FT4",
                            "ACID URIC"])) + ["-- Nhập tên mới --"]
            selected_test = st.selectbox("Tên xét nghiệm:", test_list)
            test_name = st.text_input("Nhập tên XN mới:",
                                      "HBA1C").strip().upper() if selected_test == "-- Nhập tên mới --" else selected_test
        with col2:
            db_machs = all_setups_for_dropdown["machine"].unique().tolist() if not all_setups_for_dropdown.empty else []
            machine_list = list(dict.fromkeys(db_machs + ["Abbott (Alinity)", "Roche (Cobas)"])) + [
                "-- Nhập máy khác --"]
            selected_machine = st.selectbox("Dòng máy:", machine_list)
            machine_name = st.text_input("Nhập tên máy mới:",
                                         "Sysmex").strip() if selected_machine == "-- Nhập máy khác --" else selected_machine

        saved_df = load_specific_setup(test_name, machine_name)
        existing_mach_id = str(
            saved_df["machine_id"].iloc[0]) if not saved_df.empty and "machine_id" in saved_df.columns and pd.notna(
            saved_df["machine_id"].iloc[0]) and str(saved_df["machine_id"].iloc[0]) != 'nan' else ""
        existing_unit = str(saved_df["unit"].iloc[0]) if not saved_df.empty and "unit" in saved_df.columns and pd.notna(
            saved_df["unit"].iloc[0]) and str(saved_df["unit"].iloc[0]) != 'nan' else "mmol/L"

        with col3:
            machine_id = st.text_input("Mã thiết bị:", value=existing_mach_id).strip()
        with col4:
            unit_name = st.text_input("Đơn vị đo:", value=existing_unit)
        with col5:
            global_target_power = st.selectbox(
                "Power tối thiểu:",
                options=[0.0, 0.50, 0.60, 0.70, 0.80, 0.85, 0.90, 0.95],
                format_func=lambda x: "Hiển thị tất cả" if x == 0.0 else f">= {x * 100}%",
                index=0
            )

    if not saved_df.empty:
        st.info("Đã lấy dữ liệu cũ từ Database. Bạn có thể chỉnh sửa và bấm Lưu để cập nhật.", icon=":material/sync:")
        init_input_df = pd.DataFrame({
            "Chọn chạy (Vàng)": saved_df["is_selected"].astype(bool),
            "Target concentration": saved_df["target_conc"].astype(float),
            "TEa%": saved_df["tea_pct"].astype(float),
            "Swrl = CVwrl %": saved_df["swrl_pct"].astype(float),
            "Sr = CVr %": saved_df["sr_pct"].astype(float)
        })
    else:
        init_input_df = DEFAULT_SAMPLES.get((test_name, machine_name), pd.DataFrame([
            {
                "Chọn chạy (Vàng)": True,
                "Target concentration": 6.37,
                "TEa%": 12.9,
                "Swrl = CVwrl %": 0.5,
                "Sr = CVr %": 0.3
            }
        ])).copy()

    st.markdown("### :material/edit_note: Nhập thông số và chọn mức nồng độ")
    edited_input = st.data_editor(
        init_input_df, num_rows="dynamic", use_container_width=True,
        column_config={
            "Chọn chạy (Vàng)": st.column_config.CheckboxColumn("Chọn chạy"),
            "Target concentration": st.column_config.NumberColumn("Target concentration", format="%.3f"),
            "TEa%": st.column_config.NumberColumn("TEa%", format="%.3f"),
            "Swrl = CVwrl %": st.column_config.NumberColumn("Swrl = CVwrl %", format="%.3f"),
            "Sr = CVr %": st.column_config.NumberColumn("Sr = CVr %", format="%.3f"),
        }, key=f"editor_{test_name}_{machine_name}"
    )

    selected_input = edited_input[edited_input["Chọn chạy (Vàng)"] == True]
    num_selected_levels = max(1, min(3, len(selected_input) if len(selected_input) > 0 else 1))

    if not selected_input.empty:
        chosen_final_rows = []
        for idx, row in selected_input.iterrows():
            if pd.isna(row["Target concentration"]) or pd.isna(row["TEa%"]):
                continue
            target_c = float(row["Target concentration"])
            candidates = get_candidate_options_for_row(
                tea_pct=float(row["TEa%"]),
                swrl_pct=float(row["Swrl = CVwrl %"]),
                sr_pct=float(row["Sr = CVr %"]),
                num_levels=num_selected_levels,
                desired_power=global_target_power,
                sort_mode="min_n_max_power"
            )

            with st.container(border=True):
                st.markdown(f"#### :material/target: Mức nồng độ: `{target_c:.3f} {unit_name}`")
                if len(candidates) > 1:
                    option_labels = [
                        f"Phương án {i + 1} (Tra Bảng A{num_selected_levels}): {c['n_samples']} mẫu | Power = {c['power_achieve']:.3f} | CD/Swrl = {c['cd_swrl']} | Sr/Swrl = {c['sr_swrl']} | K = {c['k_factor']} | RL = {c['rejection_limit']:.2f}%"
                        for i, c in enumerate(candidates)
                    ]
                    selected_label = st.selectbox(
                        "Chọn phương án chốt:",
                        options=option_labels,
                        index=0,
                        key=f"pick_{test_name}_{machine_name}_{idx}",
                        label_visibility="collapsed"
                    )
                    picked_cand = candidates[option_labels.index(selected_label)]
                else:
                    picked_cand = candidates[0]
                    st.info(
                        f"Tra Bảng A{num_selected_levels} | Cần {picked_cand['n_samples']} mẫu | Power = {picked_cand['power_achieve']:.3f} | CD/Swrl = {picked_cand['cd_swrl']} | Sr/Swrl = {picked_cand['sr_swrl']} | K = {picked_cand['k_factor']} | RL = {picked_cand['rejection_limit']:.2f}%",
                        icon=":material/lightbulb:"
                    )

                row_to_save = row.copy().to_dict()
                row_to_save.update({
                    "unit": unit_name,
                    "power_achieve": picked_cand["power_achieve"],
                    "cd_swrl": picked_cand["cd_swrl"],
                    "sr_swrl": picked_cand["sr_swrl"],
                    "k_factor": picked_cand["k_factor"],
                    "rejection_limit": picked_cand["rejection_limit"],
                    "n_samples": picked_cand["n_samples"],
                    "Ghi chú": picked_cand.get(
                        "note",
                        f"Đạt chuẩn (Mốc tra bảng: CD/Swrl={picked_cand['cd_swrl']}, Sr/Swrl={picked_cand['sr_swrl']})"
                    )
                })
                chosen_final_rows.append(row_to_save)

        if st.button(":material/save: Lưu cấu hình vào Database", type="primary", use_container_width=True):
            save_test_setup(test_name, machine_name, machine_id, pd.DataFrame(chosen_final_rows))
            st.success(f"Đã lưu thành công thông số và Mã thiết bị cho {test_name} - {machine_name}!", icon=":material/check_circle:")
            time.sleep(0.5)
            st.rerun()


# ==============================================================================
# TRANG 5: DANH MỤC CẤU HÌNH ĐÃ LƯU
# ==============================================================================
elif page == "5. Danh mục XN đã lưu":
    render_page_banner(
        "Danh mục Xét nghiệm đã lưu",
        "Quản lý toàn bộ thông số cấu hình Stage 1 hiện có trong cơ sở dữ liệu và hỗ trợ khôi phục dữ liệu từ file CSV/Excel."
    )

    with st.expander(":material/upload_file: Khôi phục / Import dữ liệu từ file (.csv, .xlsx)", expanded=False):
        uploaded_file = st.file_uploader("Kéo thả file dữ liệu đã xuất từ hệ thống vào đây", type=["csv", "xlsx"])
        if uploaded_file is not None:
            try:
                df_import = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(
                    uploaded_file)
                if "test_name" in df_import.columns and "target_conc" in df_import.columns:
                    st.info(f"Phát hiện **{len(df_import)}** dòng cấu hình hợp lệ.", icon=":material/info:")
                    if st.button(":material/check_circle: Bắt đầu Import Dữ liệu", type="primary", use_container_width=True):
                        import_stage1_from_dataframe(df_import)
                        st.success("Khôi phục thành công!", icon=":material/check_circle:")
                        time.sleep(1)
                        st.rerun()
                else:
                    st.error("File không đúng định dạng!", icon=":material/error:")
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}", icon=":material/error:")

    if not all_setups.empty:
        with st.container(border=True):
            st.markdown("#### :material/table_chart: Danh sách cấu hình Stage 1 hiện có")
            st.dataframe(all_setups, use_container_width=True, hide_index=True)
            c_d1, c_d2, c_d3 = st.columns([2, 2, 1])
            with c_d1:
                del_test = st.selectbox("Chọn XN cần xóa:", all_setups["test_name"].unique(), key="d_t1")
            with c_d2:
                del_mach = st.selectbox("Chọn Máy cần xóa:", all_setups["machine"].unique(), key="d_m1")
            with c_d3:
                st.write("")
                st.write("")
                if st.button(":material/delete: Xóa cấu hình", use_container_width=True):
                    delete_specific_setup(del_test, del_mach)
                    st.rerun()
    else:
        st.info("Chưa có cấu hình xét nghiệm nào được lưu.", icon=":material/info:")