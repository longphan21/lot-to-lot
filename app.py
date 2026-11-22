import subprocess
import tempfile
import os
import re
import json
import base64
import urllib.request
from datetime import date, datetime
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

# 1. Dán link thư mục Google Drive của bạn vào đây:
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1Bsrb3rxJNYFdAIFgaPm8nMSeEe5nk0km?usp=sharing"

# 2. Dán link Web App URL vừa lấy ở Bước 1 vào đây:
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbycG9IcKnNAcnFxyFYQGeglISyKNya3dS4VX251seSpMm8gWyGKBzr2xE_hm9ADp-nwVQ/exec"


# ==============================================================================
# TỰ ĐỘNG NHẬN DIỆN THIẾT BỊ (DESKTOP / MOBILE) NGAY KHI MỞ TRANG
# ==============================================================================
MOBILE_UA_PATTERN = re.compile(
    r"android|iphone|ipod|ipad|mobile|blackberry|iemobile|opera mini|windows phone",
    re.IGNORECASE,
)


def _detect_mobile_from_request() -> bool:
    """Nhận diện điện thoại từ User-Agent hoặc tham số ?view=mobile / ?view=desktop."""
    try:
        qp = st.query_params
        v_param = str(qp.get("view", "")).strip().lower()
        if v_param == "mobile":
            return True
        if v_param == "desktop":
            return False
    except Exception:
        pass

    try:
        headers = st.context.headers
        ua = ""
        if headers:
            ua = headers.get("User-Agent") or headers.get("user-agent") or ""
        return bool(MOBILE_UA_PATTERN.search(str(ua)))
    except Exception:
        return False


_INITIAL_IS_MOBILE = _detect_mobile_from_request()

st.set_page_config(
    page_title="CLSI EP26 - Lot-to-Lot",
    layout="wide",
    initial_sidebar_state="collapsed" if _INITIAL_IS_MOBILE else "expanded"
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

_dialog_decorator = getattr(st, "dialog", getattr(st, "experimental_dialog", None))


def render_html(html_str: str):
    """Xóa khoảng trắng đầu dòng để Streamlit luôn hiển thị đúng HTML."""
    clean_html = "\n".join(line.lstrip() for line in html_str.splitlines() if line.strip())
    st.markdown(clean_html, unsafe_allow_html=True)


def clean_lot_str(val) -> str:
    """Làm sạch mã LOT/SID nếu dữ liệu cũ trên Sheet bị dính đuôi .0 do ép kiểu số."""
    s = str(val).strip()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        return s[:-2]
    return s



def build_report_filename(test_name: str, machine: str, reagent: str, display_date_str: str) -> str:
    """Đổi tên file: Tên XN - Tên máy - Loại hóa chất (nếu có) - Ngày đánh giá.docx"""
    def _clean(s: str) -> str:
        return re.sub(r'[\\/*?:"<>|]+', '-', str(s).strip())

    t_clean = _clean(test_name)
    m_clean = _clean(machine)
    d_clean = _clean(display_date_str)
    r_clean = _clean(reagent) if reagent and str(reagent).strip().lower() != "nan" else ""

    if r_clean:
        return f"{t_clean} - {m_clean} - {r_clean} - {d_clean}.docx"
    return f"{t_clean} - {m_clean} - {d_clean}.docx"


def upload_word_to_drive(docx_bytes: bytes, filename: str):
    """Tự động đẩy file Word lên thư mục Google Drive cá nhân qua Apps Script."""
    if "DÁN_LINK" in DRIVE_FOLDER_URL or "DÁN_LINK" in APPS_SCRIPT_URL:
        queue_popup(
            "Chưa cấu hình Google Drive",
            "Hãy dán link thư mục Drive và Apps Script vào đầu file <code>app.py</code> để sử dụng tính năng lưu lên Drive.",
            "warning"
        )
        return
    try:
        payload = json.dumps({
            "folderUrl": DRIVE_FOLDER_URL,
            "fileName": filename,
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "fileBase64": base64.b64encode(docx_bytes).decode("utf-8")
        }).encode("utf-8")
        req = urllib.request.Request(APPS_SCRIPT_URL, data=payload, headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=25)
        queue_popup(
            "Đã lưu lên Google Drive!",
            f"File <b>{filename}</b> đã được tải lên thư mục Google Drive của bạn thành công.",
            "success"
        )
    except Exception as e:
        queue_popup(
            "Lỗi lưu lên Google Drive",
            f"Không thể đẩy file <b>{filename}</b> lên Google Drive: {e}",
            "error"
        )


@st.cache_data(ttl=3600, show_spinner=False)
def get_cached_report_docx_bytes(eval_date_str, test_name, unit, machine, machine_id, reagent_name, lot_old, lot_new, setup_df, group_df) -> bytes:
    """Lưu bộ nhớ đệm (cache) cho file Word đã tạo để mở Tab 2 (Lịch sử) nhanh tức thì."""
    docx_io = generate_report_docx(
        eval_date_str, test_name, unit, machine, machine_id,
        reagent_name, lot_old, lot_new, setup_df, group_df
    )
    return docx_io.getvalue()


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
# STATE PERSISTENCE & STEP-BY-STEP LOCKING HELPERS
# ==============================================================================
def persistent_selectbox(
    container,
    label: str,
    options: list,
    persist_key: str,
    default_value=None,
    format_func=None,
    disabled: bool = False,
    on_value_change=None,
    label_visibility: str = "visible",
):
    """Selectbox có State Persistence — không bao giờ reset về dòng đầu khi Save hay đổi tab."""
    if not options:
        return None

    saved_val = st.session_state.get(persist_key)
    if saved_val not in options:
        if default_value is not None and default_value in options:
            saved_val = default_value
        else:
            saved_val = options[0]
        st.session_state[persist_key] = saved_val

    w_key = f"_w_{persist_key}"
    st.session_state[w_key] = saved_val

    def _handle_change():
        new_val = st.session_state.get(w_key)
        if new_val in options:
            st.session_state[persist_key] = new_val
            if on_value_change is not None:
                on_value_change(new_val)

    kwargs = {
        "label": label,
        "options": options,
        "key": w_key,
        "disabled": disabled,
        "on_change": _handle_change,
        "label_visibility": label_visibility,
    }
    if format_func is not None:
        kwargs["format_func"] = format_func

    val = container.selectbox(**kwargs)
    if val in options:
        st.session_state[persist_key] = val
    return val


def persistent_text_input(
    container,
    label: str,
    persist_key: str,
    default_value: str = "",
    placeholder: str = "",
    disabled: bool = False,
) -> str:
    """Text input có State Persistence — giữ nguyên giá trị khi Save, đóng popup hoặc đổi tab."""
    if persist_key not in st.session_state:
        st.session_state[persist_key] = default_value

    w_key = f"_w_{persist_key}"
    st.session_state[w_key] = str(st.session_state[persist_key] or "")

    def _handle_change():
        st.session_state[persist_key] = str(st.session_state.get(w_key, ""))

    val = container.text_input(
        label,
        placeholder=placeholder,
        key=w_key,
        disabled=disabled,
        on_change=_handle_change,
    )
    st.session_state[persist_key] = str(val or "")
    return str(val or "").strip()


def persistent_date_input(
    container,
    label: str,
    persist_key: str,
    default_value: date = None,
    disabled: bool = False,
) -> date:
    """Date input có State Persistence."""
    fallback = default_value if isinstance(default_value, date) else date.today()
    if persist_key not in st.session_state or not isinstance(st.session_state[persist_key], date):
        st.session_state[persist_key] = fallback

    w_key = f"_w_{persist_key}"
    st.session_state[w_key] = st.session_state[persist_key]

    def _handle_change():
        picked = st.session_state.get(w_key)
        if isinstance(picked, date):
            st.session_state[persist_key] = picked

    val = container.date_input(
        label,
        key=w_key,
        format="DD/MM/YYYY",
        disabled=disabled,
        on_change=_handle_change,
    )
    if isinstance(val, date):
        st.session_state[persist_key] = val
        return val
    return fallback


def render_step_lock_header(
    title_md: str,
    lock_key: str,
    step_label: str = "Thông tin",
    is_mobile: bool = False,
) -> bool:
    """Thanh tiêu đề bước kèm nút Khóa / Mở khóa (Step-by-Step Locking)."""
    if lock_key not in st.session_state:
        st.session_state[lock_key] = False

    is_locked = bool(st.session_state[lock_key])

    def _toggle_lock():
        st.session_state[lock_key] = not bool(st.session_state.get(lock_key, False))

    lock_suffix = " &nbsp; `🔒 Đã khóa`" if is_locked else ""
    btn_label = (
        f":material/lock_open: Mở khóa {step_label}"
        if is_locked
        else f":material/lock: Khóa {step_label}"
    )

    if is_mobile:
        st.markdown(f"{title_md}{lock_suffix}")
        st.button(
            btn_label,
            key=f"_btn_lock_{lock_key}",
            type="primary" if is_locked else "secondary",
            use_container_width=True,
            on_click=_toggle_lock,
        )
    else:
        hdr_l, hdr_r = st.columns([3.6, 1.4], vertical_alignment="center")
        with hdr_l:
            st.markdown(f"{title_md}{lock_suffix}")
        with hdr_r:
            st.button(
                btn_label,
                key=f"_btn_lock_{lock_key}",
                type="primary" if is_locked else "secondary",
                use_container_width=True,
                on_click=_toggle_lock,
            )

    return is_locked


# ==============================================================================
# HỆ THỐNG POP-UP DIALOG HIỂN THỊ GIỮA MÀN HÌNH
# ==============================================================================
def queue_popup(title: str, message: str, status_type: str = "success"):
    """Lưu thông tin Pop-up vào session_state trước khi st.rerun() để không bị mất thông báo."""
    st.session_state["pending_popup"] = {
        "title": title,
        "message": message,
        "status_type": status_type
    }


if _dialog_decorator is not None:
    @_dialog_decorator("Thông báo hệ thống")
    def _show_notification_modal(title: str, message: str, status_type: str):
        icon_map = {
            "success": ("✅", "#15803D", "#DCFCE7", "#86EFAC"),
            "warning": ("⚠️", "#B45309", "#FEF3C7", "#FDE68A"),
            "error": ("🗑️", "#B91C1C", "#FEE2E2", "#FCA5A5"),
            "info": ("ℹ️", "#0369A1", "#E0F2FE", "#BAE6FD")
        }
        ic, color_text, bg_box, border_box = icon_map.get(status_type, icon_map["success"])

        render_html(f"""
        <div style="background-color:{bg_box}; border:1.5px solid {border_box}; border-radius:10px; padding:16px; text-align:center; margin-bottom:14px;">
            <div style="font-size:2rem; line-height:1.1; margin-bottom:8px;">{ic}</div>
            <div style="color:{color_text}; font-size:1.02rem; font-weight:800; margin-bottom:6px;">{title}</div>
            <div style="color:#0A2540; font-size:0.86rem; font-weight:500; line-height:1.45;">{message}</div>
        </div>
        """)
        if st.button("OK", type="primary", use_container_width=True, key="popup_ok_btn"):
            st.rerun()

    @_dialog_decorator("Tùy chọn Xuất file Word")
    def export_word_modal(docx_bytes: bytes, filename: str, key_suffix: str = ""):
        render_html(f"""
        <div style="background-color:#F4FBF6; border:1.5px solid #C8E6C9; border-radius:8px; padding:12px 16px; margin-bottom:16px; color:#0A2540; font-size:0.9rem;">
            📄 <strong>Tên file chuẩn:</strong><br>
            <code style="font-size:0.84rem; color:#1B5E20; word-break:break-all;">{filename}</code>
        </div>
        <div style="font-size:0.86rem; font-weight:600; color:#0A2540; margin-bottom:10px; text-align:center;">
            Vui lòng chọn thao tác xuất file bên dưới:
        </div>
        """)

        c1, c2, c3 = st.columns(3)
        with c1:
            if st.download_button(
                label="📥 Tải về máy",
                data=docx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"modal_dl_word_{key_suffix}"
            ):
                queue_popup(
                    "Đã tải file về máy!",
                    f"File <b>{filename}</b> đã được tải xuống máy tính của bạn.",
                    "success"
                )
                st.rerun()
        with c2:
            if st.button("☁️ Up lên Drive", type="primary", use_container_width=True, key=f"modal_up_drive_{key_suffix}"):
                with st.spinner("Đang tải file lên Google Drive..."):
                    upload_word_to_drive(docx_bytes, filename)
                st.rerun()
        with c3:
            if st.button("❌ Hủy", use_container_width=True, key=f"modal_cancel_word_{key_suffix}"):
                st.rerun()

    @_dialog_decorator("Xác nhận xóa đợt đánh giá")
    def confirm_delete_eval_modal(created_at: str, t_name: str, display_date_str: str, l_old: str, l_new: str):
        render_html(f"""
        <div style="background-color:#FEF3C7; border:1.5px solid #FDE68A; border-radius:10px; padding:14px 16px; margin-bottom:14px;">
            <div style="color:#B45309; font-size:0.95rem; font-weight:800; margin-bottom:6px;">⚠️ Bạn có chắc chắn muốn xóa đợt đánh giá này?</div>
            <div style="color:#0A2540; font-size:0.84rem; line-height:1.5;">
                • <strong>Xét nghiệm:</strong> {t_name}<br>
                • <strong>Ngày thực hiện:</strong> {display_date_str}<br>
                • <strong>Chuyển đổi LOT:</strong> <code>{l_old} ➔ {l_new}</code><br>
                <span style="color:#475569; font-size:0.78rem;">(Dữ liệu sau khi xóa sẽ được chuyển vào mục <b>6. Trash</b> ở thanh menu bên trái để khôi phục khi cần)</span>
            </div>
        </div>
        """)
        c_yes, c_no = st.columns(2)
        with c_yes:
            if st.button(":material/check: Xác nhận xóa", type="primary", use_container_width=True, key=f"modal_yes_{created_at}"):
                delete_stage2_evaluation(created_at)
                queue_popup(
                    "Đã chuyển vào Thùng rác",
                    f"Đợt đánh giá <b>{t_name}</b> ngày <b>{display_date_str}</b> (LOT: {l_old} ➔ {l_new}) đã được chuyển vào mục <b>6. Trash</b>.",
                    "warning"
                )
                st.rerun()
        with c_no:
            if st.button(":material/close: Hủy", use_container_width=True, key=f"modal_no_{created_at}"):
                st.rerun()

    @_dialog_decorator("Xác nhận xóa cấu hình Stage 1")
    def confirm_delete_setup_modal(del_test: str, del_mach: str):
        render_html(f"""
        <div style="background-color:#FEE2E2; border:1.5px solid #FCA5A5; border-radius:10px; padding:14px 16px; margin-bottom:14px;">
            <div style="color:#B91C1C; font-size:0.95rem; font-weight:800; margin-bottom:6px;">⚠️ Xác nhận xóa cấu hình xét nghiệm?</div>
            <div style="color:#0A2540; font-size:0.84rem; line-height:1.5;">
                Bạn đang chuẩn bị xóa cấu hình <strong>{del_test}</strong> trên máy <strong>{del_mach}</strong> khỏi cơ sở dữ liệu.
            </div>
        </div>
        """)
        c_yes, c_no = st.columns(2)
        with c_yes:
            if st.button(":material/delete: Xác nhận xóa", type="primary", use_container_width=True, key="modal_del_setup_yes"):
                delete_specific_setup(del_test, del_mach)
                queue_popup(
                    "Đã xóa cấu hình",
                    f"Đã xóa cấu hình xét nghiệm <b>{del_test}</b> ({del_mach}) khỏi Database.",
                    "error"
                )
                st.rerun()
        with c_no:
            if st.button(":material/close: Hủy", use_container_width=True, key="modal_del_setup_no"):
                st.rerun()


def render_pending_popup():
    """Kiểm tra và bật Pop-up ngay khi trang vừa tải lại sau thao tác."""
    if "pending_popup" in st.session_state and st.session_state["pending_popup"] is not None:
        data = st.session_state.pop("pending_popup")
        if _dialog_decorator is not None:
            _show_notification_modal(data["title"], data["message"], data["status_type"])
        else:
            st.success(f"**{data['title']}**: {data['message']}")


# ==============================================================================
# CSS GIAO DIỆN CHUNG + RESPONSIVE MOBILE UI
# ==============================================================================
def inject_custom_css(is_mobile: bool = False):
    mobile_override_css = ""
    if is_mobile:
        mobile_override_css = """
        .block-container,
        [data-testid="stMainBlockContainer"],
        [data-testid="stAppViewBlockContainer"] {
            padding-top: 2.9rem !important;
            padding-bottom: 1.5rem !important;
            padding-left: 0.45rem !important;
            padding-right: 0.45rem !important;
        }
        .page-banner {
            padding: 8px 12px !important;
            margin-bottom: 8px !important;
        }
        .page-banner-title {
            font-size: 0.98rem !important;
        }
        .page-banner-desc {
            font-size: 0.76rem !important;
        }
        [data-testid="stVerticalBlockBorderWrapper"] {
            padding: 0.55rem 0.55rem !important;
            margin-bottom: 0.55rem !important;
        }
        .main .stButton > button[kind="primary"],
        .main .stButton > button[kind="secondary"],
        .stDownloadButton > button {
            min-height: 42px !important;
            font-size: 0.84rem !important;
        }
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        div[data-baseweb="base-input"] {
            min-height: 40px !important;
            font-size: 0.86rem !important;
        }
        """

    render_html(f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600;700&display=swap');

        /* 1. TỔNG THỂ NỀN XANH LÁ NHẠT & CHỮ XANH TỐI DA TRỜI (#0A2540) */
        html, body, .stApp {{
            background-color: #F4FAF6 !important;
            color: #0A2540 !important;
            font-family: 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }}

        .block-container,
        [data-testid="stMainBlockContainer"],
        [data-testid="stAppViewBlockContainer"] {{
            padding-top: 3.8rem !important;
            padding-bottom: 2rem !important;
            padding-left: 1.6rem !important;
            padding-right: 1.6rem !important;
            max-width: 100% !important;
        }}

        .tabular-nums, [data-testid="stMetricValue"], [data-testid="stDataFrame"], table td {{
            font-variant-numeric: tabular-nums !important;
        }}

        [data-testid="stHeaderActionElements"],
        h1 a, h2 a, h3 a, h4 a, h5 a, h6 a {{
            display: none !important;
        }}

        div[role="dialog"] {{
            max-width: 440px !important;
            width: 92vw !important;
            border-radius: 14px !important;
            border: 1.5px solid #A5D6A7 !important;
            box-shadow: 0 14px 36px rgba(10, 37, 64, 0.22) !important;
            background-color: #FFFFFF !important;
        }}

        h1, h2, h3, h4, h5, h6 {{
            color: #0A2540 !important;
            font-family: 'Be Vietnam Pro', sans-serif !important;
            font-weight: 700 !important;
            letter-spacing: -0.2px;
            margin-top: 0 !important;
        }}
        h3 {{
            font-size: 0.98rem !important;
            font-weight: 800 !important;
            color: #0A2540 !important;
            margin-bottom: 0.55rem !important;
        }}
        h4, h5 {{
            font-size: 0.9rem !important;
            font-weight: 700 !important;
            color: #0A2540 !important;
            margin-bottom: 0.4rem !important;
        }}

        /* 3. NÚT ĐÓNG / MỞ SIDEBAR LUÔN HIỂN THỊ RÕ */
        [data-testid="stSidebarHeader"] {{
            display: flex !important;
            justify-content: flex-end !important;
            padding: 0.4rem 0.6rem 0 0.6rem !important;
        }}
        [data-testid="stSidebarCollapseButton"] button,
        [data-testid="collapsedControl"] button,
        [data-testid="stSidebarCollapsedControl"] button {{
            opacity: 1 !important;
            visibility: visible !important;
            background-color: #FFFFFF !important;
            color: #2E7D32 !important;
            border: 1.5px solid #66BB6A !important;
            border-radius: 8px !important;
            box-shadow: 0 1px 3px rgba(10, 37, 64, 0.08) !important;
            width: 34px !important;
            height: 34px !important;
        }}
        [data-testid="stSidebarCollapseButton"] button:hover,
        [data-testid="collapsedControl"] button:hover,
        [data-testid="stSidebarCollapsedControl"] button:hover {{
            background-color: #2E7D32 !important;
            color: #FFFFFF !important;
            border-color: #1B5E20 !important;
        }}
        [data-testid="stSidebarCollapseButton"] button svg,
        [data-testid="collapsedControl"] button svg,
        [data-testid="stSidebarCollapsedControl"] button svg {{
            fill: currentColor !important;
        }}

        /* 4. SIDEBAR MENU DỌC */
        [data-testid="stSidebar"] {{
            background-color: #E6F4EA !important;
        }}
        [data-testid="stSidebar"][aria-expanded="true"] {{
            border-right: 1.5px solid #C8E6C9 !important;
            min-width: 245px !important;
            max-width: 265px !important;
        }}
        [data-testid="stSidebar"][aria-expanded="false"] {{
            min-width: 0px !important;
            max-width: 0px !important;
            width: 0px !important;
            border-right: none !important;
        }}
        [data-testid="stSidebar"] > div:first-child {{
            background-color: #E6F4EA !important;
        }}
        .sidebar-brand {{
            background-color: #FFFFFF;
            padding: 11px 14px;
            border-radius: 10px;
            border: 1.5px solid #A5D6A7;
            border-left: 5px solid #2E7D32;
            margin-bottom: 0.9rem;
            box-shadow: 0 2px 4px rgba(10, 37, 64, 0.04);
        }}
        .sidebar-brand-title {{
            color: #0A2540 !important;
            font-size: 1.02rem;
            font-weight: 800;
            letter-spacing: -0.2px;
        }}
        .sidebar-brand-sub {{
            color: #2E7D32 !important;
            font-size: 0.75rem;
            font-weight: 600;
            margin-top: 2px;
        }}
        .sidebar-section-label {{
            color: #0A2540 !important;
            font-size: 0.72rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            padding: 0 4px;
            margin-bottom: 6px;
        }}

        [data-testid="stSidebar"] .stButton {{
            width: 100% !important;
            margin-bottom: 2px !important;
        }}
        [data-testid="stSidebar"] .stButton > button {{
            width: 100% !important;
            min-height: 40px !important;
            height: 40px !important;
            justify-content: flex-start !important;
            text-align: left !important;
            padding: 0 13px !important;
            border-radius: 8px !important;
            font-size: 0.84rem !important;
            font-weight: 700 !important;
            box-shadow: 0 1px 2px rgba(10, 37, 64, 0.04) !important;
            transition: all 0.15s ease !important;
        }}
        [data-testid="stSidebar"] .stButton > button > div {{
            justify-content: flex-start !important;
            width: 100% !important;
            gap: 8px !important;
        }}
        [data-testid="stSidebar"] .stButton > button[kind="secondary"] {{
            background-color: #FFFFFF !important;
            color: #0A2540 !important;
            border: 1.5px solid #C8E6C9 !important;
        }}
        [data-testid="stSidebar"] .stButton > button[kind="secondary"] p,
        [data-testid="stSidebar"] .stButton > button[kind="secondary"] span {{
            color: #0A2540 !important;
            font-weight: 700 !important;
            font-size: 0.84rem !important;
        }}
        [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {{
            background-color: #D7F0DD !important;
            border-color: #66BB6A !important;
        }}
        [data-testid="stSidebar"] .stButton > button[kind="primary"] {{
            background: linear-gradient(135deg, #2E7D32 0%, #388E3C 100%) !important;
            color: #FFFFFF !important;
            border: 1.5px solid #1B5E20 !important;
            box-shadow: 0 3px 6px rgba(46, 125, 50, 0.22) !important;
        }}
        [data-testid="stSidebar"] .stButton > button[kind="primary"] p,
        [data-testid="stSidebar"] .stButton > button[kind="primary"] span {{
            color: #FFFFFF !important;
            font-weight: 800 !important;
            font-size: 0.84rem !important;
        }}

        /* 5. BANNER ĐẦU TRANG */
        .page-banner {{
            background: linear-gradient(90deg, #E6F4EA 0%, #FFFFFF 100%);
            border: 1px solid #A5D6A7;
            border-left: 5px solid #2E7D32;
            border-radius: 8px;
            padding: 10px 16px;
            margin-top: 4px;
            margin-bottom: 12px;
            box-shadow: 0 1px 3px rgba(10, 37, 64, 0.03);
        }}
        .page-banner-title {{
            color: #0A2540;
            font-size: 1.05rem;
            font-weight: 800;
            margin: 0 0 2px 0;
        }}
        .page-banner-desc {{
            color: #1E3A5F;
            font-size: 0.8rem;
            font-weight: 500;
            margin: 0;
        }}

        .search-result-pill {{
            background-color: #E6F4EA;
            border: 1px solid #A5D6A7;
            color: #1B5E20;
            padding: 5px 12px;
            border-radius: 6px;
            font-size: 0.78rem;
            font-weight: 600;
            margin-top: -4px;
            margin-bottom: 10px;
            display: inline-block;
        }}

        /* 6. KHỐI CARD TRẮNG & EXPANDER */
        [data-testid="stVerticalBlockBorderWrapper"] {{
            background-color: #FFFFFF !important;
            border: 1px solid #C8E6C9 !important;
            border-radius: 8px !important;
            box-shadow: 0 1px 4px rgba(10, 37, 64, 0.03) !important;
            padding: 0.75rem 0.95rem !important;
            margin-bottom: 0.65rem !important;
        }}

        [data-testid="stExpander"] {{
            background-color: #FFFFFF !important;
            border: 1px solid #CBE6D0 !important;
            border-radius: 7px !important;
            margin-bottom: 5px !important;
            box-shadow: none !important;
            overflow: hidden;
        }}
        [data-testid="stExpander"] summary {{
            background-color: #FFFFFF !important;
            padding: 6px 12px !important;
            border-radius: 6px !important;
            transition: background-color 0.12s ease;
        }}
        [data-testid="stExpander"] summary:hover {{
            background-color: #EEF8F0 !important;
        }}
        [data-testid="stExpander"] summary p {{
            color: #0A2540 !important;
            font-weight: 600 !important;
            font-size: 0.80rem !important;
            line-height: 1.3 !important;
        }}
        [data-testid="stExpanderDetails"] {{
            padding: 8px 12px 10px 12px !important;
            border-top: 1px solid #E8F5E9 !important;
            background-color: #FAFDFB !important;
        }}

        /* 7. Ô NHẬP LIỆU, SELECTBOX & TRẠNG THÁI KHÓA */
        .stSelectbox label p, .stTextInput label p, .stDateInput label p,
        .stNumberInput label p, .stFileUploader label p {{
            color: #0A2540 !important;
            font-weight: 700 !important;
            font-size: 0.82rem !important;
        }}
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        div[data-baseweb="base-input"] {{
            background-color: #FFFFFF !important;
            border: 1px solid #A5D6A7 !important;
            border-radius: 7px !important;
            color: #0A2540 !important;
            font-weight: 600 !important;
            font-size: 0.83rem !important;
            min-height: 36px !important;
        }}
        div[data-baseweb="select"][aria-disabled="true"] > div,
        div[data-baseweb="input"]:has(input:disabled),
        div[data-baseweb="base-input"]:has(input:disabled) {{
            background-color: #EFF6F0 !important;
            border-color: #B7DFB9 !important;
            opacity: 0.92 !important;
        }}
        div[data-baseweb="select"] > div:focus-within,
        div[data-baseweb="input"] > div:focus-within {{
            border-color: #2E7D32 !important;
            box-shadow: 0 0 0 1px #2E7D32 !important;
        }}

        div[data-baseweb="select"],
        div[data-baseweb="select"] * {{
            cursor: pointer !important;
            user-select: none !important;
            -webkit-user-select: none !important;
        }}
        div[data-baseweb="select"] input {{
            caret-color: transparent !important;
            pointer-events: none !important;
            cursor: pointer !important;
            user-select: none !important;
            -webkit-user-select: none !important;
        }}

        /* 8. NÚT BẤM TRONG VÙNG NỘI DUNG CHÍNH */
        .main .stButton > button[kind="primary"],
        div[role="dialog"] .stButton > button[kind="primary"] {{
            background: linear-gradient(135deg, #388E3C 0%, #2E7D32 100%) !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            font-size: 0.82rem !important;
            border: none !important;
            border-radius: 7px !important;
            padding: 0.4rem 0.9rem !important;
            min-height: 36px !important;
        }}
        .main .stButton > button[kind="primary"] p,
        div[role="dialog"] .stButton > button[kind="primary"] p {{
            color: #FFFFFF !important;
            font-weight: 700 !important;
            font-size: 0.82rem !important;
        }}
        .main .stButton > button[kind="secondary"],
        .stDownloadButton > button,
        div[role="dialog"] .stButton > button[kind="secondary"] {{
            background-color: #E8F5E9 !important;
            color: #0A2540 !important;
            font-weight: 600 !important;
            font-size: 0.80rem !important;
            border: 1px solid #81C784 !important;
            border-radius: 7px !important;
            padding: 0.35rem 0.8rem !important;
            min-height: 36px !important;
        }}
        .main .stButton > button[kind="secondary"] p,
        .stDownloadButton > button p,
        div[role="dialog"] .stButton > button[kind="secondary"] p {{
            color: #0A2540 !important;
            font-weight: 600 !important;
            font-size: 0.80rem !important;
        }}

        /* 9. METRICS */
        [data-testid="stMetric"] {{
            background-color: #F1F8F3 !important;
            border: 1px solid #C8E6C9 !important;
            border-radius: 8px !important;
            padding: 7px 10px !important;
        }}
        [data-testid="stMetricLabel"] p {{
            color: #1B4F72 !important;
            font-weight: 700 !important;
            font-size: 0.76rem !important;
        }}
        [data-testid="stMetricValue"] {{
            color: #0A2540 !important;
            font-family: 'JetBrains Mono', monospace !important;
            font-weight: 800 !important;
            font-size: 1.05rem !important;
        }}

        /* 10. NHÃN TRẠNG THÁI (BADGES) */
        .status-badge {{
            display: inline-flex;
            align-items: center;
            gap: 3px;
            padding: 1px 8px;
            border-radius: 999px;
            font-size: 0.72rem;
            font-weight: 700;
            line-height: 1.35;
            white-space: nowrap;
        }}
        .badge-pass {{
            background-color: #DCFCE7;
            color: #15803D;
            border: 1px solid #86EFAC;
        }}
        .badge-fail {{
            background-color: #FEE2E2;
            color: #B91C1C;
            border: 1px solid #FCA5A5;
        }}
        .badge-warn {{
            background-color: #FEF3C7;
            color: #B45309;
            border: 1px solid #FDE68A;
        }}

        .eval-result-box {{
            background-color: #FFFFFF;
            border: 1px solid #C8E6C9;
            border-radius: 7px;
            padding: 7px 10px;
            margin-top: 6px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            font-size: 0.82rem;
            color: #0A2540;
        }}
        .eval-result-pass {{
            border-left: 4px solid #2E7D32;
            background-color: #F4FBF6;
        }}
        .eval-result-fail {{
            border-left: 4px solid #C62828;
            background-color: #FEF2F2;
        }}

        /* 11. LỊCH SỬ & BẢNG */
        .month-header-banner {{
            background-color: #E6F4EA;
            color: #0A2540;
            padding: 5px 12px;
            border: 1px solid #B9E2C2;
            border-left: 4px solid #2E7D32;
            border-radius: 6px;
            font-weight: 700;
            font-size: 0.82rem;
            margin-top: 8px;
            margin-bottom: 6px;
        }}
        .history-meta-bar {{
            background-color: #FFFFFF;
            border: 1px solid #DCEEE0;
            border-radius: 6px;
            padding: 6px 10px;
            margin-bottom: 8px;
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: center;
            color: #0A2540;
            font-size: 0.78rem;
        }}
        .date-pill {{
            background-color: #2E7D32;
            color: #FFFFFF;
            padding: 1px 7px;
            border-radius: 4px;
            font-family: 'JetBrains Mono', monospace;
            font-weight: 600;
            font-size: 0.75rem;
        }}
        .lot-pill {{
            background-color: #E3F2FD;
            color: #0A2540;
            border: 1px solid #90CAF9;
            padding: 1px 7px;
            border-radius: 4px;
            font-family: 'JetBrains Mono', monospace;
            font-weight: 600;
            font-size: 0.75rem;
        }}

        .lab-table-wrapper {{
            max-height: 280px;
            overflow-x: auto;
            overflow-y: auto;
            border: 1px solid #DCEEE0;
            border-radius: 6px;
            background-color: #FFFFFF;
            margin-bottom: 4px;
        }}
        .lab-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.77rem;
            color: #0A2540;
        }}
        .lab-table thead th {{
            position: sticky;
            top: 0;
            background-color: #F1F8F3;
            color: #334E68;
            font-weight: 700;
            padding: 5px 8px;
            border-bottom: 1px solid #C8E6C9;
            white-space: nowrap;
            z-index: 1;
        }}
        .lab-table tbody td {{
            padding: 4px 8px;
            border-bottom: 1px solid #F0F7F1;
            font-variant-numeric: tabular-nums;
            font-weight: 500;
            white-space: nowrap;
        }}
        .lab-table tbody tr:nth-child(even) {{
            background-color: #FAFDFB;
        }}
        .lab-table tbody tr:hover {{
            background-color: #E8F5E9;
        }}
        .text-left {{ text-align: left !important; }}
        .text-right {{ text-align: right !important; font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; }}
        .text-center {{ text-align: center !important; }}

        /* 12. DANH SÁCH TỔNG QUAN LOT (DESKTOP COMPACT LIST) */
        .overview-list-card {{
            background-color: #FFFFFF;
            border: 1px solid #DCEEE0;
            border-radius: 8px;
            margin-top: 6px;
            margin-bottom: 10px;
            overflow: hidden;
        }}
        .overview-group-header {{
            display: grid;
            grid-template-columns: 2.3fr 1.2fr 1fr 0.9fr 0.8fr;
            align-items: center;
            padding: 6px 12px;
            background-color: #F4FAF6;
            border-bottom: 1px solid #DCEEE0;
            font-size: 0.73rem;
            font-weight: 700;
            color: #52667A;
        }}
        .overview-mach-title {{
            color: #0A2540;
            font-weight: 800;
            font-size: 0.78rem;
            display: flex;
            align-items: center;
            gap: 6px;
        }}
        details.overview-item {{
            border-bottom: 1px solid #F0F5F1;
            background-color: #FFFFFF;
        }}
        details.overview-item:last-child {{
            border-bottom: none;
        }}
        summary.overview-row {{
            display: grid;
            grid-template-columns: 2.3fr 1.2fr 1fr 0.9fr 0.8fr;
            align-items: center;
            padding: 6px 12px;
            font-size: 0.78rem;
            color: #0A2540;
            cursor: pointer;
            list-style: none;
            transition: background-color 0.12s ease;
        }}
        summary.overview-row::-webkit-details-marker {{
            display: none;
        }}
        summary.overview-row:hover {{
            background-color: #F4FBF6;
        }}
        .item-name-cell {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-weight: 700;
            color: #0A2540;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        .item-avatar {{
            width: 21px;
            height: 21px;
            border-radius: 50%;
            background-color: #E6F4EA;
            border: 1px dashed #66BB6A;
            color: #2E7D32;
            font-size: 0.68rem;
            font-weight: 800;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
        }}
        .item-lot-cell {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.76rem;
            font-weight: 700;
            color: #1B5E20;
        }}
        .item-date-cell {{
            color: #64748B;
            font-size: 0.76rem;
            font-variant-numeric: tabular-nums;
        }}
        .item-action-cell {{
            text-align: right;
            font-size: 0.74rem;
            font-weight: 600;
            color: #2E7D32;
        }}
        .overview-subbox {{
            padding: 6px 14px 10px 38px;
            background-color: #FAFDFB;
            border-top: 1px dashed #DCEEE0;
        }}

        /* 13. GIAO DIỆN MOBILE CARDS CHUYÊN BIỆT */
        .mobile-topbar {{
            background: linear-gradient(135deg, #2E7D32 0%, #388E3C 100%);
            color: #FFFFFF;
            padding: 10px 14px;
            border-radius: 10px;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 3px 8px rgba(46, 125, 50, 0.22);
        }}
        .mobile-topbar-title {{
            font-size: 0.96rem;
            font-weight: 800;
            color: #FFFFFF;
        }}
        .mobile-topbar-badge {{
            background-color: rgba(255, 255, 255, 0.22);
            color: #FFFFFF;
            font-size: 0.68rem;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 999px;
        }}
        .m-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 6px;
            margin: 6px 0 8px 0;
        }}
        .m-kpi-card {{
            background-color: #F1F8F3;
            border: 1px solid #C8E6C9;
            border-radius: 8px;
            padding: 6px 8px;
            text-align: center;
        }}
        .m-kpi-lbl {{
            font-size: 0.67rem;
            font-weight: 700;
            color: #1B4F72;
            margin-bottom: 2px;
        }}
        .m-kpi-val {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.88rem;
            font-weight: 800;
            color: #0A2540;
        }}
        .m-card {{
            background-color: #FFFFFF;
            border: 1px solid #CBE6D0;
            border-radius: 10px;
            padding: 10px 12px;
            margin-bottom: 8px;
            box-shadow: 0 1px 3px rgba(10, 37, 64, 0.04);
        }}
        .m-card-head {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
            margin-bottom: 6px;
        }}
        .m-card-title {{
            font-weight: 800;
            font-size: 0.88rem;
            color: #0A2540;
            display: flex;
            align-items: center;
            gap: 7px;
        }}
        .m-card-meta {{
            font-size: 0.76rem;
            color: #334E68;
            line-height: 1.5;
        }}
        .m-sample-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 4px 0;
            border-bottom: 1px dashed #E2EFE5;
            font-size: 0.76rem;
            font-variant-numeric: tabular-nums;
        }}
        .m-sample-row:last-child {{
            border-bottom: none;
        }}
        details.m-lot-card {{
            background-color: #FFFFFF;
            border: 1px solid #CBE6D0;
            border-radius: 10px;
            margin-bottom: 8px;
            overflow: hidden;
        }}
        summary.m-lot-summary {{
            padding: 10px 12px;
            list-style: none;
            cursor: pointer;
        }}
        summary.m-lot-summary::-webkit-details-marker {{
            display: none;
        }}
        .m-lot-details {{
            padding: 8px 12px 10px 12px;
            background-color: #FAFDFB;
            border-top: 1px dashed #DCEEE0;
        }}

        /* 14. TỐI ƯU BẢNG NHẬP MẪU (st.data_editor) VỪA KHÍT MÀN HÌNH MOBILE & DESKTOP */
        .s2-level-meta {{
            display: flex;
            flex-wrap: wrap;
            gap: 6px 10px;
            align-items: center;
            background-color: #F4FAF6;
            border: 1px solid #DCEEE0;
            border-radius: 7px;
            padding: 6px 10px;
            margin-bottom: 10px;
            font-size: 0.78rem;
            color: #0A2540;
        }}
        /* Ẩn thanh công cụ nổi (+, mắt, tải xuống...) hay đè che mất dòng Giới hạn RL trên điện thoại */
        [data-testid="stDataFrame"] [data-testid="stElementToolbar"],
        [data-testid="stDataEditor"] [data-testid="stElementToolbar"] {{
            display: none !important;
        }}
        [data-testid="stDataFrame"],
        [data-testid="stDataEditor"] {{
            width: 100% !important;
            max-width: 100% !important;
        }}
        /* Giữ cặp nút + Thêm dòng / - Bớt dòng luôn nằm ngang 50/50 gọn trên điện thoại */
        [data-testid="stHorizontalBlock"]:has([class*="st-key-s2_btn_add_"]) {{
            flex-wrap: nowrap !important;
            gap: 0.45rem !important;
            margin-top: 4px !important;
        }}
        [data-testid="stHorizontalBlock"]:has([class*="st-key-s2_btn_add_"]) > [data-testid="stColumn"] {{
            min-width: 0 !important;
            flex: 1 1 50% !important;
            width: 50% !important;
        }}

        @media (max-width: 768px) {{
            .block-container,
            [data-testid="stMainBlockContainer"],
            [data-testid="stAppViewBlockContainer"] {{
                padding-top: 2.9rem !important;
                padding-left: 0.7rem !important;
                padding-right: 0.7rem !important;
            }}
        }}

        {mobile_override_css}
    </style>
    """)


def render_page_banner(title: str, desc: str, is_mobile: bool = False):
    desc_html = f'<div class="page-banner-desc">{desc}</div>' if (desc and not is_mobile) else ""
    render_html(f"""
    <div class="page-banner">
        <div class="page-banner-title">{title}</div>
        {desc_html}
    </div>
    """)


def format_status_badge_html(status_val: str) -> str:
    s = str(status_val).strip().upper()
    if s in ("PASS", "ĐẠT", "✅ ĐẠT"):
        return '<span class="status-badge badge-pass">✓ Đạt</span>'
    if s in ("FAIL", "TRƯỢT", "KHÔNG ĐẠT", "❌ TRƯỢT"):
        return '<span class="status-badge badge-fail">✕ Không đạt</span>'
    return f'<span class="status-badge badge-warn">! {status_val}</span>'


# ==============================================================================
# TRANG 1: ĐÁNH GIÁ MẪU THỰC TẾ (STAGE 2) — KHÓA CỨNG + STATE PERSISTENCE
# ==============================================================================
def _make_empty_samples_df(n_rows: int) -> pd.DataFrame:
    """Tạo bảng mẫu trống với đúng kiểu dữ liệu để mở bàn phím số gọn trên điện thoại."""
    n = max(1, int(n_rows))
    return pd.DataFrame({
        "SID (Mã mẫu)": pd.Series([""] * n, dtype="object"),
        "Kết quả LOT cũ": pd.Series([float("nan")] * n, dtype="float64"),
        "Kết quả LOT mới": pd.Series([float("nan")] * n, dtype="float64"),
    })


def render_stage2_workspace(active_setups: pd.DataFrame, is_mobile: bool = False):
    if "_persist_l2l_samples" not in st.session_state:
        st.session_state["_persist_l2l_samples"] = {}
    if "_seed_l2l_samples" not in st.session_state:
        st.session_state["_seed_l2l_samples"] = {}
    if "_l2l_row_ver" not in st.session_state:
        st.session_state["_l2l_row_ver"] = {}
    if "_l2l_clear_cnt" not in st.session_state:
        st.session_state["_l2l_clear_cnt"] = 0

    with st.container(border=True):
        info_locked = render_step_lock_header(
            title_md="#### :material/assignment: 1. Thông tin đợt đánh giá",
            lock_key="_lock_l2l_step1",
            step_label="Thông tin",
            is_mobile=is_mobile,
        )

        test_opts = active_setups["test_name"].unique().tolist()

        if is_mobile:
            eval_test = persistent_selectbox(
                st,
                label="Chọn Xét nghiệm:",
                options=test_opts,
                persist_key="_persist_l2l_test",
                default_value=test_opts[0],
                disabled=info_locked,
            )
            mach_opts = active_setups[active_setups["test_name"] == eval_test]["machine"].unique().tolist()
            eval_machine = persistent_selectbox(
                st,
                label="Dòng máy:",
                options=mach_opts,
                persist_key=f"_persist_l2l_mach::{eval_test}",
                default_value=mach_opts[0],
                disabled=info_locked,
            )
            eval_date = persistent_date_input(
                st,
                label="Ngày thực hiện:",
                persist_key="_persist_l2l_date",
                default_value=date.today(),
                disabled=info_locked,
            )
            reagent_name = persistent_text_input(
                st,
                label="Loại hóa chất thay thế (Huyết học/HbA1c):",
                persist_key="_persist_l2l_reagent",
                default_value="",
                placeholder="VD: CELL-DYN Diluent/Sheath... (Bỏ trống nếu Sinh hóa)",
                disabled=info_locked,
            )
            mc1, mc2 = st.columns(2)
            lot_old = persistent_text_input(
                mc1,
                label="Số LOT cũ:",
                persist_key="_persist_l2l_lot_old",
                default_value="82273UD00",
                disabled=info_locked,
            )
            lot_new = persistent_text_input(
                mc2,
                label="Số LOT mới:",
                persist_key="_persist_l2l_lot_new",
                default_value="86280UD00",
                disabled=info_locked,
            )
        else:
            c1, c2, c3 = st.columns([1.5, 1.5, 1])
            eval_test = persistent_selectbox(
                c1,
                label="Chọn Xét nghiệm:",
                options=test_opts,
                persist_key="_persist_l2l_test",
                default_value=test_opts[0],
                disabled=info_locked,
            )
            mach_opts = active_setups[active_setups["test_name"] == eval_test]["machine"].unique().tolist()
            eval_machine = persistent_selectbox(
                c2,
                label="Dòng máy:",
                options=mach_opts,
                persist_key=f"_persist_l2l_mach::{eval_test}",
                default_value=mach_opts[0],
                disabled=info_locked,
            )
            eval_date = persistent_date_input(
                c3,
                label="Ngày thực hiện:",
                persist_key="_persist_l2l_date",
                default_value=date.today(),
                disabled=info_locked,
            )

            c4, c5, c6 = st.columns([2, 1, 1])
            reagent_name = persistent_text_input(
                c4,
                label="Loại hóa chất thay thế (Huyết học/HbA1c):",
                persist_key="_persist_l2l_reagent",
                default_value="",
                placeholder="VD: CELL-DYN Diluent/Sheath... (Bỏ trống nếu là Sinh hóa)",
                disabled=info_locked,
            )
            lot_old = persistent_text_input(
                c5,
                label="Số LOT cũ:",
                persist_key="_persist_l2l_lot_old",
                default_value="82273UD00",
                disabled=info_locked,
            )
            lot_new = persistent_text_input(
                c6,
                label="Số LOT mới:",
                persist_key="_persist_l2l_lot_new",
                default_value="86280UD00",
                disabled=info_locked,
            )

        eval_machine_id = ""
        mach_df = active_setups[
            (active_setups["test_name"] == eval_test) & (active_setups["machine"] == eval_machine)
        ]
        if not mach_df.empty and "machine_id" in mach_df.columns and pd.notna(mach_df["machine_id"].iloc[0]):
            eval_machine_id = str(mach_df["machine_id"].iloc[0])

    test_levels_df = mach_df.sort_values("target_conc").reset_index(drop=True)
    unit_s2 = str(test_levels_df["unit"].iloc[0]) if not test_levels_df.empty else "mmol/L"

    def _clear_current_samples():
        for idx_c in range(len(test_levels_df)):
            sig_c = f"{eval_test}::{eval_machine}::{idx_c + 1}"
            st.session_state["_persist_l2l_samples"].pop(sig_c, None)
            st.session_state["_seed_l2l_samples"].pop(sig_c, None)
            st.session_state["_l2l_row_ver"][sig_c] = st.session_state["_l2l_row_ver"].get(sig_c, 0) + 1
        st.session_state["_l2l_clear_cnt"] += 1

    if is_mobile:
        st.markdown("### :material/biotech: 2. Nhập kết quả mẫu")
        st.button(
            ":material/restart_alt: Làm trống bảng mẫu (Đợt mới)",
            type="secondary",
            use_container_width=True,
            key="btn_clear_s2_samples_m",
            on_click=_clear_current_samples,
        )
    else:
        s2_h1, s2_h2 = st.columns([3.6, 1.4], vertical_alignment="center")
        with s2_h1:
            st.markdown("### :material/biotech: 2. Nhập kết quả mẫu")
        with s2_h2:
            st.button(
                ":material/restart_alt: Làm trống bảng mẫu",
                type="secondary",
                use_container_width=True,
                key="btn_clear_s2_samples_d",
                on_click=_clear_current_samples,
            )

    flat_records_to_save = []
    summary_table_rows = []
    clear_ver = st.session_state["_l2l_clear_cnt"]
    safe_sig_prefix = re.sub(r"[^a-zA-Z0-9_]", "_", f"{eval_test}_{eval_machine}")

    for idx, lv_row in test_levels_df.iterrows():
        level_num = idx + 1
        target_val = float(lv_row["target_conc"])
        rl_pct_db = float(lv_row["rejection_limit"])
        k_db = float(lv_row["k_factor"])
        n_req = max(3, int(lv_row["n_samples"]))

        sample_sig = f"{eval_test}::{eval_machine}::{level_num}"
        row_ver = st.session_state["_l2l_row_ver"].get(sample_sig, 0)
        editor_key = f"s2_tbl_{safe_sig_prefix}_L{level_num}_c{clear_ver}_r{row_ver}"

        # Chỉ nạp seed khi widget chưa tồn tại (mới mở/quay lại tab/vừa thêm bớt dòng) để không bị reset khi đang gõ
        if editor_key not in st.session_state:
            saved_df_lv = st.session_state["_persist_l2l_samples"].get(sample_sig)
            if isinstance(saved_df_lv, pd.DataFrame) and not saved_df_lv.empty:
                st.session_state["_seed_l2l_samples"][sample_sig] = saved_df_lv.copy().reset_index(drop=True)
            else:
                st.session_state["_seed_l2l_samples"][sample_sig] = _make_empty_samples_df(n_req)

        seed_df = st.session_state["_seed_l2l_samples"].get(sample_sig, _make_empty_samples_df(n_req))

        def _add_table_row_cb(sig=sample_sig):
            cur_df = st.session_state["_persist_l2l_samples"].get(sig)
            if not isinstance(cur_df, pd.DataFrame) or cur_df.empty:
                cur_df = _make_empty_samples_df(1)
            new_row_df = _make_empty_samples_df(1)
            merged = pd.concat([cur_df.reset_index(drop=True), new_row_df], ignore_index=True)
            st.session_state["_persist_l2l_samples"][sig] = merged
            st.session_state["_seed_l2l_samples"][sig] = merged.copy()
            st.session_state["_l2l_row_ver"][sig] = st.session_state["_l2l_row_ver"].get(sig, 0) + 1

        def _del_table_row_cb(sig=sample_sig):
            cur_df = st.session_state["_persist_l2l_samples"].get(sig)
            if isinstance(cur_df, pd.DataFrame) and len(cur_df) > 1:
                trimmed = cur_df.iloc[:-1].copy().reset_index(drop=True)
                st.session_state["_persist_l2l_samples"][sig] = trimmed
                st.session_state["_seed_l2l_samples"][sig] = trimmed.copy()
                st.session_state["_l2l_row_ver"][sig] = st.session_state["_l2l_row_ver"].get(sig, 0) + 1

        # Tiêu đề cột ngắn gọn + width="small" + hide_index=True + num_rows="fixed"
        # => Loại bỏ cột thừa bên trái, vừa khít 100% bề ngang điện thoại và máy tính mà không phải kéo ngang!
        col_config_lv = {
            "SID (Mã mẫu)": st.column_config.TextColumn("SID", width="small"),
            "Kết quả LOT cũ": st.column_config.NumberColumn("LOT cũ", format="%.3f", width="small"),
            "Kết quả LOT mới": st.column_config.NumberColumn("LOT mới", format="%.3f", width="small"),
        }

        with st.container(border=True):
            st.markdown(f"#### :material/target: Mức nồng độ: `{target_val:.3f} {unit_s2}`")
            render_html(f"""
            <div class="s2-level-meta">
                <span>📋 <b>Yêu cầu:</b> <code>{int(lv_row['n_samples'])} mẫu</code></span>
                <span>🎯 <b>Giới hạn RL:</b> <code>{rl_pct_db:.3f}%</code></span>
                <span>🧪 <b>LOT:</b> <span class="lot-pill">{lot_old or '-'} ➔ {lot_new or '-'}</span></span>
            </div>
            """)

            if is_mobile:
                edited_samples = st.data_editor(
                    seed_df,
                    num_rows="fixed",
                    hide_index=True,
                    use_container_width=True,
                    column_config=col_config_lv,
                    key=editor_key,
                )
                st.session_state["_persist_l2l_samples"][sample_sig] = edited_samples.copy()

                b_add, b_del = st.columns(2)
                b_add.button(
                    ":material/add: Thêm dòng",
                    key=f"s2_btn_add_m_{safe_sig_prefix}_L{level_num}",
                    type="secondary",
                    use_container_width=True,
                    on_click=_add_table_row_cb,
                )
                b_del.button(
                    ":material/remove: Bớt dòng",
                    key=f"s2_btn_del_m_{safe_sig_prefix}_L{level_num}",
                    type="secondary",
                    use_container_width=True,
                    disabled=(len(edited_samples) <= 1),
                    on_click=_del_table_row_cb,
                )

                eval_res = evaluate_stage2_level(samples_df=edited_samples, rl_pct=rl_pct_db)

                if not eval_res["has_data"]:
                    st.info("Nhập ít nhất 1 cặp kết quả (> 0) để xem đánh giá.", icon=":material/info:")
                else:
                    render_html(f"""
                    <div class="m-kpi-grid">
                        <div class="m-kpi-card">
                            <div class="m-kpi-lbl">TB |d|</div>
                            <div class="m-kpi-val">{eval_res['mean_d']:.4f}</div>
                        </div>
                        <div class="m-kpi-card">
                            <div class="m-kpi-lbl">TB |d| (%)</div>
                            <div class="m-kpi-val">{eval_res['mean_d_pct']:.3f}%</div>
                        </div>
                        <div class="m-kpi-card">
                            <div class="m-kpi-lbl">RL (%)</div>
                            <div class="m-kpi-val">{eval_res['rl_pct']:.3f}%</div>
                        </div>
                    </div>
                    """)
                    if eval_res["result"] == "PASS":
                        render_html(f"""
                        <div class="eval-result-box eval-result-pass">
                            <div><strong>Kết luận:</strong> <span class="tabular-nums">{eval_res['mean_d_pct']:.3f}% &lt; {eval_res['rl_pct']:.3f}%</span></div>
                            {format_status_badge_html('PASS')}
                        </div>
                        """)
                    else:
                        render_html(f"""
                        <div class="eval-result-box eval-result-fail">
                            <div><strong>Kết luận:</strong> <span class="tabular-nums">{eval_res['mean_d_pct']:.3f}% &ge; {eval_res['rl_pct']:.3f}%</span></div>
                            {format_status_badge_html('FAIL')}
                        </div>
                        """)
            else:
                in_col, out_col = st.columns([2.25, 1.75], gap="medium")
                with in_col:
                    edited_samples = st.data_editor(
                        seed_df,
                        num_rows="fixed",
                        hide_index=True,
                        use_container_width=True,
                        column_config=col_config_lv,
                        key=editor_key,
                    )
                    st.session_state["_persist_l2l_samples"][sample_sig] = edited_samples.copy()

                    ba1, ba2, _ = st.columns([1.2, 1.2, 1.6])
                    ba1.button(
                        ":material/add: Thêm dòng",
                        key=f"s2_btn_add_d_{safe_sig_prefix}_L{level_num}",
                        type="secondary",
                        use_container_width=True,
                        on_click=_add_table_row_cb,
                    )
                    ba2.button(
                        ":material/remove: Bớt dòng",
                        key=f"s2_btn_del_d_{safe_sig_prefix}_L{level_num}",
                        type="secondary",
                        use_container_width=True,
                        disabled=(len(edited_samples) <= 1),
                        on_click=_del_table_row_cb,
                    )

                eval_res = evaluate_stage2_level(samples_df=edited_samples, rl_pct=rl_pct_db)

                with out_col:
                    if not eval_res["has_data"]:
                        st.info("Nhập ít nhất 1 cặp kết quả (> 0) để xem đánh giá.", icon=":material/info:")
                    else:
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Trung bình |d|", f"{eval_res['mean_d']:.4f}")
                        m2.metric("TB độ chệch (%)", f"{eval_res['mean_d_pct']:.3f}%")
                        m3.metric("Giới hạn RL", f"{eval_res['rl_pct']:.3f}%")

                        if eval_res["result"] == "PASS":
                            render_html(f"""
                            <div class="eval-result-box eval-result-pass">
                                <div><strong>Kết luận:</strong> <span class="tabular-nums">{eval_res['mean_d_pct']:.3f}% &lt; {eval_res['rl_pct']:.3f}%</span></div>
                                {format_status_badge_html('PASS')}
                            </div>
                            """)
                        else:
                            render_html(f"""
                            <div class="eval-result-box eval-result-fail">
                                <div><strong>Kết luận:</strong> <span class="tabular-nums">{eval_res['mean_d_pct']:.3f}% &ge; {eval_res['rl_pct']:.3f}%</span></div>
                                {format_status_badge_html('FAIL')}
                            </div>
                            """)

            if eval_res["has_data"]:
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
            queue_popup(
                "Lưu kết quả thành công!",
                f"Đã lưu biên bản đánh giá <b>{eval_test}</b> ({eval_machine}) — LOT: <b>{lot_old} ➔ {lot_new}</b> vào lịch sử.<br>Toàn bộ thông tin đã nhập vẫn được giữ nguyên để bạn tiếp tục đối chiếu hoặc sang mục <b>2. Evaluation History</b> xuất file Word.",
                "success"
            )
            st.rerun()


# ==============================================================================
# TRANG 2: LỊCH SỬ ĐÁNH GIÁ (ĐÃ TÁCH THÙNG RÁC SANG SIDEBAR + MOBILE CARDS)
# ==============================================================================
def render_history_workspace(hist_df_master: pd.DataFrame, all_setups: pd.DataFrame, is_mobile: bool = False):
    hist_df = hist_df_master.copy()

    if hist_df.empty:
        st.info("Chưa có đợt đánh giá nào được lưu.", icon=":material/info:")
        return

    hist_search = st.text_input(
        "Tìm kiếm lịch sử đánh giá:",
        placeholder="🔍 Tra cứu (Ngày, Tên XN, Máy, Hóa chất, số LOT)..." if is_mobile else "🔍 Tra cứu nhanh (nhập Ngày, Tên XN, Máy, Hóa chất hoặc số LOT — VD: 13/11, ALT 13/11, WBC Ruby 04/11, 980645...)",
        key="search_eval_history_top",
        label_visibility="collapsed"
    ).strip().lower()

    hist_df['eval_date'] = pd.to_datetime(hist_df['eval_date'], errors='coerce')
    hist_df = hist_df.dropna(subset=['eval_date'])
    hist_df = hist_df.sort_values(by=['eval_date', 'created_at'], ascending=[False, False])
    hist_df['YearMonth'] = hist_df['eval_date'].dt.strftime('%m / %Y')
    hist_df['display_date_str'] = hist_df['eval_date'].dt.strftime('%d/%m/%Y')
    hist_df['e_date_str'] = hist_df['eval_date'].dt.strftime('%Y-%m-%d')

    search_tokens = hist_search.split() if hist_search else []
    if search_tokens:
        matched_created_at = []
        for c_at, grp in hist_df.groupby('created_at', sort=False):
            f_row = grp.iloc[0]
            l_old_c = clean_lot_str(f_row['lot_old'])
            l_new_c = clean_lot_str(f_row['lot_new'])
            reag_c = str(f_row.get('reagent_name', '')).strip()
            overall_ok = all(str(r['status']).upper() == 'PASS' for _, r in grp.iterrows())
            st_label = "đạt pass" if overall_ok else "không đạt trượt fail"
            sids_str = " ".join(clean_lot_str(s) for s in grp['sid'].tolist())

            searchable_blob = (
                f"{f_row['display_date_str']} {f_row['e_date_str']} {f_row['YearMonth']} "
                f"{f_row['test_name']} {f_row['machine']} {f_row.get('machine_id', '')} "
                f"{reag_c} {l_old_c} {l_new_c} {sids_str} {st_label}"
            ).lower()

            if all(tok in searchable_blob for tok in search_tokens):
                matched_created_at.append(c_at)

        hist_df = hist_df[hist_df['created_at'].isin(matched_created_at)]

        if hist_df.empty:
            st.warning(f"Không tìm thấy biên bản đánh giá nào khớp với từ khóa **'{hist_search}'**.", icon=":material/search_off:")
            return
        render_html(f'<div class="search-result-pill">🔍 Tìm thấy <b>{len(matched_created_at)}</b> đợt đánh giá khớp với <b>"{hist_search}"</b></div>')

    total_matched_groups = hist_df['created_at'].nunique()
    year_months = hist_df['YearMonth'].unique().tolist()

    for ym in year_months:
        render_html(f'<div class="month-header-banner">🗓️ Tháng {ym}</div>')
        ym_df = hist_df[hist_df['YearMonth'] == ym]
        grouped_hist = ym_df.groupby('created_at', sort=False)

        for created_at, group in grouped_hist:
            first_row = group.iloc[0]
            t_name = first_row['test_name']
            mach = first_row['machine']
            mach_id = first_row.get('machine_id', '')
            l_old = clean_lot_str(first_row['lot_old'])
            l_new = clean_lot_str(first_row['lot_new'])
            reagent = str(first_row.get('reagent_name', '')).strip()
            if reagent.lower() == 'nan':
                reagent = ''
            e_date_str = first_row['e_date_str']
            display_date_str = first_row['display_date_str']
            unit = first_row['unit']

            overall_pass = all(str(r['status']).upper() == 'PASS' for _, r in group.iterrows())
            status_text = "🟢 Đạt" if overall_pass else "🔴 Không đạt"
            reagent_str = f" • {reagent}" if reagent else ""

            if is_mobile:
                expander_label = f"{'🟢' if overall_pass else '🔴'} {display_date_str} • {t_name} ({mach}) | {l_old} ➔ {l_new}"
            else:
                expander_label = f":material/event: {display_date_str}  |  {t_name} ({mach}){reagent_str}  |  LOT: {l_old} ➔ {l_new}  [{status_text}]"

            auto_expand = bool(search_tokens) and (total_matched_groups <= 3)

            with st.expander(expander_label, expanded=auto_expand):
                reagent_meta = f'<div><strong>Hóa chất:</strong> {reagent}</div>' if reagent else ''
                overall_badge = format_status_badge_html("PASS" if overall_pass else "FAIL")
                render_html(f"""
                <div class="history-meta-bar">
                    <div><strong>Ngày:</strong> <span class="date-pill">{display_date_str}</span></div>
                    <div><strong>XN:</strong> {t_name} ({unit})</div>
                    <div><strong>Máy:</strong> {mach} {f'({mach_id})' if mach_id else ''}</div>
                    {reagent_meta}
                    <div><strong>LOT:</strong> <span class="lot-pill">{l_old} ➔ {l_new}</span></div>
                    <div>{overall_badge}</div>
                </div>
                """)

                if is_mobile:
                    # Giao diện Mobile dạng Thẻ dọc theo từng Level
                    for lv_name, lv_grp in group.groupby('level_name', sort=False):
                        lv_first = lv_grp.iloc[0]
                        lv_badge = format_status_badge_html(lv_first['status'])
                        sample_rows_html = []
                        for _, sr in lv_grp.iterrows():
                            sid_clean = clean_lot_str(sr['sid'])
                            sample_rows_html.append(f"""
                            <div class="m-sample-row">
                                <span><b>{sid_clean}</b></span>
                                <span>{float(sr['val_old']):.3f} ➔ {float(sr['val_new']):.3f}</span>
                                <span>|d|: <b>{float(sr['diff_d']):.3f}</b></span>
                            </div>
                            """)
                        render_html(f"""
                        <div class="m-card">
                            <div class="m-card-head">
                                <div class="m-card-title">🎯 {lv_name} — {float(lv_first['target_conc']):.3f} {unit}</div>
                                {lv_badge}
                            </div>
                            <div class="m-kpi-grid">
                                <div class="m-kpi-card">
                                    <div class="m-kpi-lbl">TB |d|</div>
                                    <div class="m-kpi-val">{float(lv_first['mean_d']):.3f}</div>
                                </div>
                                <div class="m-kpi-card">
                                    <div class="m-kpi-lbl">TB |d| (%)</div>
                                    <div class="m-kpi-val">{float(lv_first['mean_d_pct']):.3f}%</div>
                                </div>
                                <div class="m-kpi-card">
                                    <div class="m-kpi-lbl">RL (%)</div>
                                    <div class="m-kpi-val">{float(lv_first['rl_pct']):.3f}%</div>
                                </div>
                            </div>
                            {"".join(sample_rows_html)}
                        </div>
                        """)
                else:
                    rows_html = []
                    for _, r in group.iterrows():
                        badge_cell = format_status_badge_html(r['status'])
                        sid_clean = clean_lot_str(r['sid'])
                        rows_html.append(f"""
                        <tr>
                            <td class="text-left">{r['level_name']}</td>
                            <td class="text-right">{float(r['target_conc']):.3f}</td>
                            <td class="text-left">{sid_clean}</td>
                            <td class="text-right">{float(r['val_old']):.3f}</td>
                            <td class="text-right">{float(r['val_new']):.3f}</td>
                            <td class="text-right">{float(r['diff_d']):.3f}</td>
                            <td class="text-right">{float(r['mean_d']):.3f}</td>
                            <td class="text-right">{float(r['mean_d_pct']):.3f}</td>
                            <td class="text-right">{float(r['rl_pct']):.3f}</td>
                            <td class="text-center">{badge_cell}</td>
                        </tr>
                        """)

                    render_html(f"""
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
                    """)

                setup_df = all_setups[(all_setups['test_name'] == t_name) & (all_setups['machine'] == mach)]
                docx_bytes_cached = get_cached_report_docx_bytes(
                    e_date_str, t_name, unit, mach, mach_id, reagent, l_old, l_new, setup_df, group
                )
                word_filename = build_report_filename(t_name, mach, reagent, display_date_str)

                if is_mobile:
                    if st.button(
                        ":material/description: Xuất file Word / Up Drive",
                        key=f"btn_export_word_{created_at}",
                        type="primary",
                        use_container_width=True
                    ):
                        if _dialog_decorator is not None:
                            export_word_modal(docx_bytes_cached, word_filename, str(created_at))
                        else:
                            upload_word_to_drive(docx_bytes_cached, word_filename)
                            st.rerun()

                    m_b1, m_b2 = st.columns(2)
                    with m_b1:
                        if f"pdf_ready_{created_at}" not in st.session_state:
                            if st.button(":material/picture_as_pdf: Tạo PDF", key=f"create_pdf_{created_at}", use_container_width=True):
                                with st.spinner("Đang vẽ PDF..."):
                                    pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes_cached)
                                    if pdf_bytes:
                                        st.session_state[f"pdf_bytes_{created_at}"] = pdf_bytes
                                        st.session_state[f"pdf_ready_{created_at}"] = True
                                        queue_popup(
                                            "Tạo file PDF thành công!",
                                            f"Đã khởi tạo xong file PDF cho xét nghiệm <b>{t_name}</b> (LOT: {l_old} ➔ {l_new}).",
                                            "success"
                                        )
                                        st.rerun()
                                    else:
                                        queue_popup(
                                            "Không thể tạo PDF",
                                            "Môi trường Cloud hiện chưa cài đặt LibreOffice. Vui lòng sử dụng nút <b>Xuất file Word</b>.",
                                            "error"
                                        )
                                        st.rerun()
                        else:
                            st.download_button(
                                label=":material/download: Tải PDF",
                                data=st.session_state[f"pdf_bytes_{created_at}"],
                                file_name=f"LotToLot_{t_name}_{l_old}_{l_new}.pdf",
                                mime="application/pdf",
                                key=f"dl_pdf_{created_at}",
                                use_container_width=True
                            )
                    with m_b2:
                        if st.button(":material/delete: Xóa đợt này", key=f"del_hist_{created_at}", use_container_width=True):
                            if _dialog_decorator is not None:
                                confirm_delete_eval_modal(created_at, t_name, display_date_str, l_old, l_new)
                            else:
                                delete_stage2_evaluation(created_at)
                                queue_popup("Đã chuyển vào Thùng rác", f"Đã xóa đợt đánh giá {t_name}.", "warning")
                                st.rerun()
                else:
                    col_word, col_pdf, col_del = st.columns([1, 1, 1])
                    with col_word:
                        if st.button(
                            ":material/description: Xuất file Word",
                            key=f"btn_export_word_{created_at}",
                            use_container_width=True
                        ):
                            if _dialog_decorator is not None:
                                export_word_modal(docx_bytes_cached, word_filename, str(created_at))
                            else:
                                upload_word_to_drive(docx_bytes_cached, word_filename)
                                st.rerun()

                    with col_pdf:
                        if f"pdf_ready_{created_at}" not in st.session_state:
                            if st.button(":material/picture_as_pdf: Tạo file PDF", key=f"create_pdf_{created_at}", use_container_width=True):
                                with st.spinner("Đang vẽ PDF..."):
                                    pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes_cached)
                                    if pdf_bytes:
                                        st.session_state[f"pdf_bytes_{created_at}"] = pdf_bytes
                                        st.session_state[f"pdf_ready_{created_at}"] = True
                                        queue_popup(
                                            "Tạo file PDF thành công!",
                                            f"Đã khởi tạo xong file PDF cho xét nghiệm <b>{t_name}</b> (LOT: {l_old} ➔ {l_new}). Bạn hãy bấm nút <b>Tải PDF xuống</b> để lưu file.",
                                            "success"
                                        )
                                        st.rerun()
                                    else:
                                        queue_popup(
                                            "Không thể tạo PDF",
                                            "Môi trường Cloud hiện chưa cài đặt LibreOffice. Vui lòng sử dụng nút <b>Xuất file Word</b>.",
                                            "error"
                                        )
                                        st.rerun()
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
                        if st.button(":material/delete: Xóa đợt này", key=f"del_hist_{created_at}", use_container_width=True):
                            if _dialog_decorator is not None:
                                confirm_delete_eval_modal(created_at, t_name, display_date_str, l_old, l_new)
                            else:
                                delete_stage2_evaluation(created_at)
                                queue_popup("Đã chuyển vào Thùng rác", f"Đã xóa đợt đánh giá {t_name}.", "warning")
                                st.rerun()


# ==============================================================================
# TRANG 3: TỔNG QUAN LOT (DESKTOP LIST + MOBILE CARDS)
# ==============================================================================
def render_overview_workspace(hist_df_master: pd.DataFrame, is_mobile: bool = False):
    df_hist_dash = hist_df_master.copy()
    if df_hist_dash.empty:
        st.info("Chưa có dữ liệu đánh giá.", icon=":material/info:")
        return

    global_overview_search = st.text_input(
        "Tìm kiếm tổng quan LOT:",
        placeholder="🔍 Tra cứu nhanh LOT đang dùng (Tên XN, hóa chất, máy, LOT)..." if is_mobile else "🔍 Tra cứu nhanh LOT đang dùng: Nhập tên xét nghiệm, hóa chất, dòng máy hoặc số LOT (VD: ALT Cobas, MAU Alinity, WBC Ruby)...",
        key="search_overview_global_top",
        label_visibility="collapsed"
    ).strip().lower()

    df_hist_dash['reagent_name'] = df_hist_dash['reagent_name'].fillna('').astype(str).str.strip()
    df_hist_dash.loc[df_hist_dash['reagent_name'].str.lower() == 'nan', 'reagent_name'] = ''
    df_hist_dash['eval_date_obj'] = pd.to_datetime(df_hist_dash['eval_date'], errors='coerce')
    df_hist_dash = df_hist_dash.dropna(subset=['eval_date_obj'])
    df_hist_dash['date_str'] = df_hist_dash['eval_date_obj'].dt.strftime('%d/%m/%Y')

    df_unique_evals = df_hist_dash.drop_duplicates(subset=['created_at']).copy()
    df_unique_evals = df_unique_evals.sort_values(by=['eval_date_obj', 'created_at'], ascending=[False, False])

    df_sh = df_unique_evals[df_unique_evals['reagent_name'] == '']
    df_hh = df_unique_evals[df_unique_evals['reagent_name'] != '']

    def filter_overview_df(df_in: pd.DataFrame, group_by_col: str, query_str: str) -> pd.DataFrame:
        if not query_str or df_in.empty:
            return df_in
        tokens = query_str.split()
        combined_series = (
            df_in[group_by_col].astype(str) + " " +
            df_in['test_name'].astype(str) + " " +
            df_in['machine'].astype(str) + " " +
            df_in['lot_old'].astype(str) + " " +
            df_in['lot_new'].astype(str) + " " +
            df_in['date_str'].astype(str)
        ).str.lower()
        mask = pd.Series(True, index=df_in.index)
        for tok in tokens:
            mask = mask & combined_series.str.contains(tok, regex=False, na=False)
        return df_in[mask]

    def render_collapsible_overview_section(df_section, section_title, section_icon, group_by_col, search_key, placeholder_text):
        df_after_global = filter_overview_df(df_section, group_by_col, global_overview_search)
        total_items = df_after_global[group_by_col].nunique() if not df_after_global.empty else 0
        total_evals = len(df_after_global)

        auto_expand_section = bool(global_overview_search) and (total_evals > 0)

        with st.expander(f"{section_icon} {section_title}  —  {total_items} danh mục ({total_evals} đợt)", expanded=auto_expand_section):
            if df_section.empty:
                st.info("Chưa có dữ liệu trong phân hệ này.", icon=":material/info:")
                return

            local_search_query = st.text_input(
                "Tìm kiếm trong phân hệ:",
                placeholder=placeholder_text,
                key=search_key,
                label_visibility="collapsed"
            ).strip().lower()

            filtered_df = filter_overview_df(df_after_global, group_by_col, local_search_query)
            active_search = bool(global_overview_search or local_search_query)

            if filtered_df.empty:
                kw_display = global_overview_search or local_search_query
                st.warning(f"Không tìm thấy đợt Lot-to-Lot nào trong phân hệ **{section_title}** khớp với từ khóa **'{kw_display}'**.", icon=":material/search_off:")
                return

            for mach in sorted(filtered_df['machine'].unique()):
                mach_df = filtered_df[filtered_df['machine'] == mach]

                if is_mobile:
                    render_html(f'<div class="month-header-banner">🖨️ Máy: {mach}</div>')
                    cards_html = []
                    for item_name, group in mach_df.groupby(group_by_col, sort=False):
                        latest_eval = group.iloc[0]
                        first_char = str(item_name).strip()[0].upper() if str(item_name).strip() else "T"
                        latest_badge = format_status_badge_html(latest_eval['status'])
                        latest_date = latest_eval['date_str']
                        latest_lot = clean_lot_str(latest_eval['lot_new'])
                        eval_count = len(group)
                        open_attr = "open" if active_search else ""

                        hist_rows_m = []
                        for _, r in group.iterrows():
                            b_html = format_status_badge_html(r['status'])
                            r_lot_old = clean_lot_str(r['lot_old'])
                            r_lot_new = clean_lot_str(r['lot_new'])
                            hist_rows_m.append(f"""
                            <div class="m-sample-row">
                                <span><b>{r['date_str']}</b> • {r['test_name']}</span>
                                <span><code>{r_lot_old} ➔ {r_lot_new}</code></span>
                                <span>{b_html}</span>
                            </div>
                            """)

                        cards_html.append(f"""
                        <details class="m-lot-card" {open_attr}>
                            <summary class="m-lot-summary">
                                <div class="m-card-head">
                                    <div class="m-card-title">
                                        <span class="item-avatar">{first_char}</span>
                                        <span>{item_name}</span>
                                    </div>
                                    {latest_badge}
                                </div>
                                <div class="m-card-meta" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px;">
                                    <span>LOT dùng: <span class="lot-pill">{latest_lot}</span></span>
                                    <span style="color:#64748B;">📅 {latest_date} • <b>{eval_count} đợt ▾</b></span>
                                </div>
                            </summary>
                            <div class="m-lot-details">
                                {"".join(hist_rows_m)}
                            </div>
                        </details>
                        """)
                    render_html("".join(cards_html))
                else:
                    items_html = []
                    for item_name, group in mach_df.groupby(group_by_col, sort=False):
                        latest_eval = group.iloc[0]
                        first_char = str(item_name).strip()[0].upper() if str(item_name).strip() else "T"
                        latest_badge = format_status_badge_html(latest_eval['status'])
                        latest_date = latest_eval['date_str']
                        latest_lot = clean_lot_str(latest_eval['lot_new'])
                        eval_count = len(group)
                        open_attr = "open" if active_search else ""

                        sub_rows = []
                        for _, r in group.iterrows():
                            b_html = format_status_badge_html(r['status'])
                            r_lot_old = clean_lot_str(r['lot_old'])
                            r_lot_new = clean_lot_str(r['lot_new'])
                            sub_rows.append(f"""
                            <tr>
                                <td class="text-left">{r['date_str']}</td>
                                <td class="text-left">{r['test_name']}</td>
                                <td class="text-right">{r_lot_old}</td>
                                <td class="text-right">{r_lot_new}</td>
                                <td class="text-center">{b_html}</td>
                            </tr>
                            """)

                        items_html.append(f"""
                        <details class="overview-item" {open_attr}>
                            <summary class="overview-row">
                                <div class="item-name-cell">
                                    <span class="item-avatar">{first_char}</span>
                                    <span>{item_name}</span>
                                </div>
                                <div class="item-lot-cell">{latest_lot}</div>
                                <div class="item-date-cell">{latest_date}</div>
                                <div>{latest_badge}</div>
                                <div class="item-action-cell">{eval_count} đợt ▾</div>
                            </summary>
                            <div class="overview-subbox">
                                <div class="lab-table-wrapper">
                                    <table class="lab-table">
                                        <thead>
                                            <tr>
                                                <th class="text-left">Ngày thay LOT</th>
                                                <th class="text-left">Xét nghiệm</th>
                                                <th class="text-right">LOT cũ</th>
                                                <th class="text-right">LOT mới</th>
                                                <th class="text-center">Kết quả</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {"".join(sub_rows)}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </details>
                        """)

                    render_html(f"""
                    <div class="overview-list-card">
                        <div class="overview-group-header">
                            <div class="overview-mach-title">🖨️ Máy: {mach}</div>
                            <div>LOT đang dùng</div>
                            <div>Cập nhật</div>
                            <div>Trạng thái</div>
                            <div class="text-right">Lịch sử</div>
                        </div>
                        {"".join(items_html)}
                    </div>
                    """)

    render_collapsible_overview_section(
        df_section=df_sh,
        section_title="Sinh hóa - Miễn dịch",
        section_icon=":material/water_drop:",
        group_by_col="test_name",
        search_key="search_overview_sh",
        placeholder_text="🔍 Lọc riêng Sinh hóa - Miễn dịch (VD: GLUCOSE, Alinity, 70058...)"
    )

    render_collapsible_overview_section(
        df_section=df_hh,
        section_title="Huyết học - HbA1c",
        section_icon=":material/bloodtype:",
        group_by_col="reagent_name",
        search_key="search_overview_hh",
        placeholder_text="🔍 Lọc riêng Huyết học - HbA1c (VD: Eluent, HA 8190V, 6G3101...)"
    )


# ==============================================================================
# TRANG 4: CẤU HÌNH THÔNG SỐ (STAGE 1)
# ==============================================================================
def render_stage1_workspace(all_setups_for_dropdown: pd.DataFrame, is_mobile: bool = False):
    with st.container(border=True):
        st.markdown("#### :material/tune: Thông tin xét nghiệm & thiết bị")
        db_tests = all_setups_for_dropdown["test_name"].unique().tolist() if not all_setups_for_dropdown.empty else []
        test_list = list(dict.fromkeys(
            db_tests + ["CHOLESTEROL", "GLUCOSE", "TSH", "GGT", "AST", "CREATININ", "ALT", "TRIGLYXERIDE", "FT4", "ACID URIC"]
        )) + ["-- Nhập tên mới --"]

        db_machs = all_setups_for_dropdown["machine"].unique().tolist() if not all_setups_for_dropdown.empty else []
        machine_list = list(dict.fromkeys(db_machs + ["Abbott (Alinity)", "Roche (Cobas)"])) + ["-- Nhập máy khác --"]

        if is_mobile:
            selected_test = st.selectbox("Tên xét nghiệm:", test_list, key="s1_sel_test")
            test_name = st.text_input("Nhập tên XN mới:", "HBA1C").strip().upper() if selected_test == "-- Nhập tên mới --" else selected_test
            selected_machine = st.selectbox("Dòng máy:", machine_list, key="s1_sel_mach")
            machine_name = st.text_input("Nhập tên máy mới:", "Sysmex").strip() if selected_machine == "-- Nhập máy khác --" else selected_machine

            saved_df = load_specific_setup(test_name, machine_name)
            existing_mach_id = str(saved_df["machine_id"].iloc[0]) if not saved_df.empty and "machine_id" in saved_df.columns and pd.notna(saved_df["machine_id"].iloc[0]) and str(saved_df["machine_id"].iloc[0]) != 'nan' else ""
            existing_unit = str(saved_df["unit"].iloc[0]) if not saved_df.empty and "unit" in saved_df.columns and pd.notna(saved_df["unit"].iloc[0]) and str(saved_df["unit"].iloc[0]) != 'nan' else "mmol/L"

            mc1, mc2 = st.columns(2)
            machine_id = mc1.text_input("Mã thiết bị:", value=existing_mach_id).strip()
            unit_name = mc2.text_input("Đơn vị đo:", value=existing_unit)
            global_target_power = st.selectbox(
                "Power tối thiểu:",
                options=[0.0, 0.50, 0.60, 0.70, 0.80, 0.85, 0.90, 0.95],
                format_func=lambda x: "Hiển thị tất cả" if x == 0.0 else f">= {x * 100}%",
                index=0
            )
        else:
            col1, col2, col3, col4, col5 = st.columns([1.5, 1.5, 1.2, 1, 1.5])
            with col1:
                selected_test = st.selectbox("Tên xét nghiệm:", test_list, key="s1_sel_test")
                test_name = st.text_input("Nhập tên XN mới:", "HBA1C").strip().upper() if selected_test == "-- Nhập tên mới --" else selected_test
            with col2:
                selected_machine = st.selectbox("Dòng máy:", machine_list, key="s1_sel_mach")
                machine_name = st.text_input("Nhập tên máy mới:", "Sysmex").strip() if selected_machine == "-- Nhập máy khác --" else selected_machine

            saved_df = load_specific_setup(test_name, machine_name)
            existing_mach_id = str(saved_df["machine_id"].iloc[0]) if not saved_df.empty and "machine_id" in saved_df.columns and pd.notna(saved_df["machine_id"].iloc[0]) and str(saved_df["machine_id"].iloc[0]) != 'nan' else ""
            existing_unit = str(saved_df["unit"].iloc[0]) if not saved_df.empty and "unit" in saved_df.columns and pd.notna(saved_df["unit"].iloc[0]) and str(saved_df["unit"].iloc[0]) != 'nan' else "mmol/L"

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
        init_input_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Chọn chạy (Vàng)": st.column_config.CheckboxColumn("Chọn", width="small"),
            "Target concentration": st.column_config.NumberColumn("Target", format="%.3f", width="small"),
            "TEa%": st.column_config.NumberColumn("TEa%", format="%.3f", width="small"),
            "Swrl = CVwrl %": st.column_config.NumberColumn("Swrl%", format="%.3f", width="small"),
            "Sr = CVr %": st.column_config.NumberColumn("Sr%", format="%.3f", width="small"),
        },
        key=f"editor_{test_name}_{machine_name}",
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
                        f"PA {i + 1} (A{num_selected_levels}): {c['n_samples']} mẫu | Power={c['power_achieve']:.3f} | K={c['k_factor']} | RL={c['rejection_limit']:.2f}%"
                        if is_mobile else
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
            queue_popup(
                "Lưu cấu hình thành công!",
                f"Đã lưu thành công thông số Stage 1 và Mã thiết bị cho xét nghiệm <b>{test_name}</b> — Máy: <b>{machine_name}</b>.",
                "success"
            )
            st.rerun()


# ==============================================================================
# TRANG 6: THÙNG RÁC (6. Trash — ĐÃ DỜI SANG LEFT SIDEBAR)
# ==============================================================================
def render_trash_workspace(trash_df_master: pd.DataFrame, is_mobile: bool = False):
    if trash_df_master.empty:
        with st.container(border=True):
            st.info("Thùng rác hiện đang trống.", icon=":material/info:")
        return

    trash_kw = st.text_input(
        "Tìm kiếm trong thùng rác:",
        placeholder="🔍 Lọc trong thùng rác (Ngày, Tên XN, Máy, Hóa chất, số LOT)...",
        key="search_trash_top",
        label_visibility="collapsed"
    ).strip().lower()

    trash_df = trash_df_master.copy()
    grouped_trash = list(trash_df.groupby('created_at', sort=False))
    search_tokens = trash_kw.split() if trash_kw else []

    shown_count = 0
    for c_at, group in grouped_trash:
        first = group.iloc[0]
        dt_parsed = pd.to_datetime(first['eval_date'], errors='coerce')
        disp_date = dt_parsed.strftime('%d/%m/%Y') if pd.notna(dt_parsed) else str(first.get('eval_date', ''))
        t_name = str(first.get('test_name', '')).strip()
        mach = str(first.get('machine', '')).strip()
        l_old_tr = clean_lot_str(first.get('lot_old', ''))
        l_new_tr = clean_lot_str(first.get('lot_new', ''))
        reagent = str(first.get('reagent_name', '')).strip()
        if reagent.lower() == 'nan':
            reagent = ''
        reagent_str = f" ({reagent})" if reagent else ""

        if search_tokens:
            blob = f"{disp_date} {t_name} {mach} {reagent} {l_old_tr} {l_new_tr}".lower()
            if not all(tok in blob for tok in search_tokens):
                continue

        shown_count += 1
        with st.container(border=True):
            if is_mobile:
                render_html(f"""
                <div class="m-card-head" style="margin-bottom:4px;">
                    <div class="m-card-title">🗑️ {t_name}{reagent_str}</div>
                    <span class="date-pill">{disp_date}</span>
                </div>
                <div class="m-card-meta" style="margin-bottom:8px;">
                    🖨️ <b>{mach}</b><br>
                    LOT: <span class="lot-pill">{l_old_tr} ➔ {l_new_tr}</span>
                </div>
                """)
                tc1, tc2 = st.columns(2)
                with tc1:
                    if st.button(":material/restore: Khôi phục", type="primary", key=f"restore_{c_at}", use_container_width=True):
                        restore_stage2_evaluation(c_at)
                        queue_popup(
                            "Khôi phục thành công!",
                            f"Đã khôi phục đợt đánh giá <b>{t_name}{reagent_str}</b> ngày <b>{disp_date}</b> (LOT: {l_old_tr} ➔ {l_new_tr}) về lại mục <b>2. Evaluation History</b>.",
                            "success"
                        )
                        st.rerun()
                with tc2:
                    if st.button(":material/delete_forever: Xóa hẳn", key=f"perm_del_{c_at}", use_container_width=True):
                        permanently_delete_stage2_evaluation(c_at)
                        queue_popup(
                            "Đã xóa vĩnh viễn!",
                            f"Đã xóa vĩnh viễn đợt đánh giá <b>{t_name}{reagent_str}</b> ngày <b>{disp_date}</b> khỏi Thùng rác.",
                            "error"
                        )
                        st.rerun()
            else:
                c_trash1, c_restore, c_delete = st.columns([3.2, 1, 1], vertical_alignment="center")
                with c_trash1:
                    render_html(f"""
                    <div style="font-size:0.84rem; color:#0A2540;">
                        <span class="date-pill">{disp_date}</span> &nbsp;|&nbsp;
                        <strong>{t_name}</strong>{reagent_str} &nbsp;|&nbsp;
                        <span style="color:#475569;">{mach}</span> &nbsp;|&nbsp;
                        LOT: <span class="lot-pill">{l_old_tr} ➔ {l_new_tr}</span>
                    </div>
                    """)
                with c_restore:
                    if st.button(":material/restore: Khôi phục", type="primary", key=f"restore_{c_at}", use_container_width=True):
                        restore_stage2_evaluation(c_at)
                        queue_popup(
                            "Khôi phục thành công!",
                            f"Đã khôi phục đợt đánh giá <b>{t_name}{reagent_str}</b> ngày <b>{disp_date}</b> (LOT: {l_old_tr} ➔ {l_new_tr}) về lại mục <b>2. Evaluation History</b>.",
                            "success"
                        )
                        st.rerun()
                with c_delete:
                    if st.button(":material/delete_forever: Xóa hẳn", key=f"perm_del_{c_at}", use_container_width=True):
                        permanently_delete_stage2_evaluation(c_at)
                        queue_popup(
                            "Đã xóa vĩnh viễn!",
                            f"Đã xóa vĩnh viễn đợt đánh giá <b>{t_name}{reagent_str}</b> ngày <b>{disp_date}</b> khỏi Thùng rác.",
                            "error"
                        )
                        st.rerun()

    if shown_count == 0:
        st.warning(f"Không tìm thấy bản ghi nào trong thùng rác khớp với từ khóa **'{trash_kw}'**.", icon=":material/search_off:")


# ==============================================================================
# KHỞI TẠO CHẾ ĐỘ GIAO DIỆN (AUTO DETECT + MANUAL OVERRIDE), DỮ LIỆU & SIDEBAR
# ==============================================================================
if "ui_mode_override" not in st.session_state:
    st.session_state["ui_mode_override"] = "auto"

if st.session_state["ui_mode_override"] == "mobile":
    IS_MOBILE = True
elif st.session_state["ui_mode_override"] == "desktop":
    IS_MOBILE = False
else:
    IS_MOBILE = _detect_mobile_from_request()

inject_custom_css(is_mobile=IS_MOBILE)
init_db()
render_pending_popup()

all_setups = load_all_setups()
hist_df_master = load_stage2_history()
trash_df_master = load_trash_history()
trash_count = int(trash_df_master["created_at"].nunique()) if (not trash_df_master.empty and "created_at" in trash_df_master.columns) else 0

MENU_ITEMS = [
    ("1. Lot Evaluation", ":material/science:", "🧪 Đánh giá LOT"),
    ("2. Evaluation History", ":material/history:", "🕒 Lịch sử đánh giá"),
    ("3. Lot Overview", ":material/dashboard:", "📊 Tổng quan LOT"),
    ("4. Parameter Setup", ":material/tune:", "⚙️ Cấu hình Stage 1"),
    ("5. Test Directory", ":material/folder_managed:", "📂 Danh mục XN"),
    ("6. Trash", ":material/delete_outline:", f"🗑️ Thùng rác ({trash_count})" if trash_count > 0 else "🗑️ Thùng rác"),
]

valid_menu_keys = [m[0] for m in MENU_ITEMS]
if "active_menu_page" not in st.session_state or st.session_state["active_menu_page"] not in valid_menu_keys:
    st.session_state["active_menu_page"] = "1. Lot Evaluation"


def _switch_menu_page(target_page: str):
    st.session_state["active_menu_page"] = target_page


with st.sidebar:
    render_html("""
    <div class="sidebar-brand">
        <div class="sidebar-brand-title">🔬 CLSI EP26-Ed2</div>
        <div class="sidebar-brand-sub">Lot-to-Lot Evaluation System</div>
    </div>
    <div class="sidebar-section-label">MAIN MENU</div>
    """)

    for item_key, item_icon, _mob_lbl in MENU_ITEMS:
        is_active = (st.session_state["active_menu_page"] == item_key)
        disp_label = f"{item_key} ({trash_count})" if (item_key == "6. Trash" and trash_count > 0) else item_key
        st.button(
            f"{item_icon}  {disp_label}",
            key=f"nav_btn_{item_key}",
            type="primary" if is_active else "secondary",
            use_container_width=True,
            on_click=_switch_menu_page,
            args=(item_key,)
        )

    st.divider()
    render_html('<div class="sidebar-section-label">GIAO DIỆN HIỂN THỊ</div>')
    mode_choice_map = {
        "auto": f"🔄 Tự động ({'Mobile' if _detect_mobile_from_request() else 'Desktop'})",
        "desktop": "💻 Giao diện Máy tính",
        "mobile": "📱 Giao diện Điện thoại",
    }
    mode_keys = ["auto", "desktop", "mobile"]
    cur_mode_idx = mode_keys.index(st.session_state["ui_mode_override"]) if st.session_state["ui_mode_override"] in mode_keys else 0
    picked_mode = st.selectbox(
        "Chế độ hiển thị:",
        options=mode_keys,
        index=cur_mode_idx,
        format_func=lambda k: mode_choice_map[k],
        key="_w_ui_mode_sel",
        label_visibility="collapsed",
    )
    if picked_mode != st.session_state["ui_mode_override"]:
        st.session_state["ui_mode_override"] = picked_mode
        st.rerun()

# Thanh điều hướng nhanh đầu trang khi đang mở trên giao diện Điện thoại
if IS_MOBILE:
    render_html("""
    <div class="mobile-topbar">
        <div class="mobile-topbar-title">🔬 CLSI EP26 — Lot-to-Lot</div>
        <span class="mobile-topbar-badge">📱 Mobile</span>
    </div>
    """)
    mobile_label_map = {m[0]: m[2] for m in MENU_ITEMS}
    mob_nav_sel = persistent_selectbox(
        st,
        label="Chọn chức năng:",
        options=valid_menu_keys,
        persist_key="active_menu_page",
        default_value="1. Lot Evaluation",
        format_func=lambda k: mobile_label_map.get(k, k),
        label_visibility="collapsed",
    )
    page = mob_nav_sel
else:
    page = st.session_state["active_menu_page"]


# ==============================================================================
# ĐIỀU HƯỚNG 6 TRANG
# ==============================================================================
if page == "1. Lot Evaluation":
    render_page_banner(
        "Đánh giá Mẫu thực tế (Stage 2)",
        "Nhập kết quả chạy mẫu song song trên lô hóa chất cũ và mới để kiểm tra độ tương đồng theo tiêu chuẩn CLSI EP26-Ed2.",
        is_mobile=IS_MOBILE,
    )
    active_setups = all_setups[all_setups["is_selected"] == 1] if not all_setups.empty else pd.DataFrame()
    if active_setups.empty:
        st.warning("Chưa có cấu hình xét nghiệm nào. Vui lòng chọn mục **4. Parameter Setup** để thiết lập thông số ban đầu!", icon=":material/warning:")
    else:
        render_stage2_workspace(active_setups, is_mobile=IS_MOBILE)

elif page == "2. Evaluation History":
    render_page_banner(
        "Lịch sử Đánh giá Mẫu thực tế",
        "Tra cứu hồ sơ đánh giá Lot-to-Lot theo từng tháng và xuất biên bản báo cáo Word/PDF.",
        is_mobile=IS_MOBILE,
    )
    render_history_workspace(hist_df_master, all_setups, is_mobile=IS_MOBILE)

elif page == "3. Lot Overview":
    render_page_banner(
        "Tổng quan LOT (Lot Overview)",
        "Tra cứu ngay số LOT mới nhất đang chạy trên từng máy, hoặc bấm vào từng phân hệ để xem danh sách.",
        is_mobile=IS_MOBILE,
    )
    render_overview_workspace(hist_df_master, is_mobile=IS_MOBILE)

elif page == "4. Parameter Setup":
    render_page_banner(
        "Cấu hình Thông số (Stage 1)",
        "Thiết lập nồng độ đích, TEa%, Swrl% và Sr% để hệ thống tra bảng tính số lượng mẫu (n), lực thống kê (Power) và giới hạn từ chối (RL).",
        is_mobile=IS_MOBILE,
    )
    render_stage1_workspace(all_setups, is_mobile=IS_MOBILE)

elif page == "5. Test Directory":
    def _update_setup_active_status(id_to_status_map: dict):
        """Cập nhật trạng thái Active (1) / Inactive (0) trực tiếp lên Google Sheets (RAW mode)."""
        from connection import get_spreadsheet
        from database import STAGE1_HEADERS
        sh = get_spreadsheet()
        ws = sh.worksheet("stage1_setup")
        rows = ws.get_all_values()
        if not rows or len(rows) <= 1:
            return
        headers = rows[0]
        data = []
        for r in rows[1:]:
            if not any(str(cell).strip() for cell in r):
                continue
            row_dict = {h: (r[i] if i < len(r) else "") for i, h in enumerate(headers) if h}
            r_id = str(row_dict.get("id", "")).strip()
            if r_id in id_to_status_map:
                row_dict["is_selected"] = 1 if id_to_status_map[r_id] else 0
            data.append(row_dict)
        ws.clear()
        if data:
            df_to_save = pd.DataFrame(data).reindex(columns=STAGE1_HEADERS).fillna("")
            ws.update([STAGE1_HEADERS] + df_to_save.values.tolist(), value_input_option="RAW")
        else:
            ws.append_row(STAGE1_HEADERS, value_input_option="RAW")
        load_all_setups.clear()

    render_page_banner(
        "Danh mục Xét nghiệm đã lưu (Test Directory)",
        "Quản lý thông số cấu hình Stage 1: Dòng ở trạng thái Active sẽ được sử dụng ở mục 1. Lot Evaluation, dòng Inactive sẽ tạm ngưng sử dụng.",
        is_mobile=IS_MOBILE,
    )

    with st.expander(":material/upload_file: Khôi phục / Import dữ liệu từ file (.csv, .xlsx)", expanded=False):
        uploaded_file = st.file_uploader("Kéo thả file dữ liệu đã xuất từ hệ thống vào đây", type=["csv", "xlsx"])
        if uploaded_file is not None:
            try:
                df_import = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
                if "test_name" in df_import.columns and "target_conc" in df_import.columns:
                    st.info(f"Phát hiện **{len(df_import)}** dòng cấu hình hợp lệ.", icon=":material/info:")
                    if st.button(":material/check_circle: Bắt đầu Import Dữ liệu", type="primary", use_container_width=True):
                        imported_cnt = import_stage1_from_dataframe(df_import)
                        queue_popup(
                            "Import dữ liệu thành công!",
                            f"Đã khôi phục và cập nhật thành công <b>{imported_cnt}</b> dòng cấu hình Stage 1 vào cơ sở dữ liệu.",
                            "success"
                        )
                        st.rerun()
                else:
                    st.error("File không đúng định dạng!", icon=":material/error:")
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}", icon=":material/error:")

    if not all_setups.empty:
        with st.container(border=True):
            if IS_MOBILE:
                st.markdown("#### :material/table_chart: Stage 1 Parameter Directory")
                dir_kw = st.text_input(
                    "Tìm kiếm cấu hình:",
                    placeholder="🔍 Tìm theo tên XN hoặc máy...",
                    key="dir_search_m",
                    label_visibility="collapsed",
                ).strip().lower()
                df_dir_view = all_setups.copy()
                if dir_kw:
                    df_dir_view = df_dir_view[
                        df_dir_view["test_name"].astype(str).str.lower().str.contains(dir_kw, na=False)
                        | df_dir_view["machine"].astype(str).str.lower().str.contains(dir_kw, na=False)
                    ]
                cards_dir = []
                for _, r in df_dir_view.iterrows():
                    t_name = str(r.get("test_name", "")).strip()
                    mach = str(r.get("machine", "")).strip()
                    unit_val = str(r.get("unit", "")).strip()
                    target_val = f"{float(r.get('target_conc', 0)):g}"
                    tea_val = f"{float(r.get('tea_pct', 0)):g}%"
                    rl_val = f"{float(r.get('rejection_limit', 0)):g}%"
                    n_samp = int(float(r.get("n_samples", 0)))
                    is_sel = int(float(r.get("is_selected", 1))) == 1
                    st_badge = '<span class="status-badge badge-pass">✓ Active</span>' if is_sel else '<span class="status-badge badge-warn">⏸ Inactive</span>'
                    cards_dir.append(f"""
                    <div class="m-card" style="{'opacity:0.6;' if not is_sel else ''}">
                        <div class="m-card-head">
                            <div class="m-card-title">{t_name} • {target_val} {unit_val}</div>
                            {st_badge}
                        </div>
                        <div class="m-card-meta">
                            🖨️ <b>{mach}</b> &nbsp;|&nbsp; TEa: <b>{tea_val}</b> &nbsp;|&nbsp; RL: <b style="color:#1B5E20;">{rl_val}</b> &nbsp;|&nbsp; n = <b>{n_samp}</b>
                        </div>
                    </div>
                    """)
                render_html("".join(cards_dir) if cards_dir else '<div class="search-result-pill">Không tìm thấy cấu hình phù hợp.</div>')
            else:
                c_title, c_export = st.columns([3.5, 1])
                with c_title:
                    st.markdown("#### :material/table_chart: Stage 1 Parameter Directory")
                with c_export:
                    csv_backup = all_setups.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        label=":material/download: Export CSV",
                        data=csv_backup,
                        file_name="Stage1_Parameter_Setup_Backup.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

                dir_rows_html = []
                for _, r in all_setups.iterrows():
                    t_name = str(r.get("test_name", "")).strip()
                    mach = str(r.get("machine", "")).strip()
                    mach_id = clean_lot_str(r.get("machine_id", ""))
                    unit_val = str(r.get("unit", "")).strip()
                    target_val = f"{float(r.get('target_conc', 0)):g}"

                    tea_val = f"{float(r.get('tea_pct', 0)):g}%"
                    swrl_val = f"{float(r.get('swrl_pct', 0)):g}%"
                    sr_val = f"{float(r.get('sr_pct', 0)):g}%"
                    cd_swrl_val = f"{float(r.get('cd_swrl', 0)):g}"
                    sr_swrl_val = f"{float(r.get('sr_swrl', 0)):g}"
                    k_val = f"{float(r.get('k_factor', 0)):g}"
                    power_val = f"{float(r.get('power_achieve', 0)):.3f}".rstrip("0").rstrip(".")
                    rl_val = f"{float(r.get('rejection_limit', 0)):g}%"

                    n_samp = int(float(r.get("n_samples", 0)))
                    is_sel = int(float(r.get("is_selected", 1))) == 1

                    if is_sel:
                        status_badge = '<span class="status-badge badge-pass">✓ Active</span>'
                        row_style = ""
                    else:
                        status_badge = '<span class="status-badge badge-warn">⏸ Inactive</span>'
                        row_style = 'style="opacity: 0.55; background-color: #F8FAFC;"'

                    cell_bold_center = 'class="text-center" style="font-weight:700; color:#0A2540; vertical-align:middle;"'
                    cell_rl_highlight = 'class="text-center" style="font-weight:800; color:#1B5E20; background-color:#F1F8F3; vertical-align:middle;"'

                    dir_rows_html.append(f"""
                    <tr {row_style}>
                        <td class="text-center" style="font-weight:700; color:#0A2540; vertical-align:middle;">{t_name}</td>
                        <td class="text-center" style="vertical-align:middle;">{mach}</td>
                        <td class="text-center" style="vertical-align:middle; color:#475569;">{mach_id}</td>
                        <td class="text-center" style="font-weight:600; vertical-align:middle;">{target_val}</td>
                        <td class="text-center" style="vertical-align:middle; color:#475569;">{unit_val}</td>
                        <td {cell_bold_center}>{tea_val}</td>
                        <td {cell_bold_center}>{swrl_val}</td>
                        <td {cell_bold_center}>{sr_val}</td>
                        <td {cell_bold_center}>{cd_swrl_val}</td>
                        <td {cell_bold_center}>{sr_swrl_val}</td>
                        <td {cell_bold_center}>{k_val}</td>
                        <td {cell_bold_center}>{power_val}</td>
                        <td {cell_rl_highlight}>{rl_val}</td>
                        <td class="text-center" style="font-weight:600; vertical-align:middle;">{n_samp}</td>
                        <td class="text-center" style="vertical-align:middle;">{status_badge}</td>
                    </tr>
                    """)

                th_style = 'class="text-center" style="font-weight:700; color:#0A2540; vertical-align:middle;"'
                th_key_style = 'class="text-center" style="font-weight:800; color:#1B5E20; background-color:#DCFCE7; vertical-align:middle;"'

                render_html(f"""
                <div class="lab-table-wrapper" style="max-height: 460px;">
                    <table class="lab-table">
                        <thead>
                            <tr>
                                <th {th_style}>Test Name</th>
                                <th {th_style}>Analyzer</th>
                                <th {th_style}>Machine ID</th>
                                <th {th_style}>Target Conc.</th>
                                <th {th_style}>Unit</th>
                                <th {th_key_style}>TEa%</th>
                                <th {th_key_style}>Swrl</th>
                                <th {th_key_style}>Sr</th>
                                <th {th_key_style}>CD/Swrl</th>
                                <th {th_key_style}>Sr/Swrl</th>
                                <th {th_key_style}>K</th>
                                <th {th_key_style}>Power Achieve</th>
                                <th {th_key_style}>Rejection Limit</th>
                                <th {th_style}>Samples (n)</th>
                                <th {th_style}>Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            {"".join(dir_rows_html)}
                        </tbody>
                    </table>
                </div>
                """)

        col_toggle_box, col_del_box = (st.container(), st.container()) if IS_MOBILE else st.columns(2)

        with col_toggle_box:
            with st.container(border=True):
                st.markdown("#### :material/toggle_on: Bật / Tắt trạng thái sử dụng (Active / Inactive)")
                t_col1, t_col2 = st.columns(2)
                with t_col1:
                    tg_test = st.selectbox("Chọn Xét nghiệm:", all_setups["test_name"].unique(), key="tg_test_sel")
                with t_col2:
                    tg_mach_list = all_setups[all_setups["test_name"] == tg_test]["machine"].unique()
                    tg_mach = st.selectbox("Chọn Máy:", tg_mach_list, key="tg_mach_sel")

                sub_rows_tg = all_setups[
                    (all_setups["test_name"] == tg_test) & (all_setups["machine"] == tg_mach)
                ].reset_index(drop=True)

                level_options = ["Tất cả các mức nồng độ"]
                level_id_map = {"Tất cả các mức nồng độ": sub_rows_tg["id"].astype(str).tolist()}

                for _, sr in sub_rows_tg.iterrows():
                    st_txt = "Active" if int(float(sr.get("is_selected", 1))) == 1 else "Inactive"
                    lbl = f"Mức {float(sr['target_conc']):g} {sr['unit']} (Đang: {st_txt})"
                    level_options.append(lbl)
                    level_id_map[lbl] = [str(sr["id"])]

                tg_level_choice = st.selectbox("Phạm vi áp dụng:", level_options, key="tg_level_sel")
                target_ids = level_id_map[tg_level_choice]

                btn_act, btn_inact = st.columns(2)
                with btn_act:
                    if st.button(":material/check_circle: Bật Active", type="primary", use_container_width=True, key="btn_set_active"):
                        _update_setup_active_status({rid: True for rid in target_ids})
                        queue_popup(
                            "Đã bật trạng thái Active!",
                            f"Cấu hình <b>{tg_test}</b> ({tg_mach}) — <i>{tg_level_choice}</i> đã được kích hoạt và sẵn sàng sử dụng ở mục <b>1. Lot Evaluation</b>.",
                            "success"
                        )
                        st.rerun()
                with btn_inact:
                    if st.button(":material/pause_circle: Tắt Inactive", use_container_width=True, key="btn_set_inactive"):
                        _update_setup_active_status({rid: False for rid in target_ids})
                        queue_popup(
                            "Đã chuyển sang Inactive!",
                            f"Cấu hình <b>{tg_test}</b> ({tg_mach}) — <i>{tg_level_choice}</i> đã tạm ngưng sử dụng và được ẩn khỏi mục <b>1. Lot Evaluation</b>.",
                            "warning"
                        )
                        st.rerun()

        with col_del_box:
            with st.container(border=True):
                st.markdown("#### :material/delete_sweep: Xóa vĩnh viễn cấu hình khỏi Database")
                d_col1, d_col2 = st.columns(2)
                with d_col1:
                    del_test = st.selectbox("Chọn XN cần xóa:", all_setups["test_name"].unique(), key="d_t1")
                with d_col2:
                    del_mach_list = all_setups[all_setups["test_name"] == del_test]["machine"].unique()
                    del_mach = st.selectbox("Chọn Máy cần xóa:", del_mach_list, key="d_m1")

                st.markdown("<div style='font-size:0.78rem; color:#64748B; margin-bottom:10px;'>Lưu ý: Nếu chỉ muốn ẩn tạm thời dòng máy cũ, bạn nên dùng chức năng <b>Tắt Inactive</b> thay vì xóa.</div>", unsafe_allow_html=True)

                if st.button(":material/delete: Xóa vĩnh viễn cấu hình này", use_container_width=True, key="btn_del_setup_perm"):
                    if _dialog_decorator is not None:
                        confirm_delete_setup_modal(del_test, del_mach)
                    else:
                        delete_specific_setup(del_test, del_mach)
                        queue_popup("Đã xóa cấu hình", f"Đã xóa cấu hình {del_test} - {del_mach}.", "error")
                        st.rerun()
    else:
        st.info("Chưa có cấu hình xét nghiệm nào được lưu.", icon=":material/info:")

elif page == "6. Trash":
    render_page_banner(
        f"Thùng rác ({trash_count} đợt đã xóa)",
        "Khôi phục các biên bản đánh giá Lot-to-Lot đã xóa về lại Lịch sử hoặc xóa vĩnh viễn khỏi cơ sở dữ liệu.",
        is_mobile=IS_MOBILE,
    )
    render_trash_workspace(trash_df_master, is_mobile=IS_MOBILE)
