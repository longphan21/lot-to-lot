import io
import streamlit as st
import pandas as pd
from clsi_engine import evaluate_stage1_row
from database import (
    init_db,
    save_test_setup,
    load_all_setups,
    load_specific_setup,
    delete_specific_setup
)

st.set_page_config(
    page_title="CLSI EP26 - Thiết lập thông số Lot-to-Lot (Stage 1)",
    layout="wide"
)

init_db()

st.title("Hệ thống Quản lý Đánh giá Lot-to-Lot (CLSI EP26-Ed2)")
st.caption("Nhánh 1 (Stage 1 Setup): Thiết lập thông số kỹ thuật, tra cứu mẫu & giới hạn từ chối (Rejection Limit)")

# Dữ liệu mặc định từ bảng Excel GLUCOSE của bạn để dễ kiểm tra ngay lập tức
DEFAULT_SAMPLES = {
    ("GLUCOSE", "Abbott (Alinity)"): pd.DataFrame([
        {"Chọn chạy (Vàng)": False, "Target concentration": 0.39,  "TEa%": 9.300, "Swrl = CVwrl %": 1.8, "Sr = CVr %": 1.8, "k cố định (0 = Tự động)": 0.8},
        {"Chọn chạy (Vàng)": True,  "Target concentration": 3.03,  "TEa%": 9.300, "Swrl = CVwrl %": 1.1, "Sr = CVr %": 0.9, "k cố định (0 = Tự động)": 0.7},
        {"Chọn chạy (Vàng)": True,  "Target concentration": 5.88,  "TEa%": 9.300, "Swrl = CVwrl %": 0.9, "Sr = CVr %": 0.8, "k cố định (0 = Tự động)": 0.7},
        {"Chọn chạy (Vàng)": False, "Target concentration": 40.38, "TEa%": 9.300, "Swrl = CVwrl %": 0.8, "Sr = CVr %": 0.8, "k cố định (0 = Tự động)": 0.8},
    ]),
    ("GLUCOSE", "Roche (Cobas)"): pd.DataFrame([
        {"Chọn chạy (Vàng)": False, "Target concentration": 0.188, "TEa%": 9.300, "Swrl = CVwrl %": 10.0, "Sr = CVr %": 9.2, "k cố định (0 = Tự động)": 0.8},
        {"Chọn chạy (Vàng)": True,  "Target concentration": 3.57,  "TEa%": 9.300, "Swrl = CVwrl %": 0.6,  "Sr = CVr %": 0.5, "k cố định (0 = Tự động)": 0.7},
        {"Chọn chạy (Vàng)": True,  "Target concentration": 5.46,  "TEa%": 9.300, "Swrl = CVwrl %": 0.5,  "Sr = CVr %": 0.4, "k cố định (0 = Tự động)": 0.7},
        {"Chọn chạy (Vàng)": False, "Target concentration": 19.60, "TEa%": 9.300, "Swrl = CVwrl %": 0.7,  "Sr = CVr %": 0.6, "k cố định (0 = Tự động)": 0.8},
    ])
}

tab_setup, tab_database = st.tabs([
    "1. Thiết lập & Tính toán thông số (Stage 1)",
    "2. Danh mục Xét nghiệm đã lưu"
])

with tab_setup:
    col1, col2, col3, col4 = st.columns([2, 2, 1.5, 1.5])
    with col1:
        test_list = [
            "GLUCOSE", "TSH", "GGT", "AST", "CHOLESTEROL",
            "CREATININ", "ALT", "TRIGLYXERIDE", "FT4", "ACID URIC", "-- Nhập tên mới --"
        ]
        selected_test = st.selectbox("Tên xét nghiệm (TÊN XN):", test_list)
        if selected_test == "-- Nhập tên mới --":
            test_name = st.text_input("Nhập tên xét nghiệm mới:", "HBA1C").strip().upper()
        else:
            test_name = selected_test

    with col2:
        machine_list = ["Abbott (Alinity)", "Roche (Cobas)", "-- Nhập máy khác --"]
        selected_machine = st.selectbox("Dòng máy (Machine):", machine_list)
        if selected_machine == "-- Nhập máy khác --":
            machine_name = st.text_input("Nhập tên máy:", "Sysmex").strip()
        else:
            machine_name = selected_machine

    with col3:
        unit_name = st.text_input("Đơn vị đo:", "mmol/L")

    with col4:
        target_power = st.selectbox(
            "Mục tiêu Power (khi k=0):",
            options=[0.80, 0.85, 0.90, 0.95],
            index=2
        )

    st.info(
        "**Nguồn dữ liệu đầu vào:** "
        "**Sr (CVr%)** lấy từ IFU hoặc dữ liệu CVr trong hồ sơ xác nhận phương pháp; "
        "**Swrl (CVwrl%)** lấy từ IFU hoặc dữ liệu IQC theo từng mức nồng độ (tối thiểu 20 điểm dữ liệu)."
    )

    # Tải dữ liệu từ DB nếu đã lưu, nếu chưa có thì lấy dữ liệu mẫu
    saved_df = load_specific_setup(test_name, machine_name)
    if not saved_df.empty:
        init_input_df = pd.DataFrame({
            "Chọn chạy (Vàng)": saved_df["is_selected"].astype(bool),
            "Target concentration": saved_df["target_conc"].astype(float),
            "TEa%": saved_df["tea_pct"].astype(float),
            "Swrl = CVwrl %": saved_df["swrl_pct"].astype(float),
            "Sr = CVr %": saved_df["sr_pct"].astype(float),
            "k cố định (0 = Tự động)": saved_df["k_factor"].astype(float),
        })
    elif (test_name, machine_name) in DEFAULT_SAMPLES:
        init_input_df = DEFAULT_SAMPLES[(test_name, machine_name)].copy()
    else:
        init_input_df = pd.DataFrame([
            {"Chọn chạy (Vàng)": True, "Target concentration": 1.0, "TEa%": 10.0, "Swrl = CVwrl %": 1.5, "Sr = CVr %": 1.2, "k cố định (0 = Tự động)": 0.0},
            {"Chọn chạy (Vàng)": True, "Target concentration": 5.0, "TEa%": 10.0, "Swrl = CVwrl %": 1.2, "Sr = CVr %": 1.0, "k cố định (0 = Tự động)": 0.0},
        ])

    st.subheader("Bước 1: Nhập thông số từng mức nồng độ")
    edited_input = st.data_editor(
        init_input_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Chọn chạy (Vàng)": st.column_config.CheckboxColumn(
                "Chọn chạy (Vàng)",
                help="Tích chọn những mức nồng độ quyết định dùng để đánh giá Lot-to-Lot thực tế"
            ),
            "Target concentration": st.column_config.NumberColumn("Target concentration", format="%.3f"),
            "TEa%": st.column_config.NumberColumn("TEa%", format="%.3f"),
            "Swrl = CVwrl %": st.column_config.NumberColumn("Swrl = CVwrl %", format="%.3f"),
            "Sr = CVr %": st.column_config.NumberColumn("Sr = CVr %", format="%.3f"),
            "k cố định (0 = Tự động)": st.column_config.NumberColumn(
                "k (Nhập 0 để tự tra bảng)",
                min_value=0.0, max_value=0.95, step=0.05, format="%.2f"
            ),
        },
        key=f"editor_{test_name}_{machine_name}"
    )

    # Đếm số mức nồng độ được tích chọn (L = 1, 2, hoặc 3) để chọn Bảng A1, A2, A3
    num_selected_levels = int(edited_input["Chọn chạy (Vàng)"].sum())
    num_levels_for_lookup = max(1, min(3, num_selected_levels if num_selected_levels > 0 else 2))

    # Thực hiện tính toán cho từng dòng
    result_rows = []
    for _, row in edited_input.iterrows():
        if pd.isna(row["Target concentration"]) or pd.isna(row["TEa%"]):
            continue
        k_val = float(row.get("k cố định (0 = Tự động)", 0.0) or 0.0)
        manual_k = k_val if k_val > 0 else None

        calc = evaluate_stage1_row(
            tea_pct=float(row["TEa%"]),
            swrl_pct=float(row["Swrl = CVwrl %"]),
            sr_pct=float(row["Sr = CVr %"]),
            num_levels=num_levels_for_lookup,
            target_power=float(target_power),
            manual_k=manual_k
        )

        result_rows.append({
            "Chọn chạy (Vàng)": bool(row["Chọn chạy (Vàng)"]),
            "TÊN XN": test_name,
            "Machine": machine_name,
            "Target concentration": float(row["Target concentration"]),
            "Đơn vị": unit_name,
            "TEa%": float(row["TEa%"]),
            "Swrl = CVwrl %": float(row["Swrl = CVwrl %"]),
            "Sr = CVr %": float(row["Sr = CVr %"]),
            "CD% (Critical Difference)": calc["cd_pct"],
            "CD/Swrl": calc["cd_swrl"],
            "Sr/Swrl": calc["sr_swrl"],
            "Power achieve": calc["power_achieve"],
            "Number of sample each concentration": calc["n_samples"],
            "Rejection limit": calc["rejection_limit"],
            "k": calc["k_factor"],
            "Ghi chú": calc["note"]
        })

    result_df = pd.DataFrame(result_rows)

    st.subheader(
        f"Bước 2: Bảng kết quả tính toán tự động "
        f"(Đang chọn {num_selected_levels} mức nồng độ quyết định -> Tra theo Bảng A{num_levels_for_lookup})"
    )

    if not result_df.empty:
        # Tô màu vàng cho các dòng được chọn giống hệt bảng Excel của bạn
        def highlight_selected(s):
            if s["Chọn chạy (Vàng)"]:
                return ["background-color: #ffff00; color: #000000; font-weight: bold"] * len(s)
            return [""] * len(s)

        styled_df = result_df.style.apply(highlight_selected, axis=1).format({
            "Target concentration": "{:.3f}",
            "TEa%": "{:.3f}",
            "Swrl = CVwrl %": "{:.3f}",
            "Sr = CVr %": "{:.3f}",
            "CD% (Critical Difference)": "{:.3f}",
            "CD/Swrl": "{:.3f}",
            "Sr/Swrl": "{:.3f}",
            "Power achieve": "{:.3f}",
            "Rejection limit": "{:.2f}",
            "k": "{:.2f}",
        })

        st.dataframe(styled_df, use_container_width=True, hide_index=True)

        if st.button("Lưu cấu hình Xét nghiệm này vào Cơ sở dữ liệu", type="primary"):
            save_test_setup(test_name, machine_name, result_df)
            st.success(f"Đã lưu thành công bộ thông số {test_name} - {machine_name} vào cơ sở dữ liệu!")

with tab_database:
    st.subheader("Toàn bộ thông số Setup đã lưu (Sẵn sàng cho Nhánh 2 chạy đánh giá)")
    all_saved_df = load_all_setups()

    if all_saved_df.empty:
        st.warning("Chưa có xét nghiệm nào được lưu. Hãy bấm 'Lưu cấu hình' ở Tab 1.")
    else:
        only_selected = st.checkbox("Chỉ hiển thị các mức nồng độ được chọn chạy (Dòng tô vàng)", value=False)
        display_db = all_saved_df.copy()
        if only_selected:
            display_db = display_db[display_db["is_selected"] == 1]

        st.dataframe(display_db, use_container_width=True, hide_index=True)

        # Nút xuất ra file Excel tổng hợp
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            display_db.to_excel(writer, index=False, sheet_name="Stage1_Setup")
        st.download_button(
            label="Tải bảng tổng hợp Setup về file Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name="CLSI_EP26_Stage1_Setup.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()
        col_del1, col_del2, col_del3 = st.columns([2, 2, 1])
        with col_del1:
            del_test = st.selectbox("Chọn XN cần xóa:", all_saved_df["test_name"].unique(), key="del_t")
        with col_del2:
            del_mach = st.selectbox("Chọn Máy cần xóa:", all_saved_df["machine"].unique(), key="del_m")
        with col_del3:
            st.write("")
            st.write("")
            if st.button("Xóa cấu hình"):
                delete_specific_setup(del_test, del_mach)
                st.rerun()