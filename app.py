import io
import streamlit as st
import pandas as pd
from clsi_engine import get_candidate_options_for_row
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
st.caption("Nhánh 1 (Stage 1 Setup): Nhập Power mong muốn -> Tự động lọc danh sách tối ưu Số mẫu ít nhất & Power cao nhất")

# Dữ liệu mẫu mặc định
DEFAULT_SAMPLES = {
    ("GLUCOSE", "Abbott (Alinity)"): pd.DataFrame([
        {"Chọn chạy (Vàng)": False, "Target concentration": 0.39,  "TEa%": 9.300, "Swrl = CVwrl %": 1.8, "Sr = CVr %": 1.8, "Power achieve mong muốn": 0.80},
        {"Chọn chạy (Vàng)": True,  "Target concentration": 3.03,  "TEa%": 9.300, "Swrl = CVwrl %": 1.1, "Sr = CVr %": 0.9, "Power achieve mong muốn": 0.85},
        {"Chọn chạy (Vàng)": True,  "Target concentration": 5.88,  "TEa%": 9.300, "Swrl = CVwrl %": 0.9, "Sr = CVr %": 0.8, "Power achieve mong muốn": 0.85},
        {"Chọn chạy (Vàng)": False, "Target concentration": 40.38, "TEa%": 9.300, "Swrl = CVwrl %": 0.8, "Sr = CVr %": 0.8, "Power achieve mong muốn": 0.80},
    ]),
    ("GLUCOSE", "Roche (Cobas)"): pd.DataFrame([
        {"Chọn chạy (Vàng)": False, "Target concentration": 0.188, "TEa%": 9.300, "Swrl = CVwrl %": 10.0, "Sr = CVr %": 9.2, "Power achieve mong muốn": 0.80},
        {"Chọn chạy (Vàng)": True,  "Target concentration": 3.57,  "TEa%": 9.300, "Swrl = CVwrl %": 0.6,  "Sr = CVr %": 0.5, "Power achieve mong muốn": 0.85},
        {"Chọn chạy (Vàng)": True,  "Target concentration": 5.46,  "TEa%": 9.300, "Swrl = CVwrl %": 0.5,  "Sr = CVr %": 0.4, "Power achieve mong muốn": 0.85},
        {"Chọn chạy (Vàng)": False, "Target concentration": 19.60, "TEa%": 9.300, "Swrl = CVwrl %": 0.7,  "Sr = CVr %": 0.6, "Power achieve mong muốn": 0.80},
    ])
}

tab_setup, tab_database = st.tabs([
    "1. Thiết lập & Tối ưu thông số (Stage 1)",
    "2. Danh mục Xét nghiệm đã lưu"
])

with tab_setup:
    col1, col2, col3, col4 = st.columns([2, 2, 1.2, 2.5])
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
        sort_option = st.selectbox(
            "Quy tắc xếp hạng danh sách ở Bước 2:",
            options=[
                "Số mẫu chạy ít nhất -> Power achieve cao nhất",
                "Power achieve cao nhất -> Số mẫu chạy ít nhất"
            ]
        )
        sort_mode = "min_n_max_power" if "Số mẫu chạy ít nhất ->" in sort_option else "max_power_min_n"

    # Tải dữ liệu từ DB nếu đã lưu, nếu chưa có thì lấy dữ liệu mẫu
    saved_df = load_specific_setup(test_name, machine_name)
    if not saved_df.empty:
        init_input_df = pd.DataFrame({
            "Chọn chạy (Vàng)": saved_df["is_selected"].astype(bool),
            "Target concentration": saved_df["target_conc"].astype(float),
            "TEa%": saved_df["tea_pct"].astype(float),
            "Swrl = CVwrl %": saved_df["swrl_pct"].astype(float),
            "Sr = CVr %": saved_df["sr_pct"].astype(float),
            "Power achieve mong muốn": saved_df["power_achieve"].clip(upper=0.95).round(2),
        })
    elif (test_name, machine_name) in DEFAULT_SAMPLES:
        init_input_df = DEFAULT_SAMPLES[(test_name, machine_name)].copy()
    else:
        init_input_df = pd.DataFrame([
            {"Chọn chạy (Vàng)": True, "Target concentration": 0.095, "TEa%": 12.400, "Swrl = CVwrl %": 2.000, "Sr = CVr %": 1.500, "Power achieve mong muốn": 0.80},
            {"Chọn chạy (Vàng)": True, "Target concentration": 6.145, "TEa%": 12.400, "Swrl = CVwrl %": 1.500, "Sr = CVr %": 1.300, "Power achieve mong muốn": 0.80},
        ])

    st.subheader("Bước 1: Nhập thông số và Power achieve mong muốn")
    edited_input = st.data_editor(
        init_input_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Chọn chạy (Vàng)": st.column_config.CheckboxColumn(
                "Chọn chạy (Vàng)",
                help="Tích chọn những mức nồng độ quyết định dùng để đánh giá Lot-to-Lot"
            ),
            "Target concentration": st.column_config.NumberColumn("Target concentration", format="%.3f"),
            "TEa%": st.column_config.NumberColumn("TEa%", format="%.3f"),
            "Swrl = CVwrl %": st.column_config.NumberColumn("Swrl = CVwrl %", format="%.3f"),
            "Sr = CVr %": st.column_config.NumberColumn("Sr = CVr %", format="%.3f"),
            "Power achieve mong muốn": st.column_config.NumberColumn(
                "Power achieve mong muốn",
                help="Nhập mức Power tối thiểu cần đạt (ví dụ: 0.80, 0.85, 0.90)",
                min_value=0.50, max_value=0.99, step=0.05, format="%.2f"
            ),
        },
        key=f"editor_{test_name}_{machine_name}"
    )

    num_selected_levels = int(edited_input["Chọn chạy (Vàng)"].sum())
    num_levels_for_lookup = max(1, min(3, num_selected_levels if num_selected_levels > 0 else 2))

    st.subheader(
        f"Bước 2: Danh sách các phương án phù hợp "
        f"(Đang chọn {num_selected_levels} mức nồng độ quyết định -> Tra theo Bảng A{num_levels_for_lookup})"
    )

    all_options_rows = []
    chosen_final_rows = []

    # Cho phép người dùng xem toàn bộ danh sách phương án phù hợp và chọn phương án chốt cho từng mức nồng độ
    for idx, row in edited_input.iterrows():
        if pd.isna(row["Target concentration"]) or pd.isna(row["TEa%"]):
            continue

        target_c = float(row["Target concentration"])
        desired_p = float(row.get("Power achieve mong muốn", 0.80) or 0.80)
        is_sel = bool(row["Chọn chạy (Vàng)"])

        candidates = get_candidate_options_for_row(
            tea_pct=float(row["TEa%"]),
            swrl_pct=float(row["Swrl = CVwrl %"]),
            sr_pct=float(row["Sr = CVr %"]),
            num_levels=num_levels_for_lookup,
            desired_power=desired_p,
            sort_mode=sort_mode
        )

        for rank_idx, cand in enumerate(candidates, start=1):
            all_options_rows.append({
                "Thứ tự ưu tiên": f"#{rank_idx}" + (" (Tối ưu nhất)" if rank_idx == 1 else ""),
                "Chọn chạy (Vàng)": is_sel,
                "Target concentration": target_c,
                "Power mong muốn": desired_p,
                "Power achieve": cand["power_achieve"],
                "Number of sample each concentration": cand["n_samples"],
                "Rejection limit": cand["rejection_limit"],
                "k": cand["k_factor"],
                "CD% (Critical Difference)": cand["cd_pct"],
                "CD/Swrl": cand["cd_swrl"],
                "Sr/Swrl": cand["sr_swrl"],
                "Ghi chú": cand["note"]
            })

        # Mặc định lấy phương án #1 (Đứng đầu danh sách tối ưu), nhưng cho phép đổi nếu có nhiều phương án
        if is_sel and len(candidates) > 1:
            option_labels = [
                f"Hạng #{i+1}: N = {c['n_samples']} mẫu | Power = {c['power_achieve']:.3f} | k = {c['k_factor']:.2f} | RL = {c['rejection_limit']:.2f}%"
                for i, c in enumerate(candidates)
            ]
            selected_label = st.selectbox(
                f"Chọn phương án chốt cho mức nồng độ {target_c:.3f} ({unit_name}):",
                options=option_labels,
                index=0,
                key=f"pick_{test_name}_{machine_name}_{idx}"
            )
            chosen_idx = option_labels.index(selected_label)
            picked_cand = candidates[chosen_idx]
        else:
            picked_cand = candidates[0]

        chosen_final_rows.append({
            "Chọn chạy (Vàng)": is_sel,
            "TÊN XN": test_name,
            "Machine": machine_name,
            "Target concentration": target_c,
            "Đơn vị": unit_name,
            "TEa%": float(row["TEa%"]),
            "Swrl = CVwrl %": float(row["Swrl = CVwrl %"]),
            "Sr = CVr %": float(row["Sr = CVr %"]),
            "CD% (Critical Difference)": picked_cand["cd_pct"],
            "CD/Swrl": picked_cand["cd_swrl"],
            "Sr/Swrl": picked_cand["sr_swrl"],
            "Power achieve": picked_cand["power_achieve"],
            "Number of sample each concentration": picked_cand["n_samples"],
            "Rejection limit": picked_cand["rejection_limit"],
            "k": picked_cand["k_factor"],
            "Ghi chú": picked_cand["note"]
        })

    if all_options_rows:
        options_df = pd.DataFrame(all_options_rows)
        with st.expander("Xem chi tiết toàn bộ danh sách các phương án phù hợp theo thứ tự ưu tiên", expanded=True):
            st.dataframe(
                options_df.style.format({
                    "Target concentration": "{:.3f}",
                    "Power mong muốn": "{:.2f}",
                    "Power achieve": "{:.3f}",
                    "Rejection limit": "{:.2f}",
                    "k": "{:.2f}",
                    "CD% (Critical Difference)": "{:.3f}",
                    "CD/Swrl": "{:.3f}",
                    "Sr/Swrl": "{:.3f}",
                }),
                use_container_width=True,
                hide_index=True
            )

    result_df = pd.DataFrame(chosen_final_rows)

    st.subheader("Bảng cấu hình chốt để lưu (Tự động lấy phương án tối ưu đã chọn ở trên)")
    if not result_df.empty:
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