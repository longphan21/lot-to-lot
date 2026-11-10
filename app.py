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

# Cấu hình trang
st.set_page_config(
    page_title="CLSI EP26 - Thiết lập Lot-to-Lot",
    layout="wide",
    initial_sidebar_state="collapsed"
)

init_db()

st.title("🔬 Hệ thống Quản lý Đánh giá Lot-to-Lot (CLSI EP26-Ed2)")
st.caption("Nhánh 1 (Stage 1 Setup): Tối ưu hóa cấu hình đánh giá Lot-to-Lot theo khuyến cáo CLSI")
st.markdown("---")

# Mẫu dữ liệu mặc định (Đã loại bỏ cột Power)
DEFAULT_SAMPLES = {
    ("GLUCOSE", "Abbott (Alinity)"): pd.DataFrame([
        {"Chọn chạy (Vàng)": False, "Target concentration": 0.39, "TEa%": 9.300, "Swrl = CVwrl %": 1.8,
         "Sr = CVr %": 1.8},
        {"Chọn chạy (Vàng)": True, "Target concentration": 3.03, "TEa%": 9.300, "Swrl = CVwrl %": 1.1,
         "Sr = CVr %": 0.9},
        {"Chọn chạy (Vàng)": True, "Target concentration": 5.88, "TEa%": 9.300, "Swrl = CVwrl %": 0.9,
         "Sr = CVr %": 0.8},
        {"Chọn chạy (Vàng)": False, "Target concentration": 40.38, "TEa%": 9.300, "Swrl = CVwrl %": 0.8,
         "Sr = CVr %": 0.8},
    ]),
    ("GLUCOSE", "Roche (Cobas)"): pd.DataFrame([
        {"Chọn chạy (Vàng)": False, "Target concentration": 0.188, "TEa%": 9.300, "Swrl = CVwrl %": 10.0,
         "Sr = CVr %": 9.2},
        {"Chọn chạy (Vàng)": True, "Target concentration": 3.57, "TEa%": 9.300, "Swrl = CVwrl %": 0.6,
         "Sr = CVr %": 0.5},
        {"Chọn chạy (Vàng)": True, "Target concentration": 5.46, "TEa%": 9.300, "Swrl = CVwrl %": 0.5,
         "Sr = CVr %": 0.4},
        {"Chọn chạy (Vàng)": False, "Target concentration": 19.60, "TEa%": 9.300, "Swrl = CVwrl %": 0.7,
         "Sr = CVr %": 0.6},
    ])
}

tab_setup, tab_database = st.tabs([
    "⚙️ 1. Thiết lập & Tối ưu thông số (Stage 1)",
    "📂 2. Danh mục Xét nghiệm đã lưu"
])

with tab_setup:
    # -------------------------------------------------------------
    # KHU VỰC THÔNG TIN CHUNG (NẰM TRONG KHUNG VIỀN ĐẸP)
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown("#### 📌 Thông tin Xét nghiệm & Mục tiêu hệ thống")
        col1, col2, col3, col4, col5 = st.columns([2, 2, 1, 1.5, 2])

        with col1:
            test_list = ["GLUCOSE", "TSH", "GGT", "AST", "CHOLESTEROL", "CREATININ", "ALT", "TRIGLYXERIDE", "FT4",
                         "ACID URIC", "-- Nhập tên mới --"]
            selected_test = st.selectbox("Tên xét nghiệm:", test_list)
            test_name = st.text_input("Nhập tên XN mới:",
                                      "HBA1C").strip().upper() if selected_test == "-- Nhập tên mới --" else selected_test

        with col2:
            machine_list = ["Abbott (Alinity)", "Roche (Cobas)", "-- Nhập máy khác --"]
            selected_machine = st.selectbox("Dòng máy:", machine_list)
            machine_name = st.text_input("Nhập tên máy mới:",
                                         "Sysmex").strip() if selected_machine == "-- Nhập máy khác --" else selected_machine

        with col3:
            unit_name = st.text_input("Đơn vị đo:", "mmol/L")

        with col4:
            # Power được chuyển thành cài đặt chung ở đây cho gọn gàng
            global_target_power = st.selectbox("Mục tiêu Power (Tối thiểu):", options=[0.80, 0.85, 0.90, 0.95], index=0)

        with col5:
            sort_option = st.selectbox("Quy tắc đề xuất phương án:",
                                       options=["Ít mẫu nhất ➔ Power cao nhất", "Power cao nhất ➔ Ít mẫu nhất"])
            sort_mode = "min_n_max_power" if "Ít mẫu nhất" in sort_option else "max_power_min_n"

    # Tải dữ liệu từ CSDL hoặc mẫu mặc định
    saved_df = load_specific_setup(test_name, machine_name)
    if not saved_df.empty:
        init_input_df = pd.DataFrame({
            "Chọn chạy (Vàng)": saved_df["is_selected"].astype(bool),
            "Target concentration": saved_df["target_conc"].astype(float),
            "TEa%": saved_df["tea_pct"].astype(float),
            "Swrl = CVwrl %": saved_df["swrl_pct"].astype(float),
            "Sr = CVr %": saved_df["sr_pct"].astype(float),
        })
    elif (test_name, machine_name) in DEFAULT_SAMPLES:
        init_input_df = DEFAULT_SAMPLES[(test_name, machine_name)].copy()
    else:
        init_input_df = pd.DataFrame([
            {"Chọn chạy (Vàng)": True, "Target concentration": 1.0, "TEa%": 10.0, "Swrl = CVwrl %": 1.5,
             "Sr = CVr %": 1.2},
            {"Chọn chạy (Vàng)": True, "Target concentration": 5.0, "TEa%": 10.0, "Swrl = CVwrl %": 1.2,
             "Sr = CVr %": 1.0},
        ])

    # -------------------------------------------------------------
    # BƯỚC 1: BẢNG NHẬP LIỆU GỌN GÀNG HƠN
    # -------------------------------------------------------------
    st.markdown("### 📝 Bước 1: Nhập thông số và Chọn các mức nồng độ thực tế")
    edited_input = st.data_editor(
        init_input_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Chọn chạy (Vàng)": st.column_config.CheckboxColumn("✅ Chọn chạy",
                                                                help="Tích chọn những mức nồng độ quyết định dùng để chạy Lot-to-Lot"),
            "Target concentration": st.column_config.NumberColumn("Target concentration", format="%.3f"),
            "TEa%": st.column_config.NumberColumn("TEa%", format="%.3f"),
            "Swrl = CVwrl %": st.column_config.NumberColumn("Swrl = CVwrl %", format="%.3f"),
            "Sr = CVr %": st.column_config.NumberColumn("Sr = CVr %", format="%.3f"),
        },
        key=f"editor_{test_name}_{machine_name}"
    )

    selected_input = edited_input[edited_input["Chọn chạy (Vàng)"] == True]
    num_selected_levels = len(selected_input)
    num_levels_for_lookup = max(1, min(3, num_selected_levels if num_selected_levels > 0 else 1))

    # -------------------------------------------------------------
    # BƯỚC 2: TRA BẢNG & TỐI ƯU PHƯƠNG ÁN
    # -------------------------------------------------------------
    st.markdown("---")
    st.markdown("### 📊 Bước 2: Tối ưu & Chốt phương án tra cứu")

    if num_selected_levels == 0:
        st.warning(
            "⚠️ Bạn chưa tích chọn mức nồng độ nào ở Bước 1. Vui lòng tích chọn ô '✅ Chọn chạy' để hệ thống tính toán.")
    else:
        # ===== HIỂN THỊ THÔNG BÁO RÕ RÀNG VỀ BẢNG ĐANG ĐƯỢC SỬ DỤNG =====
        if num_levels_for_lookup == 1:
            table_name = "BẢNG A1 (Dành cho 1 mức nồng độ khảo sát)"
            alpha_info = "Xác suất từ chối sai (α) = 5% (0.050) trên 1 mức"
        elif num_levels_for_lookup == 2:
            table_name = "BẢNG A2 (Dành cho 2 mức nồng độ khảo sát)"
            alpha_info = "Xác suất từ chối sai (α) = 2.5% (0.025) cho mỗi mức"
        else:
            table_name = "BẢNG A3 (Dành cho 3 mức nồng độ khảo sát trở lên)"
            alpha_info = "Xác suất từ chối sai (α) = 1.67% (0.0167) cho mỗi mức"

        st.success(f"**🔬 HỆ THỐNG ĐANG TỰ ĐỘNG NỘI SUY THEO {table_name} CỦA CHUẨN CLSI EP26**  \n*({alpha_info})*")
        st.write("")

        all_options_rows = []
        chosen_final_rows = []

        for idx, row in selected_input.iterrows():
            if pd.isna(row["Target concentration"]) or pd.isna(row["TEa%"]):
                continue

            target_c = float(row["Target concentration"])

            candidates = get_candidate_options_for_row(
                tea_pct=float(row["TEa%"]),
                swrl_pct=float(row["Swrl = CVwrl %"]),
                sr_pct=float(row["Sr = CVr %"]),
                num_levels=num_levels_for_lookup,
                desired_power=global_target_power,
                sort_mode=sort_mode
            )

            for rank_idx, cand in enumerate(candidates, start=1):
                all_options_rows.append({
                    "Thứ tự": f"#{rank_idx}",
                    "Target": target_c,
                    "N (Số mẫu)": cand["n_samples"],
                    "Power": cand["power_achieve"],
                    "RL%": cand["rejection_limit"],
                    "k": cand["k_factor"],
                    "CD%": cand["cd_pct"],
                    "CD/Swrl": cand["cd_swrl"],
                    "Sr/Swrl": cand["sr_swrl"],
                    "Ghi chú": cand["note"]
                })

            # Box chọn phương án cho từng mức nồng độ
            with st.container(border=True):
                st.markdown(f"📍 **Mức nồng độ: {target_c:.3f} {unit_name}**")
                if len(candidates) > 1:
                    option_labels = [
                        f"🏆 Hạng #{i + 1}: Chọn chạy {c['n_samples']} mẫu | Đạt Power = {c['power_achieve']:.3f} | Lấy k = {c['k_factor']:.2f} | RL = {c['rejection_limit']:.2f}%"
                        for i, c in enumerate(candidates)
                    ]
                    selected_label = st.selectbox(
                        "Chọn phương án chốt cho mức này:",
                        options=option_labels,
                        index=0,
                        key=f"pick_{test_name}_{machine_name}_{idx}",
                        label_visibility="collapsed"
                    )
                    chosen_idx = option_labels.index(selected_label)
                    picked_cand = candidates[chosen_idx]
                else:
                    picked_cand = candidates[0]
                    st.info(
                        f"💡 Chỉ có 1 phương án thỏa mãn: Cần {picked_cand['n_samples']} mẫu | Power = {picked_cand['power_achieve']:.3f} | k = {picked_cand['k_factor']:.2f} | RL = {picked_cand['rejection_limit']:.2f}%")

            chosen_final_rows.append({
                "Chọn chạy (Vàng)": True,
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
            with st.expander("🔍 Click vào đây để xem toàn bộ bảng phân tích các phương án khả thi"):
                st.dataframe(
                    pd.DataFrame(all_options_rows).style.format(precision=3),
                    use_container_width=True,
                    hide_index=True
                )

        st.write("")
        st.markdown("### 💾 Cấu hình chốt chuẩn bị lưu")
        result_df = pd.DataFrame(chosen_final_rows)


        def highlight_selected(s):
            return ["background-color: #ffff00; color: #000000; font-weight: bold"] * len(s)


        styled_df = result_df.style.apply(highlight_selected, axis=1).format(precision=3)
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

        if st.button("✅ LƯU CẤU HÌNH XÉT NGHIỆM VÀO CƠ SỞ DỮ LIỆU", type="primary", use_container_width=True):
            save_test_setup(test_name, machine_name, result_df)
            st.balloons()
            st.success(
                f"🎉 Đã lưu thành công thông số **{test_name} - {machine_name}** vào cơ sở dữ liệu! Bạn có thể xem tại Tab 2.")

# -------------------------------------------------------------
# TAB 2: QUẢN LÝ DỮ LIỆU ĐÃ LƯU
# -------------------------------------------------------------
with tab_database:
    st.markdown("### 📂 Toàn bộ thông số Setup đã lưu")
    st.caption("Danh sách này sẵn sàng cung cấp thông số cấu hình chuẩn cho Bước chạy đánh giá mẫu thực tế (Stage 2).")

    all_saved_df = load_all_setups()

    if all_saved_df.empty:
        st.warning("Chưa có xét nghiệm nào được lưu. Hãy quay lại Tab 1 để thiết lập.")
    else:
        st.dataframe(all_saved_df, use_container_width=True, hide_index=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            all_saved_df.to_excel(writer, index=False, sheet_name="Stage1_Setup")

        st.download_button(
            label="📥 Xuất bảng tổng hợp về file Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name="CLSI_EP26_Stage1_Setup.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()
        st.markdown("#### 🗑️ Xóa cấu hình xét nghiệm")
        col_del1, col_del2, col_del3 = st.columns([2, 2, 1])
        with col_del1:
            del_test = st.selectbox("Chọn XN cần xóa:", all_saved_df["test_name"].unique(), key="del_t")
        with col_del2:
            del_mach = st.selectbox("Chọn Máy cần xóa:", all_saved_df["machine"].unique(), key="del_m")
        with col_del3:
            st.write("")
            st.write("")
            if st.button("❌ Xóa cấu hình", use_container_width=True):
                delete_specific_setup(del_test, del_mach)
                st.rerun()