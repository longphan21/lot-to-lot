import io
from datetime import date
import streamlit as st
import pandas as pd
from clsi_engine import get_candidate_options_for_row, evaluate_stage2_level
from report_generator import generate_report_docx
from database import (
    init_db, save_test_setup, load_all_setups, load_specific_setup,
    delete_specific_setup, save_stage2_evaluation, load_stage2_history, delete_stage2_evaluation,
    import_stage1_from_dataframe
)

st.set_page_config(page_title="CLSI EP26 - Lot-to-Lot", layout="wide", initial_sidebar_state="collapsed")
init_db()

st.title("🔬 Hệ thống Đánh giá Lot-to-Lot (CLSI EP26-Ed2)")
st.markdown("---")

DEFAULT_SAMPLES = {
    ("CHOLESTEROL", "Alinity C"): pd.DataFrame([
        {"Chọn chạy (Vàng)": True, "Target concentration": 6.37, "TEa%": 12.9, "Swrl = CVwrl %": 0.5,
         "Sr = CVr %": 0.3},
    ])
}

# SẮP XẾP LẠI THỨ TỰ TAB THEO Ý BẠN (SỬ DỤNG NHIỀU LÊN TRƯỚC)
tab_stage2, tab_history, tab_stage1, tab_db_setup = st.tabs([
    "🧪 1. Đánh giá Mẫu thực tế (Stage 2)",
    "🕒 2. Lịch sử Đánh giá",
    "⚙️ 3. Cấu hình Thông số (Stage 1)",
    "📂 4. Danh mục XN đã lưu"
])

# ==============================================================================
# TAB 1: GIAI ĐOẠN 2 - ĐÁNH GIÁ MẪU HẰNG NGÀY
# ==============================================================================
with tab_stage2:
    all_setups = load_all_setups()
    active_setups = all_setups[all_setups["is_selected"] == 1] if not all_setups.empty else pd.DataFrame()

    if active_setups.empty:
        st.warning("⚠️ Chưa có cấu hình xét nghiệm nào. Vui lòng sang **Tab 3** để thiết lập thông số ban đầu!")
    else:
        with st.container(border=True):
            st.markdown("#### 🧾 Thông tin Đợt Đánh giá")
            c1, c2, c3, c4, c5 = st.columns([1.6, 2.0, 1.4, 1.6, 1.6])

            with c1:
                eval_test = st.selectbox("Chọn Xét nghiệm:", active_setups["test_name"].unique().tolist(),
                                         key="s2_test")
            with c2:
                eval_machine = st.selectbox("Dòng máy:", active_setups[active_setups["test_name"] == eval_test][
                    "machine"].unique().tolist(), key="s2_mach")

                # Tự động lấy Mã thiết bị đã lưu từ DB
                eval_machine_id = ""
                mach_df = active_setups[
                    (active_setups["test_name"] == eval_test) & (active_setups["machine"] == eval_machine)]
                if not mach_df.empty and "machine_id" in mach_df.columns and pd.notna(mach_df["machine_id"].iloc[0]):
                    eval_machine_id = str(mach_df["machine_id"].iloc[0])

            with c3:
                eval_date = st.date_input("Ngày thực hiện:", value=date.today(), key="s2_date")
            with c4:
                lot_old = st.text_input("Số LOT cũ:", value="82273UD00", key="s2_lot_old").strip()
            with c5:
                lot_new = st.text_input("Số LOT mới:", value="86280UD00", key="s2_lot_new").strip()

        test_levels_df = mach_df.sort_values("target_conc").reset_index(drop=True)
        unit_s2 = str(test_levels_df["unit"].iloc[0]) if not test_levels_df.empty else "mmol/L"

        st.markdown("### 🧪 Nhập kết quả mẫu (Công thức: Độ chệch d = |LOT mới - LOT cũ|)")
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
                    st.markdown(f"#### 📍 Mức nồng độ: `{target_val:.3f} {unit_s2}`")
                    st.caption(
                        f"Yêu cầu: **{int(lv_row['n_samples'])} mẫu** | Giới hạn từ chối (RL) = **`{rl_pct_db:.3f}%`**")

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
                        st.info("Nhập ít nhất 1 cặp kết quả (> 0) để xem đánh giá.")
                    else:
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Trung bình d", f"{eval_res['mean_d']:.4f}")
                        m2.metric("Trung bình độ chệch (%)", f"{eval_res['mean_d_pct']:.3f}%")
                        m3.metric(f"Rejection Limit", f"{eval_res['rl_pct']:.3f}%")

                        if eval_res["result"] == "PASS":
                            st.success(
                                f"✅ **ĐẠT (PASS)**: `{eval_res['mean_d_pct']:.3f}%` < `{eval_res['rl_pct']:.3f}%`")
                        else:
                            st.error(
                                f"❌ **FAIL (TỪ CHỐI)**: `{eval_res['mean_d_pct']:.3f}%` ≥ `{eval_res['rl_pct']:.3f}%`")

                        for _, d_row in eval_res["details"].iterrows():
                            flat_records_to_save.append({
                                "eval_date": str(eval_date), "test_name": eval_test,
                                "machine": eval_machine, "machine_id": eval_machine_id,
                                "unit": unit_s2, "lot_old": lot_old, "lot_new": lot_new,
                                "level_name": f"Level {level_num}", "target_conc": target_val,
                                "sid": d_row["SID"], "val_old": d_row["LOT_old_val"], "val_new": d_row["LOT_new_val"],
                                "diff_d": d_row["d"], "mean_d": eval_res["mean_d"],
                                "mean_d_pct": eval_res["mean_d_pct"],
                                "rl_pct": eval_res["rl_pct"], "k_factor": k_db, "status": eval_res["result"]
                            })
                        summary_table_rows.append({"Kết quả": eval_res["result"]})

        if summary_table_rows:
            st.markdown("---")
            if st.button("💾 LƯU KẾT QUẢ VÀO LỊCH SỬ", type="primary", use_container_width=True):
                save_stage2_evaluation(flat_records_to_save)
                st.success("Đã lưu biên bản đánh giá! Vui lòng qua Tab 2 để xuất file Word.")

# ==============================================================================
# TAB 2: LỊCH SỬ ĐÁNH GIÁ (XUẤT WORD & XÓA)
# ==============================================================================
with tab_history:
    st.markdown("### 🕒 Lịch sử Đánh giá Mẫu thực tế")
    hist_df = load_stage2_history()

    if hist_df.empty:
        st.info("Chưa có đợt đánh giá nào được lưu.")
    else:
        hist_df['eval_date'] = pd.to_datetime(hist_df['eval_date'])
        hist_df['Year'] = hist_df['eval_date'].dt.year
        years = sorted(hist_df['Year'].unique(), reverse=True)

        for y in years:
            st.markdown(f"## 📅 Năm {y}")
            year_df = hist_df[hist_df['Year'] == y]
            grouped_hist = year_df.groupby('created_at', sort=False)

            for created_at, group in grouped_hist:
                first_row = group.iloc[0]
                t_name = first_row['test_name']
                mach = first_row['machine']
                mach_id = first_row.get('machine_id', '')
                l_old = first_row['lot_old']
                l_new = first_row['lot_new']
                e_date_str = first_row['eval_date'].strftime('%Y-%m-%d')
                unit = first_row['unit']

                with st.expander(f"📌 {e_date_str} | {t_name} ({mach}) | LOT: {l_old} ➔ {l_new}", expanded=False):
                    display_cols = ['level_name', 'target_conc', 'sid', 'val_old', 'val_new', 'diff_d', 'mean_d',
                                    'mean_d_pct', 'rl_pct', 'status']
                    st.dataframe(group[display_cols].style.format(precision=3), use_container_width=True,
                                 hide_index=True)

                    st.write("")
                    col_dl, col_del = st.columns([1, 1])

                    with col_dl:
                        # Chuẩn bị dữ liệu cho file Word
                        setup_df = all_setups[(all_setups['test_name'] == t_name) & (all_setups['machine'] == mach)]
                        docx_file = generate_report_docx(
                            e_date_str, t_name, unit, mach, mach_id, l_old, l_new, setup_df, group
                        )
                        st.download_button(
                            label="📄 XUẤT BÁO CÁO (FILE WORD)",
                            data=docx_file,
                            file_name=f"LotToLot_{t_name}_{l_old}_{l_new}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_word_{created_at}",
                            use_container_width=True
                        )

                    with col_del:
                        if st.button(f"❌ Xóa toàn bộ đợt đánh giá này", key=f"del_hist_{created_at}",
                                     use_container_width=True):
                            delete_stage2_evaluation(created_at)
                            st.rerun()

# ==============================================================================
# TAB 3: GIAI ĐOẠN 1 - SETUP THÔNG SỐ (CẬP NHẬT MÃ THIẾT BỊ)
# ==============================================================================
with tab_stage1:
    all_setups_for_dropdown = load_all_setups()

    with st.container(border=True):
        st.markdown("#### 📌 Thông tin Xét nghiệm & Thiết bị")
        col1, col2, col3, col4, col5 = st.columns([1.5, 1.5, 1.2, 1, 1.5])

        with col1:
            # Tự động lấy danh sách Xét nghiệm đã có trong DB
            db_tests = all_setups_for_dropdown[
                "test_name"].unique().tolist() if not all_setups_for_dropdown.empty else []
            default_tests = ["CHOLESTEROL", "GLUCOSE", "TSH", "GGT", "AST", "CREATININ", "ALT", "TRIGLYXERIDE", "FT4",
                             "ACID URIC"]
            test_list = list(dict.fromkeys(db_tests + default_tests)) + ["-- Nhập tên mới --"]

            selected_test = st.selectbox("Tên xét nghiệm:", test_list)
            test_name = st.text_input("Nhập tên XN mới:",
                                      "HBA1C").strip().upper() if selected_test == "-- Nhập tên mới --" else selected_test

        with col2:
            # Tự động lấy danh sách Máy đã có trong DB
            db_machs = all_setups_for_dropdown["machine"].unique().tolist() if not all_setups_for_dropdown.empty else []
            default_machs = ["Abbott (Alinity)", "Roche (Cobas)"]
            machine_list = list(dict.fromkeys(db_machs + default_machs)) + ["-- Nhập máy khác --"]

            selected_machine = st.selectbox("Dòng máy:", machine_list)
            machine_name = st.text_input("Nhập tên máy mới:",
                                         "Sysmex").strip() if selected_machine == "-- Nhập máy khác --" else selected_machine

        # TRUY XUẤT NGAY DATABASE VỚI TEST VÀ MACHINE ĐÃ CHỌN ĐỂ MATCHING
        saved_df = load_specific_setup(test_name, machine_name)

        existing_mach_id = ""
        existing_unit = "mmol/L"
        if not saved_df.empty:
            if "machine_id" in saved_df.columns and pd.notna(saved_df["machine_id"].iloc[0]):
                val = str(saved_df["machine_id"].iloc[0])
                if val != 'nan':
                    existing_mach_id = val
            if "unit" in saved_df.columns and pd.notna(saved_df["unit"].iloc[0]):
                val = str(saved_df["unit"].iloc[0])
                if val != 'nan':
                    existing_unit = val

        with col3:
            machine_id = st.text_input("Mã thiết bị:", value=existing_mach_id).strip()

        with col4:
            unit_name = st.text_input("Đơn vị đo:", value=existing_unit)

        with col5:
            global_target_power = st.selectbox("Power mục tiêu:", options=[0.80, 0.85, 0.90, 0.95], index=0)

    if not saved_df.empty:
        st.info(
            "🔄 Đã tự động lấy dữ liệu cũ từ Database. Bạn hãy nhập thêm [Mã thiết bị] rồi cuộn xuống bấm Lưu để cập nhật nhé.")
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
            {"Chọn chạy (Vàng)": True, "Target concentration": 6.37, "TEa%": 12.9, "Swrl = CVwrl %": 0.5,
             "Sr = CVr %": 0.3},
        ])

    st.markdown("### 📝 Nhập thông số và Chọn mức nồng độ")
    edited_input = st.data_editor(
        init_input_df, num_rows="dynamic", use_container_width=True,
        column_config={
            "Chọn chạy (Vàng)": st.column_config.CheckboxColumn("✅ Chọn chạy"),
            "Target concentration": st.column_config.NumberColumn("Target concentration", format="%.3f"),
            "TEa%": st.column_config.NumberColumn("TEa%", format="%.3f"),
            "Swrl = CVwrl %": st.column_config.NumberColumn("Swrl = CVwrl %", format="%.3f"),
            "Sr = CVr %": st.column_config.NumberColumn("Sr = CVr %", format="%.3f"),
        }, key=f"editor_{test_name}_{machine_name}"
    )

    selected_input = edited_input[edited_input["Chọn chạy (Vàng)"] == True]
    num_selected_levels = max(1, min(3, len(selected_input) if len(selected_input) > 0 else 1))

    if not selected_input.empty:
        all_options_rows = []
        chosen_final_rows = []

        for idx, row in selected_input.iterrows():
            if pd.isna(row["Target concentration"]) or pd.isna(row["TEa%"]): continue
            target_c = float(row["Target concentration"])
            candidates = get_candidate_options_for_row(
                tea_pct=float(row["TEa%"]), swrl_pct=float(row["Swrl = CVwrl %"]), sr_pct=float(row["Sr = CVr %"]),
                num_levels=num_selected_levels, desired_power=global_target_power, sort_mode="min_n_max_power"
            )

            with st.container(border=True):
                st.markdown(f"📍 **Mức nồng độ: {target_c:.3f} {unit_name}**")
                if len(candidates) > 1:
                    option_labels = [
                        f"🏆 Hạng #{i + 1}: Cần {c['n_samples']} mẫu | Power = {c['power_achieve']:.3f} | RL = {c['rejection_limit']:.3f}%"
                        for i, c in enumerate(candidates)]
                    selected_label = st.selectbox("Chọn phương án chốt:", options=option_labels, index=0,
                                                  key=f"pick_{test_name}_{machine_name}_{idx}",
                                                  label_visibility="collapsed")
                    picked_cand = candidates[option_labels.index(selected_label)]
                else:
                    picked_cand = candidates[0]
                    st.info(
                        f"💡 Cần {picked_cand['n_samples']} mẫu | Power = {picked_cand['power_achieve']:.3f} | RL = {picked_cand['rejection_limit']:.3f}%")

            chosen_final_rows.append({
                "Chọn chạy (Vàng)": True, "TÊN XN": test_name, "Machine": machine_name,
                "Target concentration": target_c, "Đơn vị": unit_name,
                "TEa%": float(row["TEa%"]), "Swrl = CVwrl %": float(row["Swrl = CVwrl %"]),
                "Sr = CVr %": float(row["Sr = CVr %"]),
                "CD% (Critical Difference)": picked_cand["cd_pct"], "CD/Swrl": picked_cand["cd_swrl"],
                "Sr/Swrl": picked_cand["sr_swrl"],
                "Power achieve": picked_cand["power_achieve"],
                "Number of sample each concentration": picked_cand["n_samples"],
                "Rejection limit": picked_cand["rejection_limit"], "k": picked_cand["k_factor"],
                "Ghi chú": picked_cand["note"]
            })

        if st.button("✅ LƯU CẤU HÌNH VÀO DATABASE", type="primary", use_container_width=True):
            save_test_setup(test_name, machine_name, machine_id, pd.DataFrame(chosen_final_rows))
            st.success(f"Đã lưu thành công thông số và Mã thiết bị cho {test_name} - {machine_name}!")

# ==============================================================================
# TAB 4: DANH MỤC CẤU HÌNH ĐÃ LƯU
# ==============================================================================
with tab_db_setup:
    st.markdown("### 📂 Cấu hình Stage 1 hiện có")

    # --- KHU VỰC THÊM TÍNH NĂNG IMPORT FILE ---
    with st.expander("📥 KHÔI PHỤC / IMPORT DỮ LIỆU TỪ FILE (.CSV, .XLSX)", expanded=False):
        uploaded_file = st.file_uploader("Kéo thả file dữ liệu đã xuất từ hệ thống vào đây", type=["csv", "xlsx"])
        if uploaded_file is not None:
            try:
                # Đọc file dựa vào đuôi mở rộng
                if uploaded_file.name.endswith(".csv"):
                    df_import = pd.read_csv(uploaded_file)
                else:
                    df_import = pd.read_excel(uploaded_file)

                # Kiểm tra tính hợp lệ của file (Phải có ít nhất test_name và target_conc)
                if "test_name" in df_import.columns and "target_conc" in df_import.columns:
                    st.info(
                        f"Phát hiện **{len(df_import)}** dòng cấu hình hợp lệ. Dữ liệu trùng lặp sẽ tự động được ghi đè.")
                    if st.button("✅ Bắt đầu Import Dữ liệu", type="primary", use_container_width=True):
                        count = import_stage1_from_dataframe(df_import)
                        st.success(
                            f"🎉 Đã khôi phục thành công {count} dòng cấu hình! Hệ thống sẽ tải lại sau 1 giây...")
                        import time

                        time.sleep(1)
                        st.rerun()
                else:
                    st.error(
                        "❌ File không đúng định dạng! Vui lòng chỉ tải lên file được xuất trực tiếp từ hệ thống phần mềm này.")
            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi đọc file: {e}")
    # ------------------------------------------

    all_saved_df = load_all_setups()
    if not all_saved_df.empty:
        st.dataframe(all_saved_df, use_container_width=True, hide_index=True)
        c_d1, c_d2, c_d3 = st.columns([2, 2, 1])
        with c_d1:
            del_test = st.selectbox("Chọn XN cần xóa:", all_saved_df["test_name"].unique(), key="d_t1")
        with c_d2:
            del_mach = st.selectbox("Chọn Máy cần xóa:", all_saved_df["machine"].unique(), key="d_m1")
        with c_d3:
            st.write("")
            st.write("")
            if st.button("❌ Xóa cấu hình", use_container_width=True):
                delete_specific_setup(del_test, del_mach)
                st.rerun()