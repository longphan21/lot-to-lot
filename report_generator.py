import io
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE


def generate_report_docx(eval_date_str, test_name, unit, machine, machine_id, lot_old, lot_new, stage1_df, stage2_df):
    template_path = os.path.join(os.path.dirname(__file__), 'template.docx')
    doc = Document(template_path) if os.path.exists(template_path) else Document()

    # Mã màu xanh đen chuẩn của form MEDIC
    theme_color = RGBColor(0, 32, 96)

    # 1. Chuyển đổi định dạng ngày thành DD/MM/YYYY
    d_parts = eval_date_str.split('-')
    if len(d_parts) == 3:
        d_ngay, d_thang, d_nam = d_parts[2], d_parts[1], d_parts[0]
        eval_date_formatted = f"{d_ngay}/{d_thang}/{d_nam}"
        chuoi_ngay_thang_chu_ky = f"{d_ngay} tháng {d_thang} năm {d_nam}"
    else:
        d_ngay, d_thang, d_nam = "...", "...", "..."
        eval_date_formatted = eval_date_str
        chuoi_ngay_thang_chu_ky = "[Ngày] tháng [Tháng] năm [Năm]"

    overall_pass = all(row['status'] == "PASS" for _, row in stage2_df.iterrows())
    ket_luan_tong = "2 lot thuốc thử trên tương đồng" if overall_pass else "2 lot thuốc thử trên KHÔNG tương đồng"

    replacements = {
        "[Tên thiết bị]": str(machine).strip(),
        "[Mã thiết bị]": str(machine_id).strip(),
        "[Tên xét nghiệm]": f"{test_name} ({unit})".strip(),
        "[Lô hóa chất cũ]": str(lot_old).strip(),
        "[Ngày thay lô]": eval_date_formatted,
        "[Lô hóa chất mới]": str(lot_new).strip(),
        "[Kết luận tổng]": ket_luan_tong,
        "[Ngày] tháng [Tháng] năm [Năm]": chuoi_ngay_thang_chu_ky
    }

    def replace_in_paragraph(p):
        full_text = p.text
        for key, val in replacements.items():
            if key in full_text:
                is_italic = None
                is_bold = None
                for r in p.runs:
                    if r.text.strip():
                        is_italic = r.italic
                        is_bold = r.bold
                        break

                if "Kết luận:" in full_text and key == "[Kết luận tổng]":
                    p.text = ""
                    r1 = p.add_run("Kết luận: ")
                    r1.bold = True
                    r1.font.name = 'Times New Roman'
                    r1.font.size = Pt(12)
                    r1.font.color.rgb = theme_color

                    r2 = p.add_run(val)
                    r2.italic = True
                    r2.font.name = 'Times New Roman'
                    r2.font.size = Pt(12)
                    r2.font.color.rgb = theme_color
                else:
                    p.text = full_text.replace(key, val)
                    if len(p.runs) > 0:
                        p.runs[0].font.name = 'Times New Roman'
                        p.runs[0].font.size = Pt(12)
                        if is_italic is not None: p.runs[0].italic = is_italic
                        if is_bold is not None: p.runs[0].bold = is_bold
                        p.runs[0].font.color.rgb = theme_color

    for p in doc.paragraphs: replace_in_paragraph(p)

    # Ép toàn bộ các ô trong Word (bao gồm Header) canh giữa dọc và thay chữ
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    replace_in_paragraph(p)
                for nested_table in cell.tables:
                    for n_row in nested_table.rows:
                        for n_cell in n_row.cells:
                            n_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                            for n_p in n_cell.paragraphs: replace_in_paragraph(n_p)

    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            for p in header.paragraphs: replace_in_paragraph(p)
            for table in header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs: replace_in_paragraph(p)
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            for p in footer.paragraphs: replace_in_paragraph(p)

    # 3. TỰ ĐỘNG XỬ LÝ BẢNG STAGE 1 (BẢNG D1)
    for table in doc.tables:
        if len(table.rows) > 0 and len(table.rows[0].cells) > 1 and "TEa%" in table.rows[0].cells[1].text:

            # Gắn đơn vị vào Tiêu đề cột đầu tiên
            hdr_cell = table.rows[0].cells[0]
            if unit not in hdr_cell.text:
                hdr_cell.text = hdr_cell.text.replace('\n', ' ').strip() + f"\n({unit})"
                for p in hdr_cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(12)
                        r.bold = True
                        r.font.color.rgb = theme_color

            while len(table.rows) > 1:
                table._element.remove(table.rows[-1]._element)

            stage1_selected = stage1_df[stage1_df["is_selected"] == 1]
            for _, row_data in stage1_selected.iterrows():
                new_row = table.add_row()
                new_row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                new_row.height = Cm(0.8)

                vals = [
                    str(row_data['target_conc']).strip(),  # CHỈ LẤY SỐ
                    str(row_data['tea_pct']).strip(),
                    str(row_data['cd_swrl']).strip(),
                    str(row_data['sr_swrl']).strip(),
                    str(row_data['k_factor']).strip(),
                    f"{row_data['power_achieve'] * 100:.1f}%".strip(),
                    f"{row_data['rejection_limit']}%".strip(),
                    str(row_data['n_samples']).strip()
                ]
                for c_idx, val in enumerate(vals):
                    cell = new_row.cells[c_idx]
                    cell.text = val
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(4)
                        p.paragraph_format.space_after = Pt(4)
                        if len(p.runs) > 0:
                            p.runs[0].font.name = 'Times New Roman'
                            p.runs[0].font.size = Pt(12)
                            p.runs[0].font.color.rgb = theme_color  # ÉP MÀU XANH ĐEN
            break

    # 4. TỰ ĐỘNG XỬ LÝ BẢNG STAGE 2 (BẢNG D2)
    for table in doc.tables:
        if len(table.rows) > 0 and len(table.rows[0].cells) > 1 and "SID" in table.rows[0].cells[1].text:

            # Gắn đơn vị vào Tiêu đề cột đầu tiên
            hdr_cell = table.rows[0].cells[0]
            if unit not in hdr_cell.text:
                hdr_cell.text = hdr_cell.text.replace('\n', ' ').strip() + f"\n({unit})"
                for p in hdr_cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(12)
                        r.bold = True
                        r.font.color.rgb = theme_color

            while len(table.rows) > 1:
                table._element.remove(table.rows[-1]._element)

            start_row_idx = 1
            for level_name, group in stage2_df.groupby('level_name', sort=False):
                target_conc = group['target_conc'].iloc[0]
                n_samples = len(group)
                group_start_idx = start_row_idx

                level_status = "Đạt" if all(
                    group.iloc[i]['status'] == "PASS" for i in range(n_samples)) else "Không Đạt"

                for idx_in_group in range(n_samples):
                    row_data = group.iloc[idx_in_group]
                    new_row = table.add_row()
                    new_row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                    new_row.height = Cm(0.8)
                    start_row_idx += 1

                    non_merged_vals = {
                        1: str(row_data['sid']).strip(),
                        2: str(row_data['val_old']).strip(),
                        3: str(row_data['val_new']).strip()
                    }

                    for c_idx, val in non_merged_vals.items():
                        cell = new_row.cells[c_idx]
                        cell.text = val
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                        for p in cell.paragraphs:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p.paragraph_format.space_before = Pt(4)
                            p.paragraph_format.space_after = Pt(4)
                            if len(p.runs) > 0:
                                p.runs[0].font.name = 'Times New Roman'
                                p.runs[0].font.size = Pt(12)
                                p.runs[0].font.color.rgb = theme_color  # ÉP MÀU XANH ĐEN

                merged_data = {
                    0: str(target_conc).strip(),  # CHỈ LẤY SỐ, BỎ ĐƠN VỊ
                    4: f"{group.iloc[0]['mean_d_pct']:.3f}".strip(),
                    5: f"{group.iloc[0]['rl_pct']:.3f}".strip(),
                    6: level_status
                }

                for c_idx, text_val in merged_data.items():
                    if n_samples > 1:
                        target_cell = table.cell(group_start_idx, c_idx).merge(table.cell(start_row_idx - 1, c_idx))
                    else:
                        target_cell = table.cell(group_start_idx, c_idx)

                    target_cell.text = text_val
                    target_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    for p in target_cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(4)
                        p.paragraph_format.space_after = Pt(4)
                        if len(p.runs) > 0:
                            p.runs[0].font.name = 'Times New Roman'
                            p.runs[0].font.size = Pt(12)
                            p.runs[0].font.color.rgb = theme_color  # ÉP MÀU XANH ĐEN
            break

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out