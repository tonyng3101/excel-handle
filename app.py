import streamlit as st
import pandas as pd
import os
import tempfile
import io
from openpyxl import load_workbook

st.title("🔥 Tool Tổng Hợp Học Phí Của Quỳnh")

uploaded_files = st.file_uploader(
    "📁 Chọn file Excel",
    type=["xls", "xlsx"],
    accept_multiple_files="directory"
)

# =========================
# HÀM CHECK TÊN SHEET ↔ A4
# =========================
def check_sheet_name(excel_wb, file_name):
    results = []

    for sheet_name in excel_wb.sheetnames:
        ws = excel_wb[sheet_name]
        header_value = ws["A4"].value

        if not header_value:
            status = "ERROR"
            message = "Ô A4 đang trống hoặc không hợp lệ"
        else:
            header_str = str(header_value).strip()
            if sheet_name in header_str:
                status = "OK"
                message = "Tên sheet khớp tên lớp"
            else:
                status = "WARNING"
                message = "Tên sheet chưa khớp tên lớp"

        results.append({
            "FileName": file_name,
            "SheetName": sheet_name,
            "Check_Ten_Lop": status,
            "Ghi_Chu_Ten_Lop": message
        })

    return results


if st.button("🚀 Xử lý dữ liệu"):

    if not uploaded_files:
        st.error("❌ Bạn chưa upload file nào!")
        st.stop()

    temp_dir = tempfile.mkdtemp()
    all_data = []
    all_checks = []

    for up_file in uploaded_files:

        content = up_file.read()
        safe_name = os.path.basename(up_file.name)
        file_path = os.path.join(temp_dir, safe_name)

        with open(file_path, "wb") as f:
            f.write(content)

        st.write(f"🔄 Đang xử lý {safe_name}")

        ext = os.path.splitext(file_path)[1].lower()
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"

        # =========================
        # PHẦN CŨ: ĐỌC DỮ LIỆU HỌC PHÍ
        # =========================
        xls = pd.ExcelFile(file_path, engine=engine)

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                header=None,
                engine=engine
            )

            header_row = 9
            start_row = 11

            if len(df) <= start_row:
                continue

            df_data = df.iloc[start_row:].dropna(how="all")
            if df_data.empty:
                continue

            fixed = df_data.iloc[:, :8]

            header_data = df.iloc[header_row]
            keep_idx = [
                i for i, v in enumerate(header_data)
                if pd.isna(v) and i >= 10
            ]
            keep = df_data.iloc[:, keep_idx]

            merged = pd.concat([fixed, keep], axis=1)
            merged.columns = range(merged.shape[1])

            merged["FileName"] = safe_name
            merged["SheetName"] = sheet_name

            all_data.append(merged)

        # =========================
        # PHẦN MỚI: CHECK TÊN SHEET
        # =========================
        if ext == ".xlsx":
            excel_wb = load_workbook(file_path, data_only=True)
            sheet_checks = check_sheet_name(excel_wb, safe_name)
            all_checks.extend(sheet_checks)

    if not all_data:
        st.error("❌ Không có dữ liệu hợp lệ để tổng hợp")
        st.stop()

    # =========================
    # GỘP DỮ LIỆU HỌC PHÍ
    # =========================
    final_df = pd.concat(all_data, ignore_index=True)

    # =========================
    # GỘP KẾT QUẢ CHECK VÀO TỪNG DÒNG
    # =========================
    check_df = pd.DataFrame(all_checks)

    if not check_df.empty:
        final_df = final_df.merge(
            check_df,
            on=["FileName", "SheetName"],
            how="left"
        )
    else:
        final_df["Check_Ten_Lop"] = ""
        final_df["Ghi_Chu_Ten_Lop"] = ""

    st.success("🎉 Hoàn tất xử lý!")

    # =========================
    # XUẤT FILE EXCEL
    # =========================
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        final_df.to_excel(
            writer,
            index=False,
            sheet_name="TongHopHocPhi"
        )

    buffer.seek(0)

    st.download_button(
        "⬇️ Tải file tổng hợp",
        data=buffer,
        file_name="TongHop_HocPhi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
