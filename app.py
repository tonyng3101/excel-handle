import streamlit as st
import pandas as pd
import os
import glob
import tempfile
import io

st.title("🔥 Tool Tổng Hợp Học Phí")

uploaded_files = st.file_uploader(
    "📁 Chọn file Excel",
    type=["xlsx"],
    accept_multiple_files="directory"
)

if st.button("🚀 Xử lý dữ liệu"):

    if not uploaded_files:
        st.error("Bạn chưa upload file nào!")
        st.stop()

    temp_dir = tempfile.mkdtemp()
    all_data = []

    for up_file in uploaded_files:

        # Đọc binary
        content = up_file.read()

        # 🔥 Quan trọng: loại bỏ folder ảo trong tên file
        safe_name = os.path.basename(up_file.name)

        # Lưu vào thư mục tạm
        file_path = os.path.join(temp_dir, safe_name)
        with open(file_path, "wb") as f:
            f.write(content)

        if not os.path.isfile(file_path):
            st.error(f"❌ File không tồn tại sau khi ghi: {file_path}")
            st.stop()

        st.write(f"🔄 Đang xử lý {safe_name}")

        # Đọc Excel
        xls = pd.ExcelFile(file_path, engine="openpyxl")

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            header_row = 9
            start_row = 11

            if len(df) <= start_row:
                continue

            df_data = df.iloc[start_row:].dropna(how="all")
            if df_data.empty:
                continue

            fixed = df_data.iloc[:, :8]

            header_data = df.iloc[header_row]
            keep_idx = [i for i, v in enumerate(header_data) if pd.isna(v) and i >= 10]
            keep = df_data.iloc[:, keep_idx]

            merged = pd.concat([fixed, keep], axis=1)
            merged.columns = range(merged.shape[1])
            merged["SheetName"] = sheet_name
            merged["FileName"] = safe_name

            all_data.append(merged)

    # Gộp dữ liệu
    final_df = pd.concat(all_data, ignore_index=True)
    st.success("🎉 Hoàn tất xử lý!")

    # -------------------------
    #  FIX QUAN TRỌNG CHO to_excel
    # -------------------------
    buffer = io.BytesIO()
    final_df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        "⬇️ Tải file tổng hợp",
        data=buffer,
        file_name="TongHop_HocPhi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
