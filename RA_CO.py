# COPYRIGHT: HỒ LÂM TÙNG
import io
import polars as pl
import streamlit as st

st.set_page_config(layout="wide")

st.markdown(
    """
<div style="background-color: #f8f9fa; padding: 20px; border-radius: 14px; border-left: 14px solid #0d6efd; margin-bottom: 20px;">
<h3 style="margin: 0; color: #0d6efd; font-family: sans-serif;">RÀ SOÁT DỮ LIỆU TRÙNG LẶP</h3>
<p style="margin: 10px 0 0 0; color: #6c757d; font-size: 18px;">Tải lên file Excel để tự động phân tích và phát hiện vi phạm C/O, BNN hoặc Hoá đơn TM.</p>
</div>
""",
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader("Chọn file Excel", type=["xlsx", "xls"], key="uploader")

if uploaded_file is not None:
    # Đọc file chung cho cả 2 chức năng
    file_bytes = uploaded_file.read()
    try:
        df = pl.read_excel(io.BytesIO(file_bytes))
        if "Số TK" not in df.columns:
            df = pl.read_excel(io.BytesIO(file_bytes), read_options={"has_header": True}, skip_rows=1)
    except Exception:
        df = pl.read_excel(io.BytesIO(file_bytes), skip_rows=1)

    # Tạo 2 cột cho 2 nút bấm
    col1, col2 = st.columns(2)

    # ==========================================
    # CHỨC NĂNG 1: KIỂM TRA CO-BNN (BÊN TRÁI)
    # ==========================================
    if col1.button("Kiểm tra CO-BNN", type="primary", use_container_width=True):
        with st.spinner("Đang xử lý dữ liệu CO-BNN..."):
            try:
                df_gp_dup = (
                    df.filter(
                        pl.col("Số GP").is_not_null()
                        & pl.col("Số GP").cast(pl.String).str.strip_chars().str.starts_with("BNN")
                    )
                    .group_by("Số GP").len().filter(pl.col("len") > 1)
                )
                dup_gp_list = df_gp_dup.get_column("Số GP").to_list()

                df_e2_extracted = (
                    df.filter(pl.col("Đề xuất khác").is_not_null())
                    .select([
                        pl.col("Số TK"),
                        pl.col("Đề xuất khác").cast(pl.String).str.extract_all(r"E2[A-Za-z0-9/\\-_]+").alias("Mã E2 List"),
                    ])
                    .explode("Mã E2 List")
                    .filter(pl.col("Mã E2 List").is_not_null())
                    .with_columns(pl.col("Mã E2 List").str.strip_chars(".,:; ").alias("Mã E2"))
                    .filter(pl.col("Mã E2") != "")
                )

                dup_e2_list = (
                    df_e2_extracted.group_by("Mã E2").len().filter(pl.col("len") > 1).get_column("Mã E2").to_list()
                )

                stk_with_dup_e2 = (
                    df_e2_extracted.filter(pl.col("Mã E2").is_in(dup_e2_list))
                    .group_by("Số TK")
                    .agg(pl.col("Mã E2").unique().str.join(", ").alias("Mã E2 trùng"))
                )

                final_df = (
                    df.join(stk_with_dup_e2, on="Số TK", how="left")
                    .filter(pl.col("Số GP").is_in(dup_gp_list) | pl.col("Mã E2 trùng").is_not_null())
                    .with_columns([
                        pl.when(pl.col("Số GP").is_in(dup_gp_list) & pl.col("Mã E2 trùng").is_not_null())
                        .then(pl.lit("Trùng GP & Mã E2"))
                        .when(pl.col("Số GP").is_in(dup_gp_list))
                        .then(pl.lit("Trùng Số GP (BNN)"))
                        .otherwise(pl.lit("Trùng Mã E2"))
                        .alias("Lý do trùng")
                    ])
                    .select(["Số TK", "Ngày ĐK", "Tên doanh nghiệp", "Đơn vị đối tác", "Đề xuất khác", "Mã E2 trùng", "Số GP", "Lý do trùng"])
                    .sort("Lý do trùng")
                )

                st.success(f"ĐÃ PHÁT HIỆN: {final_df.height} DÒNG VI PHẠM TRÙNG LẶP CO-BNN")
                st.dataframe(final_df.to_pandas(), use_container_width=True)

                output_file = "Ket_qua_CO_BNN.xlsx"
                final_df.write_excel(output_file)
                with open(output_file, "rb") as f:
                    st.download_button(label="Tải Báo Cáo CO-BNN", data=f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"[LỖI XỬ LÝ CO-BNN]: {str(e)}")

    # ==========================================
    # CHỨC NĂNG 2: KIỂM TRA HOÁ ĐƠN (BÊN PHẢI)
    # ==========================================
    if col2.button("Kiểm tra hoá đơn", type="primary", use_container_width=True):
        with st.spinner("Đang xử lý dữ liệu Hoá đơn..."):
            try:
                # 1. Lọc các dòng có "Số hoá đơn TM" hợp lệ (không null và không rỗng)
                df_valid = df.filter(
                    pl.col("Số hoá đơn TM").is_not_null() & 
                    (pl.col("Số hoá đơn TM").cast(pl.String).str.strip_chars() != "")
                )
                
                # 2. Tìm các nhóm trùng (Mã DN + Đơn vị đối tác + Số hoá đơn TM) có từ 2 Số TK khác nhau trở lên
                dup_groups = (
                    df_valid.group_by(["Mã DN", "Đơn vị đối tác", "Số hoá đơn TM"])
                    .agg(pl.col("Số TK").n_unique().alias("so_tk_count"))
                    .filter(pl.col("so_tk_count") > 1)
                    .select(["Mã DN", "Đơn vị đối tác", "Số hoá đơn TM"])
                )
                
                # 3. Kết xuất ra số liệu cuối cùng thoả mãn điều kiện
                final_hd_df = (
                    df_valid.join(dup_groups, on=["Mã DN", "Đơn vị đối tác", "Số hoá đơn TM"], how="inner")
                    .select(["Số TK", "Ngày ĐK", "Mã DN", "Đơn vị đối tác", "Số hoá đơn TM"])
                    .unique() # Lọc unique để loại bỏ những dòng trùng Số TK giống hệt nhau ở ví dụ 2
                    .sort(["Mã DN", "Số hoá đơn TM", "Số TK"])
                )

                st.success(f"ĐÃ PHÁT HIỆN: {final_hd_df.height} DÒNG VI PHẠM TRÙNG HOÁ ĐƠN")
                st.dataframe(final_hd_df.to_pandas(), use_container_width=True)

                output_file_hd = "Ket_qua_Hoa_Don.xlsx"
                final_hd_df.write_excel(output_file_hd)
                with open(output_file_hd, "rb") as f:
                    st.download_button(label="Tải Báo Cáo Hoá Đơn", data=f, file_name=output_file_hd, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"[LỖI XỬ LÝ HOÁ ĐƠN]: Hãy chắc chắn file có đủ các cột 'Số TK', 'Ngày ĐK', 'Mã DN', 'Đơn vị đối tác', 'Số hoá đơn TM'. Chi tiết lỗi: {str(e)}")

st.markdown("---")
st.markdown("<p style='text-align: center; color: #6c757d; font-size: 13px;'><b>COPYRIGHT: HỒ LÂM TÙNG - 0988 767413 - CHI CỤC HẢI QUAN KHU VỰC VII</b></p>", unsafe_allow_html=True)
