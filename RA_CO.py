# COPYRIGHT: HỒ LÂM TÙNG
import io
import polars as pl
import streamlit as st

st.set_page_config(layout="wide")

st.markdown(
    """
<div style="background-color: #f8f9fa; padding: 20px; border-radius: 14px; border-left: 14px solid #0d6efd; margin-bottom: 20px;">
<h3 style="margin: 0; color: #0d6efd; font-family: sans-serif;">RÀ SOÁT DỮ LIỆU TRÙNG LẶP</h3>
<p style="margin: 10px 0 0 0; color: #6c757d; font-size: 18px;">Tải lên file Excel vào đúng chức năng tương ứng để kiểm tra.</p>
</div>
""",
    unsafe_allow_html=True,
)

# Chia giao diện thành 2 cột độc lập
col1, col2 = st.columns(2)

# ==========================================
# CỘT TRÁI: KIỂM TRA CO-BNN
# ==========================================
with col1:
    st.markdown("<h4 style='color: #0d6efd;'>1. Kiểm tra C/O - BNN</h4>", unsafe_allow_html=True)
    uploaded_file_bnn = st.file_uploader("Tải file Excel (C/O - BNN)", type=["xlsx", "xls"], key="uploader_bnn")

    if uploaded_file_bnn is not None:
        with st.spinner("Đang xử lý dữ liệu CO-BNN..."):
            try:
                file_bytes = uploaded_file_bnn.read()
                try:
                    df = pl.read_excel(io.BytesIO(file_bytes))
                except Exception:
                    df = pl.read_excel(io.BytesIO(file_bytes), read_options={"has_header": True})

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
# CỘT PHẢI: KIỂM TRA HOÁ ĐƠN
# ==========================================
with col2:
    st.markdown("<h4 style='color: #198754;'>2. Kiểm tra Hoá Đơn TM</h4>", unsafe_allow_html=True)
    uploaded_file_hd = st.file_uploader("Tải file Excel (Hoá Đơn)", type=["xlsx", "xls"], key="uploader_hd")

    if uploaded_file_hd is not None:
        with st.spinner("Đang xử lý dữ liệu Hoá đơn..."):
            try:
                file_bytes = uploaded_file_hd.read()
                try:
                    df = pl.read_excel(io.BytesIO(file_bytes))
                except Exception:
                    df = pl.read_excel(io.BytesIO(file_bytes), read_options={"has_header": True})

                # 1. Lọc các dòng có "Số hoá đơn TM" hợp lệ
                df_valid = df.filter(
                    pl.col("Số hoá đơn TM").is_not_null() & 
                    (pl.col("Số hoá đơn TM").cast(pl.String).str.strip_chars() != "")
                )
                
                # 2. Tìm nhóm trùng (Mã DN + Đơn vị đối tác + Số hoá đơn TM) có >= 2 Số TK khác nhau
                dup_groups = (
                    df_valid.group_by(["Mã DN", "Đơn vị đối tác", "Số hoá đơn TM"])
                    .agg(pl.col("Số TK").n_unique().alias("so_tk_count"))
                    .filter(pl.col("so_tk_count") > 1)
                    .select(["Mã DN", "Đơn vị đối tác", "Số hoá đơn TM"])
                )
                
                # 3. Kết xuất
                final_hd_df = (
                    df_valid.join(dup_groups, on=["Mã DN", "Đơn vị đối tác", "Số hoá đơn TM"], how="inner")
                    .select(["Số TK", "Ngày ĐK", "Mã DN", "Đơn vị đối tác", "Số hoá đơn TM"])
                    .unique() 
                    .sort(["Mã DN", "Số hoá đơn TM", "Số TK"])
                )

                st.success(f"ĐÃ PHÁT HIỆN: {final_hd_df.height} DÒNG VI PHẠM TRÙNG HOÁ ĐƠN")
                st.dataframe(final_hd_df.to_pandas(), use_container_width=True)

                output_file_hd = "Ket_qua_Hoa_Don.xlsx"
                final_hd_df.write_excel(output_file_hd)
                with open(output_file_hd, "rb") as f:
                    st.download_button(label="Tải Báo Cáo Hoá Đơn", data=f, file_name=output_file_hd, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"[LỖI XỬ LÝ HOÁ ĐƠN]: {str(e)}")


st.markdown("---")
st.markdown("<p style='text-align: center; color: #6c757d; font-size: 13px;'><b>COPYRIGHT: HỒ LÂM TÙNG - 0988 767413 - CHI CỤC HẢI QUAN KHU VỰC VII</b></p>", unsafe_allow_html=True)
