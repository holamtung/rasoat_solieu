# COPYRIGHT: HỒ LÂM TÙNG
import io
import polars as pl
import streamlit as st

st.markdown(
    """
<div style="background-color: #f8f9fa; padding: 20px; border-radius: 14px; border-left: 14px solid #0d6efd; margin-bottom: 20px;">
<h3 style="margin: 0; color: #0d6efd; font-family: sans-serif;">RÀ SOÁT DỮ LIỆU TRÙNG LẶP (SỐ BNN & Số C/O)</h3>
<p style="margin: 10px 0 0 0; color: #6c757d; font-size: 20px;">Hướng dẫn: File Exel tải về từ phần mềm SLXNK bắt buộc phải xoá toàn bộ 07 dòng đầu tiên, chỉ để lại phần tiêu đề của các cột.</p>
</div>
""",
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "Chọn file Excel", type=["xlsx", "xls"], key="uploader"
)

if uploaded_file is not None:
    with st.spinner("Đang xử lý dữ liệu..."):
        try:
            file_bytes = uploaded_file.read()
            try:
                df = pl.read_excel(io.BytesIO(file_bytes))
                if "Số GP" not in df.columns:
                    df = pl.read_excel(
                        io.BytesIO(file_bytes),
                        read_options={"has_header": True},
                        skip_rows=1,
                    )
            except Exception:
                df = pl.read_excel(io.BytesIO(file_bytes), skip_rows=1)

            df_gp_dup = (
                df.filter(
                    pl.col("Số GP").is_not_null()
                    & pl.col("Số GP")
                    .cast(pl.String)
                    .str.strip_chars()
                    .str.starts_with("BNN")
                )
                .group_by("Số GP")
                .len()
                .filter(pl.col("len") > 1)
            )
            dup_gp_list = df_gp_dup.get_column("Số GP").to_list()

            df_e2_extracted = (
                df.filter(pl.col("Đề xuất khác").is_not_null())
                .select([
                    pl.col("Số TK"),
                    pl.col("Đề xuất khác")
                    .cast(pl.String)
                    .str.extract_all(r"E2[A-Za-z0-9/\\-_]+")
                    .alias("Mã E2 List"),
                ])
                .explode("Mã E2 List")
                .filter(pl.col("Mã E2 List").is_not_null())
                .with_columns(
                    pl.col("Mã E2 List")
                    .str.strip_chars(".,:; ")
                    .alias("Mã E2")
                )
                .filter(pl.col("Mã E2") != "")
            )

            dup_e2_list = (
                df_e2_extracted.group_by("Mã E2")
                .len()
                .filter(pl.col("len") > 1)
                .get_column("Mã E2")
                .to_list()
            )

            stk_with_dup_e2 = (
                df_e2_extracted.filter(pl.col("Mã E2").is_in(dup_e2_list))
                .group_by("Số TK")
                .agg(
                    pl.col("Mã E2").unique().str.join(", ").alias("Mã E2 trùng")
                )
            )

            final_df = (
                df.join(stk_with_dup_e2, on="Số TK", how="left")
                .filter(
                    pl.col("Số GP").is_in(dup_gp_list)
                    | pl.col("Mã E2 trùng").is_not_null()
                )
                .with_columns([
                    pl.when(
                        pl.col("Số GP").is_in(dup_gp_list)
                        & pl.col("Mã E2 trùng").is_not_null()
                    )
                    .then(pl.lit("Trùng GP & Mã E2"))
                    .when(pl.col("Số GP").is_in(dup_gp_list))
                    .then(pl.lit("Trùng Số GP (BNN)"))
                    .otherwise(pl.lit("Trùng Mã E2"))
                    .alias("Lý do trùng")
                ])
                .select([
                    "Số TK",
                    "Ngày ĐK",
                    "Tên doanh nghiệp",
                    "Đơn vị đối tác",
                    "Đề xuất khác",
                    "Mã E2 trùng",
                    "Số GP",
                    "Lý do trùng",
                ])
                .sort("Lý do trùng")
            )

            st.success(f"ĐÃ PHÁT HIỆN: {final_df.height} DÒNG VI PHẠM TRÙNG LẶP")
            st.dataframe(final_df.to_pandas())

            output_file = "Ket_qua_trung_lap_tong_hop.xlsx"
            final_df.write_excel(output_file)
            with open(output_file, "rb") as f:
                st.download_button(
                    label="Tải Báo Cáo rà soát",
                    data=f,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"[LỖI XỬ LÝ]: {str(e)}")

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #6c757d; font-size: 13px;'><b>COPYRIGHT: HỒ LÂM TÙNG - CHI CỤC HẢI QUAN KHU VỰC VII</b></p>",
    unsafe_allow_html=True,
)
