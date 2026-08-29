# COPYRIGHT: HỒ LÂM TÙNG
"""Rà soát dữ liệu C/O-BNN và hóa đơn thương mại từ tệp Excel."""

import hashlib
import io
import unicodedata

import pandas as pd
import polars as pl
import streamlit as st


HEADER_SCAN_ROWS = 30
E2_PATTERN = r"\bE2[A-Z0-9/_\\-]+"
INVALID_INVOICE_VALUES = ("", "NAN", "NONE")

BNN_REQUIRED_COLUMNS = {
    "Số TK",
    "Ngày ĐK",
    "Tên doanh nghiệp",
    "Đơn vị đối tác",
    "Đề xuất khác",
    "Số GP",
}
INVOICE_REQUIRED_COLUMNS = {
    "Số TK",
    "Ngày ĐK",
    "Mã DN",
    "Đơn vị đối tác",
    "Số hoá đơn TM",
}


class InputFileError(ValueError):
    """Lỗi dữ liệu đầu vào có thể hướng dẫn người dùng tự khắc phục."""


def normalise_text(value):
    """Chuẩn hoá văn bản để so khớp tiêu đề Excel không phân biệt dấu."""
    text = str(value).replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(text.lower().split())


def make_unique_columns(headers):
    """Tạo tên cột không rỗng và không trùng, giữ nguyên tên gốc khi có thể."""
    result, counts = [], {}
    for header in headers:
        name = "" if pd.isna(header) else str(header).strip()
        name = name or "Unnamed"
        counts[name] = counts.get(name, -1) + 1
        result.append(name if counts[name] == 0 else f"{name}_{counts[name]}")
    return result


def find_header_row(raw_df, header_hints):
    """Chọn dòng có nhiều từ khoá tiêu đề nhất trong phần đầu của báo cáo."""
    normalised_hints = [normalise_text(hint) for hint in header_hints]
    best_row, best_score = None, 0

    for row_index in range(min(HEADER_SCAN_ROWS, len(raw_df))):
        row_values = raw_df.iloc[row_index].dropna().tolist()
        row_text = normalise_text(" ".join(map(str, row_values)))
        score = sum(hint in row_text for hint in normalised_hints)
        if score > best_score:
            best_row, best_score = row_index, score

    if best_row is None:
        hints = ", ".join(header_hints)
        raise InputFileError(
            f"Không tìm thấy dòng tiêu đề trong {HEADER_SCAN_ROWS} dòng đầu "
            f"(cần nhận diện: {hints})."
        )
    return best_row


def read_excel_safe(file_bytes, header_hints):
    """Đọc Excel một lần, tự nhận diện tiêu đề và trả về Polars DataFrame."""
    try:
        raw_df = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)
    except Exception as exc:
        raise InputFileError(f"Không thể đọc tệp Excel: {exc}") from exc

    if raw_df.empty:
        raise InputFileError("Tệp Excel không có dữ liệu.")

    header_row = find_header_row(raw_df, header_hints)
    data_df = raw_df.iloc[header_row + 1 :].copy()
    data_df.columns = make_unique_columns(raw_df.iloc[header_row].tolist())
    data_df = data_df.dropna(how="all").fillna("")
    return pl.from_pandas(data_df, include_index=False)


def require_columns(df, required_columns):
    missing = sorted(required_columns - set(df.columns))
    if missing:
        raise InputFileError(
            "Tệp Excel thiếu cột bắt buộc: " + ", ".join(missing) + "."
        )


def clean_key(column_name, uppercase=False):
    """Chuẩn hoá chuỗi trước khi đối chiếu nhưng vẫn giữ cột hiển thị gốc."""
    expression = pl.col(column_name).cast(pl.String).fill_null("").str.strip_chars()
    return expression.str.to_uppercase() if uppercase else expression


def find_invoice_column(df):
    """Nhận diện một cột hóa đơn; dừng rõ ràng nếu tệp có nhiều cột phù hợp."""
    canonical_name = "Số hoá đơn TM"
    if canonical_name in df.columns:
        return canonical_name

    candidates = [
        column
        for column in df.columns
        if "hoa don tm" in normalise_text(column)
    ]
    if not candidates:
        raise InputFileError("Không tìm thấy cột 'Số hoá đơn TM'.")
    if len(candidates) > 1:
        raise InputFileError(
            "Có nhiều cột có thể là 'Số hoá đơn TM': " + ", ".join(candidates)
        )
    return candidates[0]


def analyse_bnn(file_bytes):
    df = read_excel_safe(file_bytes, header_hints=["Số GP", "Số TK"])
    require_columns(df, BNN_REQUIRED_COLUMNS)

    df = df.with_columns(
        clean_key("Số TK").alias("_stk_key"),
        clean_key("Số GP", uppercase=True).alias("_gp_key"),
    )

    duplicate_gp_keys = (
        df.filter(
            (pl.col("_gp_key") != "") & pl.col("_gp_key").str.starts_with("BNN")
        )
        .group_by("_gp_key")
        .len()
        .filter(pl.col("len") > 1)
        .select("_gp_key")
        .with_columns(pl.lit(True).alias("_has_duplicate_gp"))
    )

    # Một mã E2 chỉ bị coi là trùng khi nằm trên từ hai số tờ khai khác nhau.
    e2_by_declaration = (
        df.select(
            pl.col("_stk_key"),
            clean_key("Đề xuất khác", uppercase=True)
            .str.extract_all(E2_PATTERN)
            .alias("Mã E2"),
        )
        .explode("Mã E2")
        .filter(pl.col("Mã E2").is_not_null() & (pl.col("_stk_key") != ""))
        .with_columns(clean_key("Mã E2", uppercase=True).alias("Mã E2"))
        .unique(["_stk_key", "Mã E2"])
    )

    duplicate_e2_codes = (
        e2_by_declaration.group_by("Mã E2")
        .agg(pl.col("_stk_key").n_unique().alias("_declaration_count"))
        .filter(pl.col("_declaration_count") > 1)
        .select("Mã E2")
    )
    duplicate_e2_by_declaration = (
        e2_by_declaration.join(duplicate_e2_codes, on="Mã E2", how="inner")
        .group_by("_stk_key")
        .agg(pl.col("Mã E2").unique().str.join(", ").alias("Mã E2 trùng"))
    )

    return (
        df.join(duplicate_gp_keys, on="_gp_key", how="left")
        .join(duplicate_e2_by_declaration, on="_stk_key", how="left")
        .with_columns(pl.col("_has_duplicate_gp").fill_null(False))
        .filter(pl.col("_has_duplicate_gp") | pl.col("Mã E2 trùng").is_not_null())
        .with_columns(
            pl.when(
                pl.col("_has_duplicate_gp") & pl.col("Mã E2 trùng").is_not_null()
            )
            .then(pl.lit("Trùng GP & Mã E2"))
            .when(pl.col("_has_duplicate_gp"))
            .then(pl.lit("Trùng Số GP (BNN)"))
            .otherwise(pl.lit("Trùng Mã E2"))
            .alias("Lý do trùng")
        )
        .select(
            "Số TK",
            "Ngày ĐK",
            "Tên doanh nghiệp",
            "Đơn vị đối tác",
            "Đề xuất khác",
            "Mã E2 trùng",
            "Số GP",
            "Lý do trùng",
        )
        .sort(["Lý do trùng", "Số TK"])
    )


def analyse_invoices(file_bytes):
    df = read_excel_safe(file_bytes, header_hints=["đơn TM", "Số TK", "Mã DN"])
    invoice_column = find_invoice_column(df)
    if invoice_column != "Số hoá đơn TM":
        df = df.rename({invoice_column: "Số hoá đơn TM"})
    require_columns(df, INVOICE_REQUIRED_COLUMNS)

    df_valid = (
        df.with_columns(
            clean_key("Số TK").alias("_stk_key"),
            clean_key("Mã DN", uppercase=True).alias("_tax_code_key"),
            clean_key("Đơn vị đối tác", uppercase=True).alias("_partner_key"),
            clean_key("Số hoá đơn TM", uppercase=True).alias("_invoice_key"),
        )
        .filter(
            (pl.col("_stk_key") != "")
            & ~pl.col("_invoice_key").is_in(INVALID_INVOICE_VALUES)
        )
    )

    key_columns = ["_tax_code_key", "_partner_key", "_invoice_key"]
    duplicate_invoice_groups = (
        df_valid.group_by(key_columns)
        .agg(pl.col("_stk_key").n_unique().alias("_declaration_count"))
        .filter(pl.col("_declaration_count") > 1)
        .select(key_columns)
    )

    return (
        df_valid.join(duplicate_invoice_groups, on=key_columns, how="inner")
        .select("Số TK", "Ngày ĐK", "Mã DN", "Đơn vị đối tác", "Số hoá đơn TM")
        .unique()
        .sort(["Mã DN", "Số hoá đơn TM", "Số TK"])
    )


def dataframe_to_xlsx_bytes(df):
    output = io.BytesIO()
    df.write_excel(output)
    return output.getvalue()


def uploaded_file_token(uploaded_file):
    """Mã nội dung giúp xóa kết quả cũ ngay khi người dùng đổi tệp."""
    return hashlib.sha256(uploaded_file.getvalue()).hexdigest()


def reset_result_when_upload_changes(uploaded_file, prefix):
    result_key = f"{prefix}_result"
    download_key = f"{prefix}_download"
    token_key = f"{prefix}_file_token"

    if uploaded_file is None:
        st.session_state.pop(result_key, None)
        st.session_state.pop(download_key, None)
        st.session_state.pop(token_key, None)
        return

    token = uploaded_file_token(uploaded_file)
    if st.session_state.get(token_key) != token:
        st.session_state[token_key] = token
        st.session_state.pop(result_key, None)
        st.session_state.pop(download_key, None)


def render_result(prefix, report_name, download_label):
    result = st.session_state.get(f"{prefix}_result")
    if result is None:
        return

    st.success(f"ĐÃ PHÁT HIỆN: {result.height} DÒNG VI PHẠM {report_name}")
    st.dataframe(result, use_container_width=True)
    st.download_button(
        label=download_label,
        data=st.session_state[f"{prefix}_download"],
        file_name=f"Ket_qua_{prefix}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_{prefix}",
    )


st.set_page_config(layout="wide")
st.markdown(
    """
<div style="background-color: #f8f9fa; padding: 20px; border-radius: 14px;
    border-left: 14px solid #0d6efd; margin-bottom: 20px;">
  <h3 style="margin: 0; color: #0d6efd; font-family: sans-serif;">
    RÀ SOÁT DỮ LIỆU TRÙNG LẶP
  </h3>
  <p style="margin: 10px 0 0 0; color: #6c757d; font-size: 18px;">
    Tải tệp Excel xuất từ SLXNK. Công cụ tự nhận diện dòng tiêu đề trong 30 dòng đầu.
  </p>
</div>
""",
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)

with col1:
    st.markdown("<h4 style='color: #0d6efd;'>1. Kiểm tra C/O - BNN</h4>", unsafe_allow_html=True)
    uploaded_file_bnn = st.file_uploader(
        "Tải file Excel (C/O - BNN)", type=["xlsx", "xls"], key="uploader_bnn"
    )
    reset_result_when_upload_changes(uploaded_file_bnn, "CO_BNN")

    if uploaded_file_bnn and st.button(
        "Kiểm tra dữ liệu C/O - BNN", key="btn_bnn", type="primary"
    ):
        with st.spinner("Đang xử lý dữ liệu C/O-BNN..."):
            try:
                result = analyse_bnn(uploaded_file_bnn.getvalue())
                st.session_state["CO_BNN_result"] = result
                st.session_state["CO_BNN_download"] = dataframe_to_xlsx_bytes(result)
            except InputFileError as exc:
                st.error(f"[LỖI DỮ LIỆU C/O-BNN]: {exc}")
            except Exception as exc:
                st.error(f"[LỖI XỬ LÝ C/O-BNN]: {exc}")

    render_result("CO_BNN", "TRÙNG LẶP C/O-BNN", "Tải Báo Cáo C/O-BNN")

with col2:
    st.markdown("<h4 style='color: #198754;'>2. Kiểm tra Hoá Đơn TM</h4>", unsafe_allow_html=True)
    uploaded_file_hd = st.file_uploader(
        "Tải file Excel (Hoá Đơn)", type=["xlsx", "xls"], key="uploader_hd"
    )
    reset_result_when_upload_changes(uploaded_file_hd, "Hoa_Don")

    if uploaded_file_hd and st.button(
        "Kiểm tra dữ liệu Hoá Đơn", key="btn_hd", type="primary"
    ):
        with st.spinner("Đang xử lý dữ liệu hoá đơn..."):
            try:
                result = analyse_invoices(uploaded_file_hd.getvalue())
                st.session_state["Hoa_Don_result"] = result
                st.session_state["Hoa_Don_download"] = dataframe_to_xlsx_bytes(result)
            except InputFileError as exc:
                st.error(f"[LỖI DỮ LIỆU HOÁ ĐƠN]: {exc}")
            except Exception as exc:
                st.error(f"[LỖI XỬ LÝ HOÁ ĐƠN]: {exc}")

    render_result("Hoa_Don", "TRÙNG HOÁ ĐƠN", "Tải Báo Cáo Hoá Đơn")

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #6c757d; font-size: 20px;'>"
    "BẢN QUYỀN: HỒ LÂM TÙNG - 0988 767413 - CHI CỤC HẢI QUAN KHU VỰC VII</p>",
    unsafe_allow_html=True,
)
