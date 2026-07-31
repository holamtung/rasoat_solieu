# @title
# 1. Cài đặt và import các thư viện cần thiết
pip -q install polars openpyxl xlsxwriter ipywidgets fastexcel

import io
import time
import ipywidgets as widgets
from IPython.display import display, clear_output
import polars as pl
from google.colab import files

# Cấu hình hiển thị đầy đủ bảng trên Colab
pl.Config.set_tbl_rows(-1)
pl.Config.set_fmt_str_lengths(100)

# Biến toàn cục để lưu kết quả
final_df = None

# 2. Xây dựng các thành phần UI (Widgets)
style = {"description_width": "initial"}

title_html = widgets.HTML(
    value="""
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 5px solid #0d6efd; margin-bottom: 15px;">
        <h3 style="margin: 0; color: #0d6efd; font-family: sans-serif;">RÀ SOÁT DỮ LIỆU TRÙNG LẶP (SỐ BNN & Số C/O)</h3>
        <p style="margin: 5px 0 0 0; color: #6c757d; font-size: 14px;">Tải lên file Excel để tự động phân tích và phát hiện vi phạm.</p>
    </div>
    """,
)

uploader = widgets.FileUpload(
    description="Chọn file Excel",
    accept=".xlsx, .xls",
    multiple=False,
    button_style="info",
    icon="upload",
    style=style,
)

run_button = widgets.Button(
    description="Kiểm tra Dữ Liệu",
    disabled=True,
    button_style="success",
    icon="play",
    tooltip="Nhấn để bắt đầu phân tích dữ liệu",
)

download_button = widgets.Button(
    description="Tải Báo Cáo rà soát",
    disabled=True,
    button_style="warning",
    icon="download",
    tooltip="Nhấn để tải file kết quả Excel",
)

progress_bar = widgets.FloatProgress(
    value=0.0,
    min=0.0,
    max=1.0,
    description='Đang xử lý:',
    bar_style='info',
    orientation='horizontal',
    layout=widgets.Layout(width='100%', visibility='hidden')
)

# Khu vực hiển thị log và bảng kết quả
output_area = widgets.Output()

# 3. Xử lý sự kiện
def on_file_upload(change):
  if uploader.value:
    run_button.disabled = False
    download_button.disabled = True
    with output_area:
      clear_output()
      file_name = list(uploader.value.keys())[0]
      file_size = len(uploader.value[file_name]["content"]) / 1024
      print(f"Đã nhận file: '{file_name}' ({file_size:.2f} KB). Nhấn nút 'Chạy Rà Soát' để bắt đầu.")

uploader.observe(on_file_upload, names="value")

def on_button_clicked(b):
  global final_df
  progress_bar.layout.visibility = 'visible'
  progress_bar.value = 0.1
  with output_area:
    clear_output()
    if not uploader.value:
      print("[CẢNH BÁO] Bạn chưa chọn file nào. Vui lòng tải file lên!")
      progress_bar.layout.visibility = 'hidden'
      return

    try:
      file_name = list(uploader.value.keys())[0]
      file_bytes = uploader.value[file_name]["content"]

      progress_bar.value = 0.3
      try:
        df = pl.read_excel(io.BytesIO(file_bytes))
        if "Số GP" not in df.columns:
          df = pl.read_excel(io.BytesIO(file_bytes), read_options={"has_header": True}, skip_rows=1)
      except Exception:
        df = pl.read_excel(io.BytesIO(file_bytes), skip_rows=1)

      progress_bar.value = 0.5
      df_gp_dup = df.filter(pl.col("Số GP").is_not_null() & pl.col("Số GP").cast(pl.String).str.strip_chars().str.starts_with("BNN")).group_by("Số GP").len().filter(pl.col("len") > 1)
      dup_gp_list = df_gp_dup.get_column("Số GP").to_list()

      progress_bar.value = 0.7
      df_e2_extracted = df.filter(pl.col("Đề xuất khác").is_not_null()).select([pl.col("Số TK"), pl.col("Đề xuất khác").cast(pl.String).str.extract_all(r"E2[A-Za-z0-9/\\-_]+").alias("Mã E2 List")]).explode("Mã E2 List").filter(pl.col("Mã E2 List").is_not_null()).with_columns(pl.col("Mã E2 List").str.strip_chars(".,:; ").alias("Mã E2")).filter(pl.col("Mã E2") != "")
      dup_e2_list = df_e2_extracted.group_by("Mã E2").len().filter(pl.col("len") > 1).get_column("Mã E2").to_list()
      stk_with_dup_e2 = df_e2_extracted.filter(pl.col("Mã E2").is_in(dup_e2_list)).group_by("Số TK").agg(pl.col("Mã E2").unique().str.join(", ").alias("Mã E2 trùng"))

      progress_bar.value = 0.9
      final_df = df.join(stk_with_dup_e2, on="Số TK", how="left").filter(pl.col("Số GP").is_in(dup_gp_list) | pl.col("Mã E2 trùng").is_not_null()).with_columns([pl.when(pl.col("Số GP").is_in(dup_gp_list) & pl.col("Mã E2 trùng").is_not_null()).then(pl.lit("Trùng GP & Mã E2")).when(pl.col("Số GP").is_in(dup_gp_list)).then(pl.lit("Trùng Số GP (BNN)")).otherwise(pl.lit("Trùng Mã E2")).alias("Lý do trùng")]).select(["Số TK", "Ngày ĐK", "Tên doanh nghiệp", "Đơn vị đối tác", "Đề xuất khác", "Mã E2 trùng", "Số GP", "Lý do trùng"]).sort("Lý do trùng")

      progress_bar.value = 1.0
      print(f"ĐÃ PHÁT HIỆN: {final_df.height} DÒNG VI PHẠM TRÙNG LẶP\n")
      display(final_df)
      download_button.disabled = False

    except Exception as e:
      print(f"[LỖI XỬ LÝ]: {str(e)}")
    finally:
      time.sleep(1)
      progress_bar.layout.visibility = 'hidden'

def on_download_clicked(b):
  if final_df is not None:
    output_file = "Ket_qua_trung_lap_tong_hop.xlsx"
    final_df.write_excel(output_file)
    files.download(output_file)

run_button.on_click(on_button_clicked)
download_button.on_click(on_download_clicked)

# 5. Hiển thị
controls = widgets.VBox([
    widgets.HBox([uploader, run_button, download_button]),
    progress_bar,
    output_area
])
display(title_html, controls)
