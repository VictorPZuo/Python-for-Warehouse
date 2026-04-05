# -*- coding: utf-8 -*-
import base64
import io
import re
from typing import List

import pandas as pd
import streamlit as st
from barcode.codex import Code128
from barcode.writer import ImageWriter, SVGWriter
from PIL import Image
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


# =========================
# 页面配置：横向 4 x 6 inch
# 实际 PDF：宽 6 inch，高 4 inch
# =========================
PAGE_WIDTH = 6 * inch
PAGE_HEIGHT = 4 * inch
PAGE_SIZE = (PAGE_WIDTH, PAGE_HEIGHT)

MARGIN_X = 0.30 * inch
AVAILABLE_WIDTH = PAGE_WIDTH - 2 * MARGIN_X

FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

# 固定版式位置
Y1 = PAGE_HEIGHT - 0.55 * inch   # 集装箱号
Y2 = PAGE_HEIGHT - 1.05 * inch   # 客户代码
Y3 = PAGE_HEIGHT - 1.70 * inch   # SKU
BARCODE_CENTER_Y = PAGE_HEIGHT - 2.72 * inch
Y5 = PAGE_HEIGHT - 3.60 * inch   # Qty


def normalize_text(text: str) -> str:
    return str(text or "").strip()


def sanitize_filename(text: str) -> str:
    text = re.sub(r"[^\w\-]+", "_", normalize_text(text))
    return text or "receiving_labels"


def split_container(container_no: str):
    container_no = normalize_text(container_no)
    if len(container_no) <= 4:
        return "", container_no
    return container_no[:-4], container_no[-4:]


def fit_font_size(
    text: str,
    font_name: str,
    start_size: float,
    max_width: float,
    min_size: float = 8
) -> float:
    size = start_size
    while size > min_size and stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    return max(size, min_size)


def calc_sku_font_size(sku: str) -> float:
    """
    SKU 字号尽可能大，但需保证：
    1. 不超出页面可用宽度
    2. 整体 1-5 行仍保持在页面高度内
    """
    sku = normalize_text(sku)
    max_size = 34.0
    min_size = 8.0
    max_width = AVAILABLE_WIDTH
    max_visual_height = 0.50 * inch

    size = max_size
    while size >= min_size:
        text_width = stringWidth(sku, FONT_BOLD, size)
        visual_height = size * 1.15
        if text_width <= max_width and visual_height <= max_visual_height:
            return size
        size -= 0.5

    return min_size


def draw_centered_text(
    c: canvas.Canvas,
    text: str,
    y: float,
    font_name: str,
    font_size: float
):
    c.setFont(font_name, font_size)
    c.drawCentredString(PAGE_WIDTH / 2, y, text)


def draw_container_line(c: canvas.Canvas, container_no: str):
    """
    第一行：
    集装箱号，后四位加粗并加下划线，整体居中
    """
    prefix, last4 = split_container(container_no)

    prefix_size = 20
    last4_size = 20

    prefix_width = stringWidth(prefix, FONT_REGULAR, prefix_size)
    last4_width = stringWidth(last4, FONT_BOLD, last4_size)
    total_width = prefix_width + last4_width

    start_x = (PAGE_WIDTH - total_width) / 2

    if prefix:
        c.setFont(FONT_REGULAR, prefix_size)
        c.drawString(start_x, Y1, prefix)

    x_last4 = start_x + prefix_width
    c.setFont(FONT_BOLD, last4_size)
    c.drawString(x_last4, Y1, last4)

    underline_y = Y1 - 2
    c.setLineWidth(1)
    c.line(x_last4, underline_y, x_last4 + last4_width, underline_y)


def draw_client_line(c: canvas.Canvas, client_code: str):
    draw_centered_text(c, client_code, Y2, FONT_BOLD, 20)


def draw_sku_line(c: canvas.Canvas, sku: str):
    font_size = calc_sku_font_size(sku)
    font_size = fit_font_size(
        sku,
        FONT_BOLD,
        font_size,
        AVAILABLE_WIDTH,
        min_size=8
    )
    draw_centered_text(c, sku, Y3, FONT_BOLD, font_size)


def barcode_png_bytes(sku: str) -> bytes:
    """
    使用 python-barcode 生成 PNG 条码
    针对仓库扫码进行增强：
    - 条码高度 >= 1 inch
    - module_width 更大
    - 留白更充分
    """
    sku = normalize_text(sku)
    fp = io.BytesIO()

    barcode = Code128(sku, writer=ImageWriter())
    barcode.write(
        fp,
        options={
            "write_text": False,
            "module_width": 0.35,
            "module_height": 25.0,
            "quiet_zone": 2.0,
            "background": "white",
            "foreground": "black",
            "dpi": 300,
            "format": "PNG",
        },
    )
    fp.seek(0)
    return fp.getvalue()


def barcode_svg_data_uri(sku: str) -> str:
    """
    生成可直接用于 HTML <img> 的 SVG data URI
    用于页面预览，避免直接显示 SVG 源代码
    """
    sku = normalize_text(sku)
    fp = io.BytesIO()

    barcode = Code128(sku, writer=SVGWriter())
    barcode.write(
        fp,
        options={
            "write_text": False,
            "module_width": 0.35,
            "module_height": 25.0,
            "quiet_zone": 2.0,
            "background": "white",
            "foreground": "black",
        },
    )
    fp.seek(0)
    svg_bytes = fp.read()
    svg_b64 = base64.b64encode(svg_bytes).decode("utf-8")
    return f"data:image/svg+xml;base64,{svg_b64}"


def draw_barcode_line(c: canvas.Canvas, sku: str):
    """
    第四行：真实 Code128 条码
    条码打印高度至少 1 inch
    """
    png_bytes = barcode_png_bytes(sku)
    image = Image.open(io.BytesIO(png_bytes))

    img_w, img_h = image.size

    target_h = 1.0 * inch
    scale = target_h / img_h
    target_w = img_w * scale

    max_w = AVAILABLE_WIDTH * 0.95
    if target_w > max_w:
        scale = max_w / img_w
        target_w = img_w * scale
        target_h = img_h * scale

    x = (PAGE_WIDTH - target_w) / 2
    y = BARCODE_CENTER_Y - target_h / 2

    c.drawInlineImage(image, x, y, width=target_w, height=target_h)


def draw_qty_line(c: canvas.Canvas):
    qty_text = "Qty: __________"
    font_size = 15
    text_width = stringWidth(qty_text, FONT_BOLD, font_size)
    x = PAGE_WIDTH - MARGIN_X - text_width
    c.setFont(FONT_BOLD, font_size)
    c.drawString(x, Y5, qty_text)


def draw_one_label(
    c: canvas.Canvas,
    container_no: str,
    client_code: str,
    sku: str
):
    draw_container_line(c, container_no)
    draw_client_line(c, client_code)
    draw_sku_line(c, sku)
    draw_barcode_line(c, sku)
    draw_qty_line(c)


def generate_pdf(records: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=PAGE_SIZE)

    for _, row in records.iterrows():
        container_no = normalize_text(row["集装箱号"])
        client_code = normalize_text(row["客户代码"])
        sku = normalize_text(row["SKU"])
        qty = int(row["标签数量"])

        for _ in range(qty):
            draw_one_label(c, container_no, client_code, sku)
            c.showPage()

    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def validate_records(df: pd.DataFrame) -> pd.DataFrame:
    expected = ["集装箱号", "客户代码", "SKU", "标签数量"]
    df = df.copy()
    df.columns = [normalize_text(col) for col in df.columns]

    missing_cols = [col for col in expected if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Excel 缺少必要列：{', '.join(missing_cols)}")

    df = df[expected].copy()
    df["集装箱号"] = df["集装箱号"].astype(str).map(normalize_text)
    df["客户代码"] = df["客户代码"].astype(str).map(normalize_text)
    df["SKU"] = df["SKU"].astype(str).map(normalize_text)

    df["标签数量"] = pd.to_numeric(df["标签数量"], errors="coerce")
    invalid_qty = df["标签数量"].isna() | (df["标签数量"] < 1)
    if invalid_qty.any():
        bad_rows = (df.index[invalid_qty] + 2).tolist()
        raise ValueError(f"标签数量存在无效值，请检查 Excel 行号：{bad_rows}")

    blank_mask = (
        (df["集装箱号"] == "") |
        (df["客户代码"] == "") |
        (df["SKU"] == "")
    )
    if blank_mask.any():
        bad_rows = (df.index[blank_mask] + 2).tolist()
        raise ValueError(f"存在空白必填字段，请检查 Excel 行号：{bad_rows}")

    df["标签数量"] = df["标签数量"].astype(int)
    return df.reset_index(drop=True)


def create_excel_template() -> bytes:
    template = pd.DataFrame(
        [
            {"集装箱号": "TCLU1234567", "客户代码": "JDL", "SKU": "SKU-001", "标签数量": 2},
            {"集装箱号": "TGHU7654321", "客户代码": "BASEUS", "SKU": "SKU-002-ABC", "标签数量": 3},
        ]
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        template.to_excel(writer, index=False, sheet_name="labels")
    output.seek(0)
    return output.getvalue()


def render_label_preview(container_no: str, client_code: str, sku: str):
    """
    页面预览：显示真正的标签效果，而不是 SVG 代码文本
    """
    prefix, last4 = split_container(container_no)

    safe_prefix = prefix.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    safe_last4 = last4.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    safe_client = client_code.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    safe_sku = sku.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    font_px = int(calc_sku_font_size(sku) * 1.35)
    font_px = max(font_px, 14)

    barcode_uri = barcode_svg_data_uri(sku)

    html = f"""
    <div style="
        width: 600px;
        height: 400px;
        border: 2px solid #d9d9d9;
        border-radius: 10px;
        background: white;
        padding: 18px 26px;
        margin: 8px 0 18px 0;
        box-sizing: border-box;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    ">
        <div style="text-align:center; font-size:28px; line-height:1.15; margin-top:4px;">
            <span>{safe_prefix}</span>
            <span style="font-weight:700; text-decoration: underline;">{safe_last4}</span>
        </div>

        <div style="text-align:center; font-size:28px; line-height:1.1; font-weight:700; margin-top:8px;">
            {safe_client}
        </div>

        <div style="
            text-align:center;
            font-size:{font_px}px;
            line-height:1.05;
            font-weight:700;
            margin-top:10px;
            word-break: break-all;
            overflow-wrap: anywhere;
        ">
            {safe_sku}
        </div>

        <div style="
            display:flex;
            justify-content:center;
            align-items:center;
            margin-top:8px;
            min-height:100px;
        ">
            <img src="{barcode_uri}" style="width:90%; max-height:90px;" />
        </div>

        <div style="text-align:right; font-size:22px; font-weight:700; margin-bottom:4px;">
            Qty: __________
        </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


def collect_manual_records() -> pd.DataFrame:
    st.subheader("手动输入")

    with st.form("manual_form", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            container_no = st.text_input("集装箱号", placeholder="例如：TCLU1234567")
            client_code = st.text_input("客户代码", placeholder="例如：JDL / BASEUS / ABC")

        with col2:
            sku = st.text_input("SKU", placeholder="例如：SKU-123456789")
            qty = st.number_input("标签数量", min_value=1, max_value=10000, value=1, step=1)

        add_row = st.form_submit_button("加入本次生成清单")

    if "manual_rows" not in st.session_state:
        st.session_state.manual_rows = []

    if add_row:
        container_no = normalize_text(container_no)
        client_code = normalize_text(client_code)
        sku = normalize_text(sku)

        missing = []
        if not container_no:
            missing.append("集装箱号")
        if not client_code:
            missing.append("客户代码")
        if not sku:
            missing.append("SKU")

        if missing:
            st.error("请补全以下字段：" + "、".join(missing))
        else:
            st.session_state.manual_rows.append(
                {
                    "集装箱号": container_no,
                    "客户代码": client_code,
                    "SKU": sku,
                    "标签数量": int(qty),
                }
            )
            st.success("已加入本次生成清单。")

    if st.session_state.manual_rows:
        st.markdown("**当前手动输入清单**")
        st.dataframe(pd.DataFrame(st.session_state.manual_rows), use_container_width=True)

        if st.button("清空手动输入清单"):
            st.session_state.manual_rows = []
            st.rerun()

    if st.session_state.manual_rows:
        return pd.DataFrame(st.session_state.manual_rows)

    return pd.DataFrame(columns=["集装箱号", "客户代码", "SKU", "标签数量"])


def collect_excel_records() -> pd.DataFrame:
    st.subheader("Excel 批量导入")
    st.caption("Excel 必须包含以下列名：集装箱号、客户代码、SKU、标签数量")

    template_bytes = create_excel_template()
    st.download_button(
        "下载 Excel 模板",
        data=template_bytes,
        file_name="receiving_label_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

    if uploaded_file is None:
        return pd.DataFrame(columns=["集装箱号", "客户代码", "SKU", "标签数量"])

    try:
        df_excel = pd.read_excel(uploaded_file, engine="openpyxl")
        df_excel = validate_records(df_excel)
        st.success(f"Excel 导入成功，共 {len(df_excel)} 条记录。")
        st.dataframe(df_excel, use_container_width=True)
        return df_excel
    except Exception as e:
        st.error(f"Excel 读取失败：{e}")
        return pd.DataFrame(columns=["集装箱号", "客户代码", "SKU", "标签数量"])


def combine_sources(manual_df: pd.DataFrame, excel_df: pd.DataFrame) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []

    if not manual_df.empty:
        frames.append(manual_df)

    if not excel_df.empty:
        frames.append(excel_df)

    if not frames:
        return pd.DataFrame(columns=["集装箱号", "客户代码", "SKU", "标签数量"])

    result = pd.concat(frames, ignore_index=True)
    result["标签数量"] = result["标签数量"].astype(int)
    return result


def preview_records(df: pd.DataFrame):
    if df.empty:
        st.info("当前还没有可预览的数据。请先手动加入记录或上传 Excel。")
        return

    st.subheader("标签预览")
    st.caption("每个 SKU 仅预览 1 张，用于确认版式；实际 PDF 会按“标签数量”生成多页。")

    preview_df = df.drop_duplicates(subset=["集装箱号", "客户代码", "SKU"]).reset_index(drop=True)

    for idx, row in preview_df.iterrows():
        st.markdown(f"**预览 {idx + 1}**")
        render_label_preview(
            normalize_text(row["集装箱号"]),
            normalize_text(row["客户代码"]),
            normalize_text(row["SKU"]),
        )


# =========================
# Streamlit 页面
# =========================
st.set_page_config(
    page_title="收货标签生成器",
    page_icon="🏷️",
    layout="wide"
)

st.title("🏷️ 收货标签生成器")
st.caption("支持手动输入 + Excel 批量导入；支持同一 PDF 内生成多组不同 SKU 标签；支持先预览后下载 PDF。")

left, right = st.columns([1, 1])

with left:
    manual_df = collect_manual_records()

with right:
    excel_df = collect_excel_records()

all_records = combine_sources(manual_df, excel_df)

st.markdown("---")
st.subheader("本次 PDF 生成清单")

if all_records.empty:
    st.info("暂无待生成记录。")
else:
    st.dataframe(all_records, use_container_width=True)
    total_pages = int(all_records["标签数量"].sum())
    st.write(f"总记录数：**{len(all_records)}**")
    st.write(f"预计 PDF 总页数：**{total_pages}**")

st.markdown("---")
preview_records(all_records)

st.markdown("---")
st.subheader("生成 PDF")

if st.button("生成并下载 PDF", type="primary", disabled=all_records.empty):
    try:
        pdf_bytes = generate_pdf(all_records)
        st.success("PDF 已生成。")
        st.download_button(
            label="下载 PDF",
            data=pdf_bytes,
            file_name="receiving_labels_batch.pdf",
            mime="application/pdf",
        )
    except Exception as e:
        st.exception(e)

st.info(
    "打印建议：\n"
    "- 纸张方向：横向\n"
    "- 页面大小：4 × 6 inch\n"
    "- 缩放：100% / Actual Size\n"
    "- 条码高度已增强，适合仓库扫码场景\n"
    "- 每页 1 张标签"
)
