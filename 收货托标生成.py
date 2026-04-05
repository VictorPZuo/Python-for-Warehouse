# -*- coding: utf-8 -*-
import io
import re
from typing import Tuple

import streamlit as st
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.graphics.barcode import code128
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Paragraph
from reportlab.pdfgen import canvas


# =========================
# 基础配置
# =========================
PAGE_WIDTH = 6 * inch   # 横向 6 inch
PAGE_HEIGHT = 4 * inch  # 横向 4 inch
PAGE_SIZE = (PAGE_WIDTH, PAGE_HEIGHT)

MARGIN_X = 0.35 * inch
TOP_MARGIN = 0.28 * inch
BOTTOM_MARGIN = 0.25 * inch
AVAILABLE_WIDTH = PAGE_WIDTH - 2 * MARGIN_X

# 行高与版式
CONTAINER_LINE_H = 0.52 * inch
CLIENT_LINE_H = 0.40 * inch
SKU_LINE_H = 0.72 * inch
BARCODE_LINE_H = 1.05 * inch
QTY_LINE_H = 0.36 * inch

# 字体
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"


def sanitize_filename(text: str) -> str:
    text = re.sub(r"[^\w\-]+", "_", text.strip())
    return text or "receiving_labels"


def normalize_text(text: str) -> str:
    return (text or "").strip()


def split_container_display(container_no: str) -> Tuple[str, str]:
    """
    返回：(前缀, 后四位)
    若长度不足 4，则前缀为空，整体视为“后四位部分”
    """
    container_no = normalize_text(container_no)
    if len(container_no) <= 4:
        return "", container_no
    return container_no[:-4], container_no[-4:]


def calc_sku_font_size(sku: str) -> float:
    """
    根据 SKU 长度自适应字号。
    同时后续还会用实际宽度做二次收缩，避免溢出。
    """
    length = len(sku)
    if length <= 8:
        return 28
    if length <= 12:
        return 24
    if length <= 18:
        return 20
    if length <= 25:
        return 16
    return 13


def fit_font_size_to_width(text: str, font_name: str, start_size: float, max_width: float, min_size: float = 8) -> float:
    """
    按实际文本宽度动态收缩字号，保证内容不超出最大宽度。
    """
    size = start_size
    while size > min_size and stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    return max(size, min_size)


def draw_centered_paragraph(c: canvas.Canvas, paragraph: Paragraph, y_top: float, frame_width: float) -> None:
    """
    将 Paragraph 按指定宽度绘制在页面水平居中位置，y_top 为顶部坐标。
    """
    w, h = paragraph.wrap(frame_width, PAGE_HEIGHT)
    x = (PAGE_WIDTH - frame_width) / 2
    paragraph.drawOn(c, x, y_top - h)


def build_styles():
    styles = getSampleStyleSheet()

    container_style = ParagraphStyle(
        "container_style",
        parent=styles["Normal"],
        fontName=FONT_REGULAR,
        fontSize=20,
        leading=22,
        alignment=TA_CENTER,
    )

    client_style = ParagraphStyle(
        "client_style",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=20,
        leading=22,
        alignment=TA_CENTER,
    )

    qty_style = ParagraphStyle(
        "qty_style",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=15,
        leading=17,
        alignment=TA_RIGHT,
    )

    return container_style, client_style, qty_style


def draw_one_label(
    c: canvas.Canvas,
    container_no: str,
    client_code: str,
    sku: str,
) -> None:
    container_no = normalize_text(container_no)
    client_code = normalize_text(client_code)
    sku = normalize_text(sku)

    container_style, client_style, qty_style = build_styles()

    y = PAGE_HEIGHT - TOP_MARGIN

    # -------------------------
    # 第 1 行：集装箱号（后四位加粗+下划线）
    # -------------------------
    prefix, last4 = split_container_display(container_no)
    # 使用 XML 转义友好的最小替换
    prefix_safe = prefix.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    last4_safe = last4.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    container_html = f'{prefix_safe}<u><b>{last4_safe}</b></u>'
    p1 = Paragraph(container_html, container_style)
    draw_centered_paragraph(c, p1, y, AVAILABLE_WIDTH)
    y -= CONTAINER_LINE_H

    # -------------------------
    # 第 2 行：客户代码（居中）
    # -------------------------
    client_safe = client_code.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    p2 = Paragraph(f"<b>{client_safe}</b>", client_style)
    draw_centered_paragraph(c, p2, y, AVAILABLE_WIDTH)
    y -= CLIENT_LINE_H

    # -------------------------
    # 第 3 行：SKU（自适应字号，居中）
    # -------------------------
    sku_font_size = calc_sku_font_size(sku)
    sku_font_size = fit_font_size_to_width(
        sku,
        FONT_BOLD,
        sku_font_size,
        max_width=AVAILABLE_WIDTH,
        min_size=8,
    )

    c.setFont(FONT_BOLD, sku_font_size)
    c.drawCentredString(PAGE_WIDTH / 2, y - sku_font_size, sku)
    y -= SKU_LINE_H

    # -------------------------
    # 第 4 行：Code128 条码（根据 SKU）
    # -------------------------
    barcode_max_width = AVAILABLE_WIDTH * 0.82
    barcode_height = 0.72 * inch

    # 先按一个较保守的 barWidth 生成，再做宽度检查
    bar_width = 1.0
    barcode_obj = code128.Code128(
        sku,
        barWidth=bar_width,
        barHeight=barcode_height,
        humanReadable=False,
    )

    # 如果条码太宽，则按比例缩小 barWidth
    if barcode_obj.width > barcode_max_width and barcode_obj.width > 0:
        scale = barcode_max_width / barcode_obj.width
        bar_width *= scale
        barcode_obj = code128.Code128(
            sku,
            barWidth=bar_width,
            barHeight=barcode_height,
            humanReadable=False,
        )

    barcode_x = (PAGE_WIDTH - barcode_obj.width) / 2
    barcode_y = y - barcode_height + 2
    barcode_obj.drawOn(c, barcode_x, barcode_y)
    y -= BARCODE_LINE_H

    # -------------------------
    # 第 5 行：Qty: __________（右对齐）
    # 十个空白字符以“下划线长度”呈现，便于打印后手写
    # -------------------------
    qty_text = "Qty: __________"
    p5 = Paragraph(f"<b>{qty_text}</b>", qty_style)
    w5, h5 = p5.wrap(AVAILABLE_WIDTH, PAGE_HEIGHT)
    p5.drawOn(c, PAGE_WIDTH - MARGIN_X - w5, max(BOTTOM_MARGIN, y - h5))


def generate_pdf(container_no: str, client_code: str, sku: str, qty: int) -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=PAGE_SIZE)

    for _ in range(qty):
        draw_one_label(c, container_no, client_code, sku)
        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="收货标签生成器", page_icon="🏷️", layout="centered")

st.title("🏷️ 收货标签生成器")
st.caption("输入收货信息，生成可直接打印的 PDF 标签。每页 1 张，横向 4 inch × 6 inch。")

with st.form("label_form"):
    container_no = st.text_input("1. 集装箱号", placeholder="例如：TCLU1234567")
    client_code = st.text_input("2. 客户代码", placeholder="例如：JDL / BASEUS / ABC")
    sku = st.text_input("3. SKU", placeholder="例如：SKU-123456789")
    qty = st.number_input("4. 所需标签数量", min_value=1, max_value=5000, value=1, step=1)

    submitted = st.form_submit_button("生成 PDF")

if submitted:
    container_no = normalize_text(container_no)
    client_code = normalize_text(client_code)
    sku = normalize_text(sku)

    missing_fields = []
    if not container_no:
        missing_fields.append("集装箱号")
    if not client_code:
        missing_fields.append("客户代码")
    if not sku:
        missing_fields.append("SKU")

    if missing_fields:
        st.error("请补全以下字段：" + "、".join(missing_fields))
    else:
        try:
            pdf_bytes = generate_pdf(
                container_no=container_no,
                client_code=client_code,
                sku=sku,
                qty=int(qty),
            )

            filename = f"{sanitize_filename(container_no)}_{sanitize_filename(client_code)}_{sanitize_filename(sku)}_labels.pdf"

            st.success("PDF 已生成。")
            st.download_button(
                label="下载 PDF",
                data=pdf_bytes,
                file_name=filename,
                mime="application/pdf",
            )

            st.info(
                "打印建议：\n"
                "- 纸张方向：横向\n"
                "- 页面尺寸：4 × 6 inch\n"
                "- 缩放：100% / Actual Size\n"
                "- 每页 1 张标签"
            )
        except Exception as e:
            st.exception(e)
