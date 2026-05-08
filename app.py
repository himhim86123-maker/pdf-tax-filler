"""
PDF智能填表系统 v3.0 - 最终修正版
=====================================
1. 所有右对齐字段使用精确的方格右边线 (x=565.5)
2. 所有居中字段使用方格中心点
3. 覆盖区域覆盖整个方格，确保原文字完全被清除
4. 新增4个签章字段（经办人、身份证号、受理人、受理日期）
5. 自动添加两位小数功能（從業人數除外）
"""

import streamlit as st
import fitz  # PyMuPDF
import os
import io

st.set_page_config(page_title="PDF智能填表系统", layout="wide")

# ============================================================
# 字体处理
# ============================================================

def extract_font_from_pdf(pdf_path):
    """从PDF中提取SimSun字体数据"""
    doc = fitz.open(pdf_path)
    try:
        font_data = None
        for xref in range(1, doc.xref_length()):
            try:
                obj = doc.xref_object(xref)
                if obj and "FontFile2" in obj:
                    font_data = doc.xref_stream(xref)
                    if font_data and len(font_data) > 1000:
                        break
            except:
                continue
        if not font_data:
            for xref in range(1, doc.xref_length()):
                try:
                    if doc.xref_get_key(xref, "Subtype")[1] == "/TrueType":
                        font_data = doc.xref_stream(xref)
                        if font_data and len(font_data) > 1000:
                            break
                except:
                    continue
        if not font_data:
            best_size = 0
            for xref in range(1, doc.xref_length()):
                try:
                    data = doc.xref_stream(xref)
                    if data and len(data) > best_size:
                        best_size = len(data)
                        font_data = data
                except:
                    continue
        return font_data
    finally:
        doc.close()

def save_font_temp(pdf_path):
    """保存字体到临时文件"""
    font_data = extract_font_from_pdf(pdf_path)
    if font_data:
        font_path = "/tmp/simsun.ttf"
        with open(font_path, "wb") as f:
            f.write(font_data)
        return font_path
    return None

# ============================================================
# 两位小数格式化
# ============================================================

def format_two_decimal(value, field_key):
    """自动添加两位小数"""
    if not value or not str(value).strip():
        return value
    # 從業人數不加小数
    if field_key.startswith("eq"):
        return str(value).strip()
    # 签章文字字段不加小数
    if field_key in ["agent_name", "agent_id", "receiver", "receive_date"]:
        return str(value).strip()
    text = str(value).strip().replace(",", "")
    # 已有小数点或百分号，保持不变
    if "." in text or "%" in text:
        return str(value).strip()
    # 纯数字，添加千分位和两位小数
    try:
        num = float(text)
        if num == int(num):
            return f"{int(num):,}.00"
        else:
            return f"{num:,.2f}"
    except ValueError:
        return str(value).strip()

# ============================================================
# 字段坐标配置
# ============================================================

RIGHT_EDGE = 565.5
BASELINE_OFFSET = 6.9

FIELD_CFG = {
    # === 從業人數 (居中) ===
    "eq1s":  (0, 149.9, 170.2, "c"),
    "eq1e":  (0, 201.6, 170.2, "c"),
    "eq2s":  (0, 253.4, 170.2, "c"),
    "eq2e":  (0, 305.1, 170.2, "c"),
    "eq3e":  (0, 563.5, 170.2, "c"),
    
    # === 資產總額 (右對齊) ===
    "aq1s":  (0, 151.9, 189.0, "r"),
    "aq1e":  (0, 203.6, 189.0, "r"),
    "aq2s":  (0, 255.3, 189.0, "r"),
    "aq2e":  (0, 307.0, 189.0, "r"),
    "aq3e":  (0, RIGHT_EDGE, 189.0, "r"),
    
    # === 預繳稅款計算 (右對齊) ===
    "L1":    (0, RIGHT_EDGE, 297.7, "r"),
    "L2":    (0, RIGHT_EDGE, 316.5, "r"),
    "L3":    (0, RIGHT_EDGE, 335.2, "r"),
    "L4":    (0, RIGHT_EDGE, 354.0, "r"),
    "L5":    (0, RIGHT_EDGE, 372.7, "r"),
    "L6":    (0, RIGHT_EDGE, 391.5, "r"),
    "L7":    (0, RIGHT_EDGE, 410.2, "r"),
    "L8":    (0, RIGHT_EDGE, 429.0, "r"),
    "L9":    (0, RIGHT_EDGE, 447.7, "r"),
    "L10":   (0, RIGHT_EDGE, 466.5, "r"),
    "L11":   (0, RIGHT_EDGE, 485.2, "r"),
    "L12":   (0, RIGHT_EDGE, 504.0, "r"),
    "L13":   (0, RIGHT_EDGE, 522.7, "r"),
    "L13_1": (0, RIGHT_EDGE, 541.5, "r"),
    "L14":   (0, RIGHT_EDGE, 560.2, "r"),
    "L15":   (0, RIGHT_EDGE, 579.0, "r"),
    "L16":   (0, RIGHT_EDGE, 597.7, "r"),
    
    # === 匯總納稅 ===
    "L17":   (0, RIGHT_EDGE, 631.5, "r"),
    "L18":   (0, RIGHT_EDGE, 650.2, "r"),
    "L19":   (0, RIGHT_EDGE, 669.0, "r"),
    "L20":   (0, RIGHT_EDGE, 687.7, "r"),
    "L21":   (0, RIGHT_EDGE, 706.5, "r"),
    "L22":   (0, RIGHT_EDGE, 725.2, "r"),
    
    # === 附註 ===
    "FZ1":   (0, RIGHT_EDGE, 762.7, "r"),
    "FZ2":   (0, RIGHT_EDGE, 781.5, "r"),
    "L23":   (0, RIGHT_EDGE, 800.0, "r"),
    
    # === 第2頁 簽章區域 ===
    "agent_name":   (1, 104.8, 103.4, "l"),
    "agent_id":     (1, 184.8, 113.0, "l"),
    "receiver":     (1, 383.7, 103.4, "l"),
    "receive_date": (1, 403.7, 122.6, "l"),
}

# ============================================================
# PDF填写核心函数
# ============================================================

def fill_pdf_core(pdf_bytes, font_path, values):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    FONT_NAME = "SimSun"
    FONT_SIZE = 8
    TEXT_COLOR = (0, 0, 0)
    COVER_COLOR = (1, 1, 1)
    
    try:
        for key, raw_value in values.items():
            if not raw_value or key not in FIELD_CFG:
                continue
            new_value = format_two_decimal(raw_value, key)
            if not new_value:
                continue
            page_num, x_ref, y_top, align = FIELD_CFG[key]
            page = doc[page_num]
            text = str(new_value)
            tw = len(text) * FONT_SIZE * 0.6
            baseline_y = y_top + BASELINE_OFFSET
            text_height = 8
            
            if align == "c":
                cover_left = x_ref - 28
                cover_right = x_ref + 28
                write_x = x_ref - tw / 2
            elif align == "r":
                if key.startswith(("aq", "eq")) and key != "aq3e":
                    cover_left = x_ref - 55
                elif key.startswith("eq"):
                    cover_left = x_ref - 55
                else:
                    cover_left = 514.5
                cover_right = x_ref + 3
                write_x = x_ref - tw
            else:
                cover_left = x_ref - 2
                cover_right = x_ref + tw + 10
                write_x = x_ref
            
            cover_rect = fitz.Rect(cover_left, y_top - 1, cover_right, y_top + text_height + 2)
            shape = page.new_shape()
            shape.draw_rect(cover_rect)
            shape.finish(color=COVER_COLOR, fill=COVER_COLOR)
            shape.commit()
            
            kwargs = {"fontname": FONT_NAME, "fontsize": FONT_SIZE, "color": TEXT_COLOR}
            if font_path and os.path.exists(font_path):
                kwargs["fontfile"] = font_path
            page.insert_text((write_x, baseline_y), text, **kwargs)
        
        output = io.BytesIO()
        doc.save(output, garbage=4, deflate=True)
        output.seek(0)
        return output.getvalue()
    finally:
        doc.close()

# ============================================================
# Streamlit UI
# ============================================================

def main():
    st.title("📄 PDF智能填表系统")
    st.markdown("上传企业所得税申报表PDF模板，修改指定字段，生成字体完全一致的新PDF")
    
    st.header("1️⃣ 上传PDF模板")
    uploaded_file = st.file_uploader("选择PDF文件", type=["pdf"])
    
    if uploaded_file is None:
        st.info("👆 请先上传PDF模板文件")
        st.markdown("""
        **使用步骤：**
        1. 上传原始PDF模板
        2. 在下方输入框填写新数值（留空=不修改）
        3. 点击「生成填好的PDF」按钮
        4. 下载生成的PDF文件
        **自动两位小数：** 输入 12345 → 自动变为 12,345.00
        """)
        return
    
    pdf_bytes = uploaded_file.getvalue()
    
    with st.spinner("正在提取字体..."):
        tmp_pdf = "/tmp/uploaded_template.pdf"
        with open(tmp_pdf, "wb") as f:
            f.write(pdf_bytes)
        font_path = save_font_temp(tmp_pdf)
    
    if font_path:
        st.success("✅ SimSun字体提取成功")
    else:
        st.warning("⚠️ 无法提取字体，将使用系统默认字体")
        font_path = None
    
    st.header("2️⃣ 填写字段数据")
    st.caption("留空表示不修改。数字自动添加两位小数（從業人數和签章字段除外）")
    
    values = {}
    
    with st.expander("👥 第1行 - 從業人數（不加小数）", expanded=True):
        cols = st.columns(5)
        labels = ["Q1季初", "Q1季末", "Q2季初", "Q2季末", "Q3季末(總)"]
        keys = ["eq1s", "eq1e", "eq2s", "eq2e", "eq3e"]
        for col, label, key in zip(cols, labels, keys):
            with col:
                values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("💰 第2行 - 資產總額（萬元）", expanded=True):
        cols = st.columns(5)
        labels = ["Q1季初", "Q1季末", "Q2季初", "Q2季末", "Q3季末(總)"]
        keys = ["aq1s", "aq1e", "aq2s", "aq2e", "aq3e"]
        for col, label, key in zip(cols, labels, keys):
            with col:
                values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("📊 第3-16行 - 預繳稅款計算（自动两位小数）", expanded=True):
        col_left, col_right = st.columns(2)
        left_fields = [
            ("L1", "3 營業收入"), ("L2", "4 營業成本"), ("L3", "5 利潤總額"),
            ("L4", "6 特定業務"), ("L5", "7 不徵稅收入"),
            ("L6", "8 減：資產加速折舊"), ("L7", "9 減：免稅收入"),
        ]
        right_fields = [
            ("L8", "10 減：所得減免"), ("L9", "11 減：所得減免其他"),
            ("L10", "12 實際利潤額"), ("L11", "13 稅率(25%)"),
            ("L12", "14 應納所得稅額"), ("L13", "15 減免所得稅額"),
            ("L13_1", "15.1 減免明細"),
        ]
        with col_left:
            for key, label in left_fields:
                values[key] = st.text_input(label, value="", key=key)
        with col_right:
            for key, label in right_fields:
                values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("📋 第16-25行 - 其他稅額及附註"):
        col_left, col_right = st.columns(2)
        left_fields2 = [
            ("L14", "16 本期預繳"), ("L15", "17 減免其他"),
            ("L16", "18 本期應補(退)"), ("L17", "19 總機構本期"),
            ("L18", "20 總機構分攤"), ("L19", "21 財政集中"),
        ]
        right_fields2 = [
            ("L20", "22 分支機構"), ("L21", "23 分攤比例"),
            ("L22", "24 稅率"), ("FZ1", "附1 中央級收入"),
            ("FZ2", "附2 地方級收入"), ("L23", "25 減免地方"),
        ]
        with col_left:
            for key, label in left_fields2:
                values[key] = st.text_input(label, value="", key=key)
        with col_right:
            for key, label in right_fields2:
                values[key] = st.text_input(label, value="", key=key)
    
    with st.expander("✏️ 第2頁 - 签章信息（不加小数）"):
        col_left, col_right = st.columns(2)
        with col_left:
            values["agent_name"] = st.text_input("经办人", value="", key="agent_name")
            values["agent_id"] = st.text_input("经办人身份证号", value="", key="agent_id")
        with col_right:
            values["receiver"] = st.text_input("受理人", value="", key="receiver")
            values["receive_date"] = st.text_input("受理日期", value="", key="receive_date",
                                                   placeholder="如：2025年04月14日")
    
    st.header("3️⃣ 生成PDF")
    filled_values = {k: v.strip() for k, v in values.items() if v.strip()}
    
    if filled_values:
        st.success(f"✅ 已填写 {len(filled_values)} 个字段")
    else:
        st.info("💡 尚未填写任何字段")
    
    if st.button("🚀 生成填好的PDF", type="primary", disabled=len(filled_values) == 0):
        with st.spinner("正在生成PDF..."):
            try:
                result = fill_pdf_core(pdf_bytes, font_path, filled_values)
                st.success("✅ PDF生成成功！")
                st.download_button(
                    label="📥 点击下载填好的PDF",
                    data=result,
                    file_name="filled_tax_form.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                try:
                    preview_doc = fitz.open(stream=result, filetype="pdf")
                    st.subheader("📋 预览")
                    for i, p in enumerate(preview_doc):
                        pix = p.get_pixmap(dpi=100)
                        st.image(pix.tobytes("png"), caption=f"第{i+1}页")
                    preview_doc.close()
                except Exception as e:
                    st.warning(f"预览失败: {e}")
            except Exception as e:
                st.error(f"❌ 生成失败: {e}")
                st.exception(e)
    
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        **自动两位小数：**
        - 输入 123456 → 显示 123,456.00
        - 输入已带小数点 → 保持不变
        - 從業人數和签章字段 → 不加小数
        
        **字段总数：** 40个
        - 從業人數：5个
        - 資產總額：5个
        - 預繳稅款：21个
        - 汇总纳税：6个
        - 附註：3个
        - 签章信息：4个（第2页）
        """)
    
    st.markdown("---")
    st.markdown("<center>PDF智能填表系统 v3.0 | 坐标精确修复 | 自动两位小数 | 40字段支持</center>",
                unsafe_allow_html=True)

if __name__ == "__main__":
    main()
