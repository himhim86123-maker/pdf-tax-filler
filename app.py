import streamlit as st
import fitz
import io
import zipfile
import re
import os
import base64

st.set_page_config(page_title="PDF智能填表系统 v15.1", layout="wide")

CHAR_W = 4.0
COMMA_W = 2.5
DOT_W = 2.0
FONTSIZE = 8.0

COMMA_FONT_B64 = (
    "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"
)

STAMP_LABELS = {
    'agent_name': '经办人：',
    'agent_id': '经办人身份证号：',
    'receiver': '受理人：',
    'receive_date': '受理日期：',
}

FIELD_CFG = {
    "eq1s": (0, 100.6, 152.3, 164.8, 183.6, 0.5),
    "eq1e": (0, 152.6, 204.0, 164.8, 183.6, 0.5),
    "eq2s": (0, 204.3, 255.7, 164.8, 183.6, 0.5),
    "eq2e": (0, 256.0, 307.4, 164.8, 183.6, 0.5),
    "eq3e": (0, 514.5, 565.9, 164.8, 183.6, 0.5),
    "aq1s": (0, 100.6, 152.3, 183.6, 202.3, 0.5),
    "aq1e": (0, 152.6, 204.0, 183.6, 202.3, 0.5),
    "aq2s": (0, 204.3, 255.7, 183.6, 202.3, 0.5),
    "aq2e": (0, 256.0, 307.4, 183.6, 202.3, 0.5),
    "aq3e": (0, 514.5, 565.9, 183.6, 202.3, 0.5),
    "L1": (0, 514.5, 565.9, 297.7, 305.7, 0.2),
    "L2": (0, 514.5, 565.9, 316.5, 324.5, 0.2),
    "L3": (0, 514.5, 565.9, 335.2, 343.2, 0.2),
    "L4": (0, 514.5, 565.9, 354.0, 362.0, 0.2),
    "L5": (0, 514.5, 565.9, 372.7, 380.7, 0.2),
    "L6": (0, 514.5, 565.9, 391.5, 399.5, 0.2),
    "L7": (0, 514.5, 565.9, 410.2, 418.2, 0.2),
    "L8": (0, 514.5, 565.9, 429.0, 437.0, 0.2),
    "L9": (0, 514.5, 565.9, 447.7, 455.7, 0.2),
    "L10": (0, 514.5, 565.9, 466.5, 474.5, 0.2),
    "L11": (0, 514.5, 565.9, 485.2, 493.2, 0.2),
    "L12": (0, 514.5, 565.9, 504.0, 512.0, 0.2),
    "L13": (0, 514.5, 565.9, 522.7, 530.7, 0.2),
    "L13_1": (0, 514.5, 565.9, 541.5, 549.5, 0.2),
    "L14": (0, 514.5, 565.9, 560.2, 568.2, 0.2),
    "L15": (0, 514.5, 565.9, 579.0, 587.0, 0.2),
    "L16": (0, 514.5, 565.9, 597.7, 605.7, 0.2),
    "L17": (0, 514.5, 565.9, 631.5, 639.5, 0.2),
    "L18": (0, 514.5, 565.9, 650.2, 658.2, 0.2),
    "L19": (0, 514.5, 565.9, 669.0, 677.0, 0.2),
    "L20": (0, 514.5, 565.9, 687.7, 695.7, 0.2),
    "L21": (0, 514.5, 565.9, 706.5, 714.5, 0.2),
    "L22": (0, 514.5, 565.9, 725.2, 733.2, 0.2),
    "FZ1": (0, 514.5, 565.9, 762.7, 770.7, 0.2),
    "FZ2": (0, 514.5, 565.9, 781.5, 789.5, 0.2),
    "L23": (0, 514.5, 565.9, 800.0, 808.0, 0.2),
    "L23_2": (1, 514.5, 565.9, 10.0, 27.6, 0.2),
    "FZ3": (1, 514.5, 565.9, 43.0, 51.0, 0.2),
    "L24": (1, 514.5, 565.9, 61.7, 69.7, 0.2),
    "agent_name": (1, 80.8, 184.8, 103.4, 111.4, 2.0),
    "agent_id": (1, 112.8, 264.8, 113.0, 121.0, 2.0),
    "receiver": (1, 355.7, 463.7, 103.4, 111.4, 2.0),
    "receive_date": (1, 355.7, 483.7, 122.6, 130.6, 2.0),
    "A201_R1C1": (2, 224.0, 288.0, 104.0, 112.0, 0.2),
    "A201_R1C2": (2, 288.0, 352.0, 104.0, 112.0, 0.2),
    "A201_R1C3": (2, 352.0, 416.0, 104.0, 112.0, 0.2),
    "A201_R1C4": (2, 416.0, 480.0, 104.0, 112.0, 0.2),
    "A201_R1C5": (2, 480.0, 566.0, 104.0, 112.0, 0.2),
    "A201_R2C1": (2, 224.0, 288.0, 120.7, 128.7, 0.2),
    "A201_R2C2": (2, 288.0, 352.0, 120.7, 128.7, 0.2),
    "A201_R2C3": (2, 352.0, 416.0, 120.7, 128.7, 0.2),
    "A201_R2C4": (2, 416.0, 480.0, 120.7, 128.7, 0.2),
    "A201_R2C5": (2, 480.0, 566.0, 120.7, 128.7, 0.2),
    "A201_R3C1": (2, 224.0, 288.0, 135.7, 143.7, 0.2),
    "A201_R3C2": (2, 288.0, 352.0, 135.7, 143.7, 0.2),
    "A201_R3C3": (2, 352.0, 416.0, 135.7, 143.7, 0.2),
    "A201_R3C4": (2, 416.0, 480.0, 135.7, 143.7, 0.2),
    "A201_R3C5": (2, 480.0, 566.0, 135.7, 143.7, 0.2),
}

def extract_font_data(doc):
    for xref in range(1, doc.xref_length()):
        try:
            obj = doc.xref_object(xref)
            if not obj:
                continue
            if "/FontFile2" in obj:
                match = re.search(r'/FontFile2\s+(\d+)', obj)
                if match:
                    ff2_xref = int(match.group(1))
                    data = doc.xref_stream(ff2_xref)
                    if data and len(data) > 1000:
                        return data
        except:
            continue
    for xref in range(2, min(20, doc.xref_length())):
        try:
            data = doc.xref_stream(xref)
            if data and len(data) > 10000:
                return data
        except:
            continue
    return None

def make_comma_font():
    subset_path = '/tmp/simsun_subset_comma_v14.ttf'
    if os.path.exists(subset_path) and os.path.getsize(subset_path) > 1000:
        return subset_path
    try:
        font_data = base64.b64decode(COMMA_FONT_B64)
        with open(subset_path, 'wb') as f:
            f.write(font_data)
        st.info(f"内嵌字体释放成功: {len(font_data)} bytes")
        return subset_path
    except Exception as e:
        st.error(f"内嵌字体释放失败: {e}")
        return None

def fmt_decimal(value, field_key):
    if not value or not str(value).strip():
        return value
    if field_key.startswith("eq"):
        return str(value).strip()
    if field_key in ["agent_name", "agent_id", "receiver", "receive_date"]:
        return str(value).strip()
    text = str(value).strip().replace(",", "")
    if "." in text or "%" in text:
        return text
    try:
        return "{:,.2f}".format(int(text))
    except:
        return text

def calc_text_width(text, font, fontsize):
    total = 0.0
    for char in text:
        if char == ',':
            total += COMMA_W
        elif char == '.':
            total += DOT_W
        else:
            try:
                total += font.text_length(char, fontsize=fontsize)
            except:
                total += CHAR_W
    return total

def cover_stamp_area_whole_line(page, x0, x1, y0, y1):
    try:
        cover_rect = fitz.Rect(x0 - 40, y0 - 2, x1 + 2, y1 + 2)
        shape = page.new_shape()
        shape.draw_rect(cover_rect)
        shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
        shape.commit()
        return True
    except Exception as e:
        st.warning(f"白色覆盖失败: {e}")
        return False

def fill_and_render(pdf_bytes, values, dpi=300, fmt="png"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if len(doc) > 1:
        page1 = doc[1]
        try:
            page_dict_str = doc.xref_object(page1.xref)
            contents_match = re.search(r'/Contents\s+(\d+)\s+0\s+R', page_dict_str)
            if contents_match:
                contents_xref = int(contents_match.group(1))
                stream_data = doc.xref_stream(contents_xref)
                content_str = stream_data.decode('latin-1', errors='replace')
                if '[(14403xdzswj)] TJ' in content_str:
                    new_content = content_str.replace('[(14403xdzswj)] TJ', '[()] TJ')
                    new_data = new_content.encode('latin-1')
                    doc.update_stream(contents_xref, new_data, compress=True)
        except Exception as e:
            st.warning(f"内容流替换跳过: {e}")
    font_simsun = None
    font_data = extract_font_data(doc)
    if font_data:
        try:
            font_simsun = fitz.Font(fontbuffer=font_data)
            st.info(f"提取字体成功: {len(font_data)} bytes")
        except Exception as e:
            st.warning(f"提取字体失败: {e}")
    comma_font_path = make_comma_font()
    font_embedded = None
    if comma_font_path:
        try:
            font_embedded = fitz.Font(fontfile=comma_font_path)
            st.info("内嵌字体载入成功（含逗号/星号/中文）")
        except Exception as e:
            st.warning(f"内嵌字体载入失败: {e}")
    use_fallback = (font_simsun is None)
    if use_fallback:
        st.warning("使用备份字体 china-ss")
    try:
        for key, raw_value in values.items():
            if not raw_value or key not in FIELD_CFG:
                continue
            new_value = fmt_decimal(raw_value, key)
            if not new_value:
                continue
            page_num, x0, x1, y0, y1, right_margin = FIELD_CFG[key]
            page = doc[page_num]
            text = str(new_value)
            origin_y = y0 + (y1 - y0) * 0.75
            if key in STAMP_LABELS:
                label = STAMP_LABELS[key]
                full_text = label + text
                cover_stamp_area_whole_line(page, x0, x1, y0, y1)
                write_x = x0 + 2.0
                if use_fallback:
                    page.insert_text((write_x, origin_y), full_text,
                                     fontname="china-ss", fontsize=FONTSIZE, color=(0, 0, 0))
                else:
                    tw = fitz.TextWriter(page.rect)
                    font_to_use = font_embedded if font_embedded else font_simsun
                    current_x = write_x
                    for char in full_text:
                        if char in [',', '.', '*'] and font_embedded is not None:
                            tw.append((current_x, origin_y), char,
                                      fontsize=FONTSIZE, font=font_embedded)
                        else:
                            tw.append((current_x, origin_y), char,
                                      fontsize=FONTSIZE, font=font_to_use)
                        try:
                            char_w = font_to_use.text_length(char, fontsize=FONTSIZE)
                        except:
                            char_w = CHAR_W if char not in [',', '.'] else (COMMA_W if char == ',' else DOT_W)
                        current_x += char_w
                    tw.write_text(page, color=(0, 0, 0))
            else:
                INSET = 0.0
                match_x0 = x0 - 3
                match_x1 = x1 + 3
                match_y0 = y0 - 2
                match_y1 = y1 + 2
                for b in page.get_text("dict")["blocks"]:
                    if "lines" not in b:
                        continue
                    for line in b["lines"]:
                        for span in line["spans"]:
                            sb = span["bbox"]
                            overlap_x = not (sb[2] < match_x0 or sb[0] > match_x1)
                            overlap_y = not (sb[3] < match_y0 or sb[1] > match_y1)
                            if overlap_x and overlap_y:
                                cl = max(sb[0] + INSET, x0 + 1.0)
                                cr = min(sb[2] - INSET, x1 - 1.0)
                                ct = max(sb[1] + INSET, y0 + 1.0)
                                cb = min(sb[3] - INSET, y1 - 1.0)
                                if cr > cl and cb > ct:
                                    rect = fitz.Rect(cl, ct, cr, cb)
                                    shape = page.new_shape()
                                    shape.draw_rect(rect)
                                    shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
                                    shape.commit()
                font_for_calc = font_simsun if font_simsun else (font_embedded if font_embedded else None)
                if font_for_calc and not use_fallback:
                    total_width = calc_text_width(text, font_for_calc, FONTSIZE)
                else:
                    total_width = 0
                    for char in text:
                        if char == ',':
                            total_width += COMMA_W
                        elif char == '.':
                            total_width += DOT_W
                        else:
                            total_width += CHAR_W
                current_x = (x1 - right_margin) - total_width
                if use_fallback:
                    for char in text:
                        page.insert_text((current_x, origin_y), char,
                                         fontname="china-ss", fontsize=FONTSIZE, color=(0, 0, 0))
                        if char == ',':
                            current_x += COMMA_W
                        elif char == '.':
                            current_x += DOT_W
                        else:
                            current_x += CHAR_W
                else:
                    tw = fitz.TextWriter(page.rect)
                    for char in text:
                        if char in [',', '.', '*'] and font_embedded is not None:
                            tw.append((current_x, origin_y), char,
                                      fontsize=FONTSIZE, font=font_embedded)
                        else:
                            tw.append((current_x, origin_y), char,
                                      fontsize=FONTSIZE, font=font_simsun)
                        if char == ',':
                            current_x += COMMA_W
                        elif char == '.':
                            current_x += DOT_W
                        else:
                            try:
                                current_x += font_simsun.text_length(char, fontsize=FONTSIZE)
                            except:
                                current_x += CHAR_W
                    tw.write_text(page, color=(0, 0, 0))
        output_pdf = io.BytesIO()
        doc.save(output_pdf, garbage=0, deflate=False, clean=False)
        output_pdf.seek(0)
        pdf_result = output_pdf.getvalue()
        doc_for_render = fitz.open(stream=pdf_result, filetype="pdf")
        images = []
        for page in doc_for_render:
            pix = page.get_pixmap(dpi=dpi)
            if fmt == "png":
                img_bytes = pix.tobytes("png")
            else:
                img_bytes = pix.tobytes("jpeg")
            images.append(img_bytes)
        doc_for_render.close()
        return images, pdf_result
    finally:
        doc.close()

def main():
    st.title("PDF智能填表系统 v15.1")
    st.markdown("**v15.1 修复版** | 整行覆盖 | 保留原始 PDF 结构 | SimSun 字体")
    st.header("1️⃣ 上传PDF模板")
    uploaded_file = st.file_uploader("选择PDF文件", type=["pdf"])
    if uploaded_file is None:
        st.info("👆 请先上传PDF模板文件")
        return
    pdf_bytes = uploaded_file.getvalue()
    st.header("2️⃣ 填写字段数据")
    st.caption("留空表示不修改。数字自动添加两位小数")
    values = {}
    with st.expander("👥 第1行 - 从业人数", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(总)"],
                                   ["eq1s","eq1e","eq2s","eq2e","eq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)
    with st.expander("💰 第2行 - 资产总额", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(总)"],
                                   ["aq1s","aq1e","aq2s","aq2e","aq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)
    with st.expander("📊 第3-16行（预缴税款计算）", expanded=True):
        col1, col2 = st.columns(2)
        left = [("L1","3 营业收入"),("L2","4 营业成本"),("L3","5 利润总额"),("L4","6 特定业务"),
                ("L5","7 不征税收入"),("L6","8 减：资产加速折旧"),("L7","9 减：免税收入")]
        right = [("L8","10 减：所得减免"),("L9","11 减：所得减免其他"),("L10","12 实际利润额"),
                 ("L11","13 税率(25%)"),("L12","14 应纳所得税额"),("L13","15 减免所得税额"),("L13_1","15.1 减免明细")]
        with col1:
            for k,l in left: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right: values[k] = st.text_input(l, value="", key=k)
    with st.expander("📋 第16-25行"):
        col1, col2 = st.columns(2)
        left2 = [("L14","16 本期预缴"),("L15","17 减免其他"),("L16","18 本期应补(退)"),
                 ("L17","19 总机构本期"),("L18","20 总机构分摊"),("L19","21 财政集中")]
        right2 = [("L20","22 分支机构"),("L21","23 分摊比例"),("L22","24 税率"),
                  ("FZ1","附1 中央级收入"),("FZ2","附2 地方级收入"),("L23","25 减免地方")]
        with col1:
            for k,l in left2: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right2: values[k] = st.text_input(l, value="", key=k)
    with st.expander("📄 第2页 - 补充字段"):
        col1, col2, col3 = st.columns(3)
        with col1: values["L23_2"] = st.text_input("23.2 本年累计应减免", value="", key="L23_2")
        with col2: values["FZ3"] = st.text_input("FZ3 地方级收入实际应纳税额", value="", key="FZ3")
        with col3: values["L24"] = st.text_input("24 实际应补(退)所得税额", value="", key="L24")
    with st.expander("✏️ 第2页 - 签章信息（v15.1 修复：整行覆盖，无残留）"):
        c1, c2 = st.columns(2)
        with c1:
            values["agent_name"] = st.text_input("经办人", value="", key="agent_name",
                help="写入格式：经办人：XXX（标签+值一起写入，覆盖旧文字）")
            values["agent_id"] = st.text_input("经办人身份证号", value="", key="agent_id",
                help="写入格式：经办人身份证号：XXX")
        with c2:
            values["receiver"] = st.text_input("受理人", value="", key="receiver",
                help="写入格式：受理人：XXX（先清除旧 14403xdzswj）")
            values["receive_date"] = st.text_input("受理日期", value="", key="receive_date",
                help="写入格式：受理日期：XXX")
    with st.expander("📎 第3页 - A201020 资产加速折旧附表"):
        st.caption("行1: 加速折旧")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R1C1"] = st.text_input("R1 账载折旧", value="", key="A201_R1C1")
        with c2: values["A201_R1C2"] = st.text_input("R1 税收一般规定", value="", key="A201_R1C2")
        with c3: values["A201_R1C3"] = st.text_input("R1 加速政策计算", value="", key="A201_R1C3")
        with c4: values["A201_R1C4"] = st.text_input("R1 纳税调减金额", value="", key="A201_R1C4")
        with c5: values["A201_R1C5"] = st.text_input("R1 加速优惠金额", value="", key="A201_R1C5")
        st.caption("行2: 一次性扣除")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R2C1"] = st.text_input("R2 账载折旧", value="", key="A201_R2C1")
        with c2: values["A201_R2C2"] = st.text_input("R2 税收一般规定", value="", key="A201_R2C2")
        with c3: values["A201_R2C3"] = st.text_input("R2 加速政策计算", value="", key="A201_R2C3")
        with c4: values["A201_R2C4"] = st.text_input("R2 纳税调减金额", value="", key="A201_R2C4")
        with c5: values["A201_R2C5"] = st.text_input("R2 加速优惠金额", value="", key="A201_R2C5")
        st.caption("行3: 合计")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R3C1"] = st.text_input("R3 账载折旧", value="", key="A201_R3C1")
        with c2: values["A201_R3C2"] = st.text_input("R3 税收一般规定", value="", key="A201_R3C2")
        with c3: values["A201_R3C3"] = st.text_input("R3 加速政策计算", value="", key="A201_R3C3")
        with c4: values["A201_R3C4"] = st.text_input("R3 纳税调减金额", value="", key="A201_R3C4")
        with c5: values["A201_R3C5"] = st.text_input("R3 加速优惠金额", value="", key="A201_R3C5")
    st.header("3️⃣ 输出设置")
    c1, c2 = st.columns(2)
    with c1:
        output_fmt = st.radio("图片格式", ["PNG（无损推荐）", "JPG"], horizontal=True)
    with c2:
        dpi = st.select_slider("DPI", options=[150, 200, 300, 400], value=300)
    fmt = "png" if "PNG" in output_fmt else "jpeg"
    ext = "png" if fmt == "png" else "jpg"
    filled = {k: v.strip() for k, v in values.items() if v.strip()}
    if filled:
        st.success(f"✅ 已填写 {len(filled)} 个字段")
    else:
        st.info("💡 尚未填写任何字段")
    if st.button("🚀 生成图片", type="primary", disabled=len(filled)==0):
        with st.spinner("正在渲染..."):
            try:
                images, pdf_data = fill_and_render(pdf_bytes, filled, dpi=dpi, fmt=fmt)
                st.success(f"✅ 成功生成 {len(images)} 页图片（{dpi} DPI）")
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    if len(images) == 1:
                        st.download_button(f"📥 下载图片（{ext.upper()}）", images[0],
                            f"filled_tax_form_page1.{ext}", mime=f"image/{ext}", use_container_width=True)
                    else:
                        zip_buf = io.BytesIO()
                        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                            for i, img_data in enumerate(images):
                                zf.writestr(f"filled_tax_form_page{i+1}.{ext}", img_data)
                        zip_buf.seek(0)
                        st.download_button(f"📥 下载全部图片（ZIP）", zip_buf.getvalue(),
                            "filled_tax_form.zip", mime="application/zip", use_container_width=True)
                with col_dl2:
                    st.download_button("📥 下载 PDF（保留原始结构）", pdf_data,
                        "filled_tax_form.pdf", mime="application/pdf", use_container_width=True)
                st.subheader("预览（第1页）")
                st.image(images[0], use_container_width=True)
            except Exception as e:
                st.error(f"❌ 生成失败: {e}")
                st.exception(e)
    st.markdown("---")
    st.markdown("<center>PDF智能填表系统 v15.1 | 整行覆盖修复 | 保留原始PDF结构 | SimSun 字体</center>",
                unsafe_allow_html=True)

if __name__ == "__main__":
    main()
