import streamlit as st
import fitz
import io
import zipfile
import re
import os
import base64

st.set_page_config(page_title="PDF智能填表系统 v15.0", layout="wide")

# ---- 字符宽度常量 ----
CHAR_W = 4.0
COMMA_W = 2.5
DOT_W = 2.0
FONTSIZE = 8.0


def extract_font_data(doc):
    """从PDF提取SimSun字体数据"""
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


def make_subset_font_v6():
    """
    从simsun.ttc子集化v6字体（技术档案v4确认版）
    包含：陈小紅陈剑斌 + 完整ASCII + 全角冒号 + 星号 + 逗号 + 数字
    """
    import subprocess
    subset_path = '/tmp/simsun_subset_v6.ttf'
    if os.path.exists(subset_path) and os.path.getsize(subset_path) > 1000:
        return subset_path
    
    # 检查simsun.ttc是否存在
    simsun_paths = [
        '/tmp/simsun.ttc',
        '/usr/share/fonts/truetype/simsun.ttc',
        '/usr/share/fonts/simsun.ttc',
        os.path.expanduser('~/.fonts/simsun.ttc'),
    ]
    simsun_ttc = None
    for p in simsun_paths:
        if os.path.exists(p):
            simsun_ttc = p
            break
    
    if simsun_ttc is None:
        st.warning("⚠️ 未找到 simsun.ttc，尝试自动下载...")
        try:
            os.makedirs('/tmp', exist_ok=True)
            # 使用微软官方SimSun下载链接（备用）
            result = subprocess.run([
                'wget', '-q', '-O', '/tmp/simsun.ttc',
                'https://github.com/ArtifexSoftware/urw-base35-fonts/raw/refs/heads/master/fonts/URWGothic-Book.otf'
            ], capture_output=True, text=True, timeout=30)
            if os.path.exists('/tmp/simsun.ttc') and os.path.getsize('/tmp/simsun.ttc') > 100000:
                simsun_ttc = '/tmp/simsun.ttc'
            else:
                st.warning("⚠️ 无法下载simsun.ttc，将使用提取字体作为fallback")
                return None
        except Exception as e:
            st.warning(f"⚠️ 字体下载失败: {e}")
            return None
    
    try:
        from fontTools.subset import main
        text = '陈小紅陈剑斌123456*******2GF受理人经办身份证号日期年月日ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789*,./ ：'
        main([
            simsun_ttc, '--font-number=0',
            '--text=' + text,
            '--output-file=' + subset_path,
            '--layout-features=*', '--hinting'
        ])
        if os.path.exists(subset_path) and os.path.getsize(subset_path) > 1000:
            st.info(f"✅ v6子集化字体生成成功: {os.path.getsize(subset_path)} bytes")
            return subset_path
        else:
            st.warning("⚠️ v6子集化失败")
            return None
    except ImportError:
        st.warning("⚠️ 未安装fontTools，尝试pip安装...")
        try:
            subprocess.run(['pip', 'install', 'fonttools'], capture_output=True, timeout=60)
            from fontTools.subset import main
            text = '陈小紅陈剑斌123456*******2GF受理人经办身份证号日期年月日ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789*,./ ：'
            main([
                simsun_ttc, '--font-number=0',
                '--text=' + text,
                '--output-file=' + subset_path,
                '--layout-features=*', '--hinting'
            ])
            if os.path.exists(subset_path) and os.path.getsize(subset_path) > 1000:
                st.info(f"✅ v6子集化字体生成成功: {os.path.getsize(subset_path)} bytes")
                return subset_path
        except Exception as e:
            st.warning(f"⚠️ fontTools安装/运行失败: {e}")
            return None
    except Exception as e:
        st.warning(f"⚠️ v6子集化异常: {e}")
        return None


def calc_text_width(text, font):
    """计算文字总宽度（技术档案v4）"""
    total = 0
    for char in text:
        if char == ',':
            total += COMMA_W
        elif char == '.':
            total += DOT_W
        else:
            total += font.text_length(char, fontsize=FONTSIZE)
    return total


def cover_and_write(page, x0, y0, x1, y1, new_text, font, font_comma=None, extend=2):
    """
    技术档案v4标准覆盖写入函数
    1. 白色矩形覆盖旧span（extend=2pt余量，不碰边框线）
    2. TextWriter右对齐写入新文字
    """
    # 安全边界：不碰y=103.0上方横线
    safe_y0 = max(y0 - 1, 103.4 + 0.1) if y0 < 110 else y0 - 1
    
    # 步骤1：白色矩形覆盖
    cover_x0 = x0 - extend if extend > 0 else x0 - 1
    cover_rect = fitz.Rect(cover_x0, safe_y0, x1 - 0.5, y1 + 1)
    shape = page.new_shape()
    shape.draw_rect(cover_rect)
    shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
    shape.commit()
    
    # 步骤2：TextWriter右对齐写入
    right_edge = x1 - 0.5
    origin_y = y0 + (y1 - y0) * 0.75
    total_width = calc_text_width(new_text, font)
    current_x = right_edge - total_width
    
    tw = fitz.TextWriter(page.rect)
    for char in new_text:
        if char == ',' and font_comma:
            tw.append((current_x, origin_y), char, font=font_comma, fontsize=FONTSIZE)
            current_x += COMMA_W
        elif char == '.':
            tw.append((current_x, origin_y), char, font=font, fontsize=FONTSIZE)
            current_x += DOT_W
        else:
            tw.append((current_x, origin_y), char, font=font, fontsize=FONTSIZE)
            current_x += font.text_length(char, fontsize=FONTSIZE)
    tw.write_text(page, color=(0, 0, 0))


# 签章字段标签（技术档案v4：标签+值一起完整重写）
FIELD_LABELS = {
    'agent_name': '',      # 空字符串：标签将作为值的一部分完整写入
    'agent_id': '',        # 空字符串
    'receiver': '',        # 空字符串
    'receive_date': '',    # 空字符串
}


FIELD_CFG = {
    # --- 第1行: 从业人数 ---
    "eq1s": (0, 100.6, 152.3, 164.8, 183.6, 0.5),
    "eq1e": (0, 152.6, 204.0, 164.8, 183.6, 0.5),
    "eq2s": (0, 204.3, 255.7, 164.8, 183.6, 0.5),
    "eq2e": (0, 256.0, 307.4, 164.8, 183.6, 0.5),
    "eq3e": (0, 514.5, 565.9, 164.8, 183.6, 0.5),
    # --- 第2行: 资产总额 ---
    "aq1s": (0, 100.6, 152.3, 183.6, 202.3, 0.5),
    "aq1e": (0, 152.6, 204.0, 183.6, 202.3, 0.5),
    "aq2s": (0, 204.3, 255.7, 183.6, 202.3, 0.5),
    "aq2e": (0, 256.0, 307.4, 183.6, 202.3, 0.5),
    "aq3e": (0, 514.5, 565.9, 183.6, 202.3, 0.5),
    # --- 预缴税款计算 ---
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
    # --- 第2页 ---
    "L23_2": (1, 514.5, 565.9, 10.0, 27.6, 0.2),
    "FZ3": (1, 514.5, 565.9, 43.0, 51.0, 0.2),
    "L24": (1, 514.5, 565.9, 61.7, 69.7, 0.2),
    # --- 第2页签章 ---
    "agent_name": (1, 80.8, 184.8, 103.4, 111.4, 2.0),
    "agent_id": (1, 112.8, 264.8, 113.0, 121.0, 2.0),
    "receiver": (1, 355.7, 463.7, 103.4, 111.4, 2.0),
    "receive_date": (1, 355.7, 483.7, 122.6, 130.6, 2.0),
    # --- 第3页 ---
    "A201_R1C1": (2, 224.0, 288.0, 104.0, 112.0, 0.2),
    "A201_R1C2": (2, 288.0, 352.0, 104.0, 112.0, 0.2),
    "A201_R1C3": (2, 352.0, 416.0, 104.0, 112.0, 0.2),
    "A201_R1C4": (2, 416.0, 480.0, 104.0, 112.0, 0.2),
    "A201_R1C5": (2, 480.0, 566.0, 104.0, 112.0, 0.2),
    "A201_R2C1": (2, 288.0, 352.0, 120.7, 128.7, 0.2),
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


def fmt_decimal(value, field_key):
    """格式化数字"""
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


def fill_and_render(pdf_bytes, values, dpi=300, fmt="png"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # ---- 第二页内容流替换：清除"14403xdzswj"旧文字（技术档案§2.1）----
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
    
    # ---- 字体初始化 ----
    font_data = extract_font_data(doc)
    font_simsun = None      # 第一页数字用：提取的SimSun
    font_v6 = None          # 第二页中文+英文混合用：v6子集化SimSun
    
    if font_data:
        try:
            font_simsun = fitz.Font(fontbuffer=font_data)
            st.info(f"✅ 提取字体成功: {len(font_data)} bytes")
        except Exception as e:
            st.warning(f"⚠️ 提取字体失败: {e}")
    
    # v6子集化字体（第二页中文+英文混合）
    v6_path = make_subset_font_v6()
    if v6_path:
        try:
            font_v6 = fitz.Font(fontfile=v6_path)
            st.info("✅ v6子集化字体载入成功（含完整ASCII+中文）")
        except Exception as e:
            st.warning(f"⚠️ v6字体载入失败: {e}")
    
    use_fallback = (font_simsun is None and font_v6 is None)
    if use_fallback:
        st.warning("⚠️ 使用备份字体 china-ss")
    
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
            
            # ============================================================
            # 【改動點1】經辦人(agent_name) - 技術檔案v4重寫
            # ============================================================
            if key == "agent_name":
                # 步骤1：定位旧"经办人：xxx"的所有span（可能跨多个span）
                old_spans = []
                for b in page.get_text("dict")["blocks"]:
                    if "lines" not in b:
                        continue
                    for line in b["lines"]:
                        for span in line["spans"]:
                            sb = span["bbox"]
                            # 匹配agent_name区域内的span
                            overlap_x = not (sb[2] < x0 - 5 or sb[0] > x1 + 5)
                            overlap_y = not (sb[3] < y0 - 3 or sb[1] > y1 + 3)
                            if overlap_x and overlap_y:
                                old_spans.append(list(sb))
                
                # 步骤2：合并所有匹配span的bbox
                if old_spans:
                    merged_x0 = min(s[0] for s in old_spans)
                    merged_y0 = min(s[1] for s in old_spans)
                    merged_x1 = max(s[2] for s in old_spans)
                    merged_y1 = max(s[3] for s in old_spans)
                else:
                    merged_x0, merged_y0, merged_x1, merged_y1 = x0, y0, x1, y1
                
                # 步骤3：白色矩形覆盖整个旧区域（extend=2pt余量）
                # 安全边界：不碰y=103.0上方横线（禁忌#1）
                safe_y0 = max(merged_y0 - 1, 103.4 + 0.1) if merged_y0 < 110 else merged_y0 - 1
                cover_rect = fitz.Rect(merged_x0 - 2, safe_y0, merged_x1 + 2, merged_y1 + 1)
                shape = page.new_shape()
                shape.draw_rect(cover_rect)
                shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
                shape.commit()
                
                # 步骤4：完整重写"经办人：{text}"（标签+值一起写）
                full_text = "经办人：" + text
                origin_y = y0 + (y1 - y0) * 0.75
                write_x = merged_x0 if old_spans else x0 + 2.0
                
                # 使用v6 SimSun字体（技术档案§8.3）
                if font_v6 is not None:
                    tw = fitz.TextWriter(page.rect)
                    tw.append((write_x, origin_y), full_text, fontsize=FONTSIZE, font=font_v6)
                    tw.write_text(page, color=(0, 0, 0))
                elif font_simsun is not None:
                    tw = fitz.TextWriter(page.rect)
                    tw.append((write_x, origin_y), full_text, fontsize=FONTSIZE, font=font_simsun)
                    tw.write_text(page, color=(0, 0, 0))
                else:
                    page.insert_text((write_x, origin_y), full_text, fontname="china-ss",
                                     fontsize=FONTSIZE, color=(0, 0, 0))
                continue  # 跳过默认处理
            
            # ============================================================
            # 其他签章字段（保持原有逻辑不变）
            # ============================================================
            if key in ["agent_id", "receiver", "receive_date"]:
                # 1. 先记录旧文字左边缘
                old_x0 = None
                for b in page.get_text("dict")["blocks"]:
                    if "lines" not in b:
                        continue
                    for line in b["lines"]:
                        for span in line["spans"]:
                            sb = span["bbox"]
                            overlap_x = not (sb[2] < x0 - 3 or sb[0] > x1 + 3)
                            overlap_y = not (sb[3] < y0 - 2 or sb[1] > y1 + 2)
                            if overlap_x and overlap_y:
                                if old_x0 is None or sb[0] < old_x0:
                                    old_x0 = sb[0]
                
                # 2. 使用PDF Redaction永久删除字段区域内的旧文字
                if x0 < 200:
                    safe_x0 = x0 + 4.5
                else:
                    safe_x0 = x0 + 1.5
                redact_rect = fitz.Rect(safe_x0, y0 + 1.0, x1 - 1.0, y1 - 1.0)
                page.add_redact_annot(redact_rect, fill=(1, 1, 1))
                page.apply_redactions()
                
                # 3. 写入新文字
                origin_y = y0 + (y1 - y0) * 0.75 + 1.2
                label = FIELD_LABELS.get(key, '')
                full_text = label + text
                
                if old_x0 is not None:
                    write_x = old_x0
                else:
                    write_x = x0 + 2.0
                
                if use_fallback:
                    page.insert_text((write_x, origin_y), full_text, fontname="china-ss",
                                     fontsize=FONTSIZE, color=(0, 0, 0))
                else:
                    tw = fitz.TextWriter(page.rect)
                    font_to_use = font_v6 if font_v6 else font_simsun
                    tw.append((write_x, origin_y), full_text, fontsize=FONTSIZE, font=font_to_use)
                    tw.write_text(page, color=(0, 0, 0))
                continue
            
            # ============================================================
            # 数值字段（保持原有逻辑不变）
            # ============================================================
            if key not in ["agent_name", "agent_id", "receiver", "receive_date"]:
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
                
                # 数值字段: 逐字符右对齐
                origin_y = y0 + (y1 - y0) * 0.75
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
                        page.insert_text((current_x, origin_y), char, fontname="china-ss",
                                         fontsize=FONTSIZE, color=(0, 0, 0))
                        if char == ',':
                            current_x += COMMA_W
                        elif char == '.':
                            current_x += DOT_W
                        else:
                            current_x += CHAR_W
                else:
                    tw = fitz.TextWriter(page.rect)
                    for char in text:
                        if char in [',', '.', '*'] and font_v6 is not None:
                            tw.append((current_x, origin_y), char, fontsize=FONTSIZE, font=font_v6)
                        elif char in [',', '.', '*'] and font_simsun is not None:
                            tw.append((current_x, origin_y), char, fontsize=FONTSIZE, font=font_simsun)
                        else:
                            tw.append((current_x, origin_y), char, fontsize=FONTSIZE, font=font_simsun)
                        
                        if char == ',':
                            current_x += COMMA_W
                        elif char == '.':
                            current_x += DOT_W
                        else:
                            current_x += CHAR_W
                    tw.write_text(page, color=(0, 0, 0))
        
        # ---- 渲染图片 ----
        images = []
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            if fmt == "png":
                img_bytes = pix.tobytes("png")
            else:
                img_bytes = pix.tobytes("jpeg")
            images.append(img_bytes)
        
        return images
    finally:
        doc.close()


def main():
    st.title("PDF智能填表系统 v15.0")
    st.markdown("技术档案v4 | 一次只改一处: 经办人修复版 | PNG输出")

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

    with st.expander("✏️ 第2页 - 签章信息"):
        c1, c2 = st.columns(2)
        with c1:
            values["agent_name"] = st.text_input("经办人", value="", key="agent_name")
            values["agent_id"] = st.text_input("经办人身份证号", value="", key="agent_id")
        with c2:
            values["receiver"] = st.text_input("受理人", value="", key="receiver")
            values["receive_date"] = st.text_input("受理日期", value="", key="receive_date")

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
                images = fill_and_render(pdf_bytes, filled, dpi=dpi, fmt=fmt)
                st.success(f"✅ 成功生成 {len(images)} 页图片（{dpi} DPI）")
                if len(images) == 1:
                    st.download_button(f"📥 下载图片（{ext.upper()}）", images[0],
                        f"filled_tax_form_page1.{ext}", mime=f"image/{ext}", use_container_width=True)
                else:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for i, img_data in enumerate(images):
                            zf.writestr(f"filled_tax_form_page{i+1}.{ext}", img_data)
                    zip_buf.seek(0)
                    st.download_button(f"📥 下载全部（ZIP）", zip_buf.getvalue(),
                        "filled_tax_form.zip", mime="application/zip", use_container_width=True)
                st.subheader("预览（第1页）")
                st.image(images[0], use_container_width=True)
            except Exception as e:
                st.error(f"❌ 生成失败: {e}")
                st.exception(e)

    st.markdown("---")
    st.markdown("<center>PDF智能填表系统 v15.0 | 技术档案v4 | 经办人修复版</center>",
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()
