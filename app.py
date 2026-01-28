import streamlit as st
import sys
import os

# === 1. ページ設定 ===
st.set_page_config(page_title="Universe PDF Converter (PyMuPDF)", layout="wide")

# UIスタイル
st.markdown("""
<style>
    .stButton>button { border-radius: 8px; font-weight: bold; border: 2px solid #4CAF50; }
    h1 { color: #2E7D32; }
</style>
""", unsafe_allow_html=True)

# === 2. ライブラリ読み込みチェック ===
try:
    import cv2
    import numpy as np
    import pytesseract
    from pytesseract import Output
    import fitz  # ★ここが新エンジン (PyMuPDF)
    from PIL import Image
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR
    import io
    import zipfile
    from streamlit_image_comparison import image_comparison
    from deep_translator import GoogleTranslator
except ImportError as e:
    st.error(f"⚠️ ライブラリ不足: {e}")
    st.info("requirements.txt に 'pymupdf' が含まれているか確認してください。")
    st.stop()

st.title("🪐 Universe PDF Converter (脱Poppler版)")

# ===========================
# 関数定義 (PyMuPDFエンジン)
# ===========================

@st.cache_data(show_spinner=False)
def load_pdf_images(file_bytes):
    """
    ★Popplerを使わず、PyMuPDFでPDFを画像化する（エラー知らず！）
    """
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        images = []
        for page in doc:
            # 解像度設定 (zoom=2 で 144dpi相当, zoom=3 で 216dpi相当)
            # OCR精度のため少し高めに設定
            zoom = 2.0  
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Pixmap -> PIL Image変換
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)
        return images
    except Exception as e:
        st.error(f"PDF読み込みエラー: {e}")
        return []

@st.cache_data(show_spinner=False)
def get_pdf_info_simple(file_bytes):
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        return {"Pages": len(doc)}
    except:
        return {"Pages": 0}

@st.cache_data(show_spinner=False)
def translate_text(text, target_lang='ja'):
    try:
        if not text.strip(): return text
        translator = GoogleTranslator(source='auto', target=target_lang)
        return translator.translate(text)
    except: return text

def preprocess_image_for_ocr(cv_img, zoom_factor=1.0):
    # PyMuPDFですでに拡大しているので、ここのZoomは控えめでOK
    h, w = cv_img.shape[:2]
    new_w, new_h = int(w * zoom_factor), int(h * zoom_factor)
    resized = cv2.resize(cv_img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    enhanced = clahe.apply(gray)
    _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary, new_w, new_h

def extract_objects(cv_img, min_area=5000):
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (10, 10))
    dilated = cv2.dilate(thresh, kernel, iterations=3)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    objects = []
    h, w = cv_img.shape[:2]
    for cnt in contours:
        x, y, cw, ch = cv2.boundingRect(cnt)
        if cw * ch < min_area: continue
        if cw > w * 0.95 or ch > h * 0.95: continue 
        objects.append((x, y, cw, ch))
    return objects

def sort_blocks_smart(blocks):
    if not blocks: return []
    block_list = []
    for bid, data in blocks.items():
        data['id'] = bid
        data['avg_x'] = min(data['left'])
        data['avg_y'] = min(data['top'])
        block_list.append(data)
    if not block_list: return []
    max_x = max([b['avg_x'] + max(b['width']) for b in block_list])
    center_x = max_x / 2
    left_col = [b for b in block_list if b['avg_x'] < center_x]
    right_col = [b for b in block_list if b['avg_x'] >= center_x]
    left_col.sort(key=lambda x: x['avg_y'])
    right_col.sort(key=lambda x: x['avg_y'])
    return left_col + right_col

def merge_text_blocks_ordered(sorted_blocks, spacing_limit=50):
    if not sorted_blocks: return {}
    merged = {}
    curr_m_id = 0
    merged[curr_m_id] = {k: v[:] if isinstance(v, list) else v for k, v in sorted_blocks[0].items()}
    for i in range(1, len(sorted_blocks)):
        curr = sorted_blocks[i]
        prev = merged[curr_m_id]
        prev_bottom = max([t + h for t, h in zip(prev['top'], prev['height'])])
        curr_top = min(curr['top'])
        prev_left = min(prev['left'])
        curr_left = min(curr['left'])
        if 0 < (curr_top - prev_bottom) < spacing_limit and abs(curr_left - prev_left) < 60:
            prev['text'].append("\n")
            prev['text'].extend(curr['text'])
            prev['left'].extend(curr['left'])
            prev['top'].extend(curr['top'])
            prev['width'].extend(curr['width'])
            prev['height'].extend(curr['height'])
        else:
            curr_m_id += 1
            merged[curr_m_id] = {k: v[:] if isinstance(v, list) else v for k, v in curr.items()}
    return merged

# ===========================
# サイドバー & メイン処理
# ===========================
st.sidebar.header("🎛️ 設定パネル")

with st.sidebar.expander("🌐 言語・翻訳"):
    do_translate = st.checkbox("自動翻訳", value=False)
    target_lang = st.selectbox("翻訳先", ["ja (日本語)", "en (英語)", "zh-CN (中国語)"])
    target_lang_code = target_lang.split()[0]
    ocr_lang = "jpn+eng"

with st.sidebar.expander("🛠️ デザイン・モード", expanded=True):
    mode = st.radio("モード", ["分解モード", "通常モード"])
    target_font = st.selectbox("フォント", ["Meiryo", "Yu Gothic", "Arial"])
    detect_title = st.checkbox("タイトル強調", value=True)

with st.sidebar.expander("⚙️ 詳細設定"):
    use_erase = st.checkbox("ロゴ消し", value=True)
    erase_w = st.slider("右端カット", 0, 800, 350)
    erase_h = st.slider("下端カット", 0, 500, 180)
    min_area = st.slider("図形感度", 1000, 20000, 5000)
    smart_sort = st.checkbox("2段組み補正", value=True)
    merge_lines = st.checkbox("段落結合", value=True)

uploaded_file = st.file_uploader("PDFファイルをアップロード", type="pdf")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    
    # PyMuPDFで読み込み
    info = get_pdf_info_simple(file_bytes)
    total_pages = info["Pages"]
    
    with st.spinner("PDFエンジン(PyMuPDF)で読み込み中..."):
        images = load_pdf_images(file_bytes)
        if not images: st.stop()

    # --- プレビュー ---
    st.divider()
    preview_idx = st.number_input("プレビュー", 1, total_pages, 1) - 1
    
    img_orig = np.array(images[preview_idx])
    img_proc = cv2.cvtColor(np.array(images[preview_idx]), cv2.COLOR_RGB2BGR)
    h, w = img_proc.shape[:2]

    if use_erase:
        mask = np.zeros((h, w), np.uint8)
        cv2.rectangle(mask, (w - erase_w, h - erase_h), (w, h), 255, -1)
        img_proc = cv2.inpaint(img_proc, mask, 3, cv2.INPAINT_TELEA)

    if mode.startswith("分解"):
        objs = extract_objects(img_proc, min_area)
        for (ox, oy, ow, oh) in objs:
            cv2.rectangle(img_proc, (ox, oy), (ox+ow, oy+oh), (255, 120, 0), 4)

    if detect_title:
        cv2.line(img_proc, (0, int(h*0.2)), (w, int(h*0.2)), (0, 255, 0), 2)

    image_comparison(
        img1=img_orig,
        img2=cv2.cvtColor(img_proc, cv2.COLOR_BGR2RGB),
        label1="Original",
        label2="Processed",
        width=700,
        in_memory=True
    )

    # --- 変換 ---
    st.divider()
    c1, c2 = st.columns([2, 1])
    with c1:
        page_range = st.slider("範囲", 1, total_pages, (1, min(total_pages, 5)))
    with c2:
        st.write("")
        btn = st.button("🚀 変換スタート", type="primary", use_container_width=True)

    if btn:
        start_p, end_p = page_range
        process_cnt = end_p - start_p + 1
        p_bar = st.progress(0)
        status_area = st.empty() # 安定化のためmarkdown固定エリア

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            target_images = images[start_p-1 : end_p]
            
            for i, image in enumerate(target_images):
                p_num = start_p + i
                status_area.markdown(f"**Processing... {i+1}/{process_cnt} (P.{p_num})**")
                
                cv_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                h_orig, w_orig = cv_img.shape[:2]
                scale_x = prs.slide_width / w_orig
                scale_y = prs.slide_height / h_orig
                
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                if use_erase:
                    mask = np.zeros((h_orig, w_orig), np.uint8)
                    cv2.rectangle(mask, (w_orig - erase_w, h_orig - erase_h), (w_orig, h_orig), 255, -1)
                    cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

                object_rects = []
                if mode.startswith("分解"):
                    objs = extract_objects(cv_img, min_area)
                    for idx, (ox, oy, ow, oh) in enumerate(objs):
                        crop = cv_img[oy:oy+oh, ox:ox+ow]
                        zf.writestr(f"assets/P{p_num}_fig{idx}.jpg", cv2.imencode(".jpg", crop)[1].tobytes())
                        s = io.BytesIO(cv2.imencode(".jpg", crop)[1].tobytes())
                        slide.shapes.add_picture(s, int(ox*scale_x), int(oy*scale_y), width=int(ow*scale_x), height=int(oh*scale_y))
                        object_rects.append((ox, oy, ow, oh))
                else:
                    s = io.BytesIO(cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), 80])[1].tofile(s))
                    slide.shapes.add_picture(s, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # OCR (Tesseract)
                # Tesseractはpackages.txtが必要ですが、PyMuPDFへの移行で少なくとも画像化までは成功します
                # 万が一OCRが動かなくてもアプリは落ちないようにtry-catch
                try:
                    ocr_img, _, _ = preprocess_image_for_ocr(cv_img, 1.0) # PyMuPDFですでに高画質なのでzoom=1
                    d = pytesseract.image_to_data(ocr_img, lang=ocr_lang, output_type=Output.DICT)
                except:
                    d = {'text': [], 'conf': []}

                raw_blocks = {}
                if 'text' in d:
                    for j in range(len(d['text'])):
                        txt = d['text'][j].strip()
                        if int(d['conf'][j]) > 40 and txt != "" and not (len(txt)==1 and txt in ".,-_|"):
                            bid = d['block_num'][j]
                            if bid not in raw_blocks: raw_blocks[bid] = {'text':[], 'left':[], 'top':[], 'width':[], 'height':[]}
                            raw_blocks[bid]['text'].append(txt)
                            raw_blocks[bid]['left'].append(d['left'][j])
                            raw_blocks[bid]['top'].append(d['top'][j])
                            raw_blocks[bid]['width'].append(d['width'][j])
                            raw_blocks[bid]['height'].append(d['height'][j])
                
                sorted_list = sort_blocks_smart(raw_blocks) if smart_sort else sorted(raw_blocks.values(), key=lambda x: min(x['top']))
                final_blocks = merge_text_blocks_ordered(sorted_list) if merge_lines else {i:v for i,v in enumerate(sorted_list)}

                notes_content = []
                for bid, bdata in final_blocks.items():
                    original_text = "".join(bdata['text'])
                    display_text = original_text
                    if do_translate:
                        translated = translate_text(original_text, target_lang_code)
                        display_text = translated
                        notes_content.append(f"[原] {original_text}\n[訳] {translated}\n")
                    else:
                        notes_content.append(original_text + "\n")

                    ox = min(bdata['left'])
                    oy = min(bdata['top'])
                    ow = max([l+w for l,w in zip(bdata['left'], bdata['width'])]) - ox
                    # PyMuPDFの画像はOCR用とサイズ一致していると仮定(zoom調整済み)
                    orig_x, orig_y, orig_w = ox, oy, ow 

                    if mode.startswith("分解"):
                        cx, cy = orig_x + orig_w/2, orig_y + int(max(bdata['height'])/2)
                        if any(rox < cx < rox+row and roy < cy < roy+roh for (rox,roy,row,roh) in object_rects): continue

                    pp_x = int(orig_x * scale_x)
                    pp_y = int(orig_y * scale_y)
                    pp_w = int(orig_w * scale_x)

                    if pp_w > Inches(0.2):
                        try:
                            is_title = detect_title and (orig_y < h_orig * 0.2)
                            avg_h = sum(bdata['height'])/len(bdata['height'])
                            est_pt = (prs.slide_height.inches*72)*(avg_h/h_orig)*0.75
                            font_pt = max(9, min(est_pt, 80))
                            if is_title: font_pt *= 1.2
                            
                            box_h = Inches(font_pt/72 * 1.5 * (display_text.count('\n') + 1))
                            txBox = slide.shapes.add_textbox(pp_x, pp_y, pp_w, box_h)
                            tf = txBox.text_frame
                            tf.word_wrap = True
                            p = tf.paragraphs[0]
                            p.text = display_text
                            p.font.size = Pt(font_pt)
                            p.font.name = target_font
                            p.font.language_id = 1041
                            
                            if is_title:
                                p.font.bold = True
                                p.font.color.rgb = RGBColor(0, 51, 102)
                            else:
                                y1,y2 = max(0,orig_y), min(h_orig,orig_y+int(max(bdata['height'])))
                                x1,x2 = max(0,orig_x), min(w_orig,orig_x+orig_w)
                                roi = cv_img[y1:y2, x1:x2]
                                if roi.size > 0:
                                    pixels = np.float32(roi.reshape(-1, 3))
                                    mask = np.all(pixels < 230, axis=1)
                                    t_pix = pixels[mask]
                                    if len(t_pix)>0:
                                        r,g,b = np.mean(t_pix, axis=0)
                                        p.font.color.rgb = RGBColor(int(r), int(g), int(b))
                        except: pass
                
                slide.notes_slide.notes_text_frame.text = "\n".join(notes_content)
                p_bar.progress((i+1)/process_cnt)

        status_area.success("✅ 変換完了！")
        st.balloons()
        
        col_d1, col_d2 = st.columns(2)
        out_ppt = io.BytesIO()
        prs.save(out_ppt)
        out_ppt.seek(0)
        col_d1.download_button(f"📥 スライド ({target_lang})", out_ppt, "Converted_Slides.pptx", type="primary", use_container_width=True)

        if mode.startswith("分解"):
            zip_buffer.seek(0)
            col_d2.download_button(f"🗂️ 画像素材", zip_buffer, "assets.zip", use_container_width=True)
