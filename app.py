import streamlit as st
import os
import cv2
import numpy as np
import pytesseract
from pytesseract import Output
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

# === ページ設定 ===
st.set_page_config(page_title="God Mode PDF Converter", layout="wide", initial_sidebar_state="expanded")

st.title("⚡ 真・完全版 PDF変換ツール (God Mode)")
st.markdown("""
**最終進化:**
1. **リアルタイムプレビュー**: 設定変更がその場で目に見えます。
2. **段落結合AI**: バラバラの行をひとつのテキストボックスにまとめます。
3. **世界最高峰のOCR**: ノイズ除去と配置精度を極限まで高めました。
""")

# ===========================
# サイドバー設定
# ===========================
st.sidebar.header("🎛️ コントロールパネル")

tab1, tab2, tab3 = st.sidebar.tabs(["調整＆プレビュー", "高度設定", "出力"])

with tab1:
    st.info("👇 ここを動かすと右の画像が変わります")
    # ロゴ消し
    use_erase = st.checkbox("ロゴ/不要領域を消す", value=True)
    erase_width = st.slider("右端カット (px)", 0, 800, 350, step=10)
    erase_height = st.slider("下端カット (px)", 0, 500, 180, step=10)
    
    st.divider()
    st.write("**プレビュー用ページ**")
    preview_page = st.number_input("確認するページ番号", min_value=1, value=1)

with tab2:
    mode = st.radio("処理モード", ["分解モード（図と文字を分離）", "通常モード（背景一枚絵）"])
    min_area_size = st.slider("図形認識サイズ", 1000, 20000, 5000)
    # 新機能：段落結合
    merge_lines = st.checkbox("段落結合 (バラバラの行をまとめる)", value=True, help="近い行を一つのテキストボックスにします")
    line_spacing_limit = st.slider("行間許容値 (px)", 10, 100, 40, help="この距離以内の行は結合します")

with tab3:
    template_file = st.file_uploader("PPTXテンプレート (任意)", type="pptx")
    jpeg_quality = st.slider("画像品質", 10, 100, 85)

# ===========================
# メインエリア：PDFアップロード
# ===========================
uploaded_file = st.file_uploader("PDFをアップロードしてください", type="pdf")

# ===========================
# 関数定義
# ===========================
def preprocess_image_for_ocr(cv_img, zoom_factor=2.0):
    h, w = cv_img.shape[:2]
    new_w, new_h = int(w * zoom_factor), int(h * zoom_factor)
    resized = cv2.resize(cv_img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    enhanced = clahe.apply(gray)
    _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary, new_w, new_h

def get_dominant_color(img_crop):
    try:
        pixels = np.float32(img_crop.reshape(-1, 3))
        mask = np.all(pixels < 230, axis=1)
        target_pixels = pixels[mask]
        if len(target_pixels) > 0:
            avg = np.mean(target_pixels, axis=0)
            return int(avg[0]), int(avg[1]), int(avg[2])
        return 0, 0, 0
    except: return 0, 0, 0

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
        if cw > w * 0.95 and ch > h * 0.95: continue
        objects.append((x, y, cw, ch))
    return objects

def merge_text_blocks(blocks, spacing_limit=40):
    """
    バラバラのテキストブロックを、位置関係に基づいて結合する関数
    """
    if not blocks: return {}
    
    # Y座標（上からの位置）でソート
    sorted_ids = sorted(blocks.keys(), key=lambda k: min(blocks[k]['top']))
    
    merged_blocks = {}
    current_merge_id = 0
    
    # 最初のブロックを登録
    last_id = sorted_ids[0]
    merged_blocks[current_merge_id] = blocks[last_id].copy()
    
    for i in range(1, len(sorted_ids)):
        curr_id = sorted_ids[i]
        curr = blocks[curr_id]
        prev = merged_blocks[current_merge_id]
        
        # 判定基準:
        # 1. 左端の位置が近い (インデントが同じ)
        # 2. 上のブロックの下端と、今のブロックの上端が近い (行間)
        
        prev_bottom = max([t + h for t, h in zip(prev['top'], prev['height'])])
        curr_top = min(curr['top'])
        prev_left = min(prev['left'])
        curr_left = min(curr['left'])
        
        vertical_dist = curr_top - prev_bottom
        horizontal_diff = abs(curr_left - prev_left)
        
        if 0 < vertical_dist < spacing_limit and horizontal_diff < 50:
            # 結合する！
            prev['text'].append("\n") # 改行を入れる
            prev['text'].extend(curr['text'])
            prev['left'].extend(curr['left'])
            prev['top'].extend(curr['top'])
            prev['width'].extend(curr['width'])
            prev['height'].extend(curr['height'])
        else:
            # 新しいブロックとして開始
            current_merge_id += 1
            merged_blocks[current_merge_id] = curr.copy()
            
    return merged_blocks

# ===========================
# メイン処理フロー
# ===========================

if uploaded_file is not None:
    # PDF基本情報
    pdf_bytes = uploaded_file.read()
    pdf_info = pdfinfo_from_bytes(pdf_bytes)
    max_pages = pdf_info["Pages"]
    
    # --- プレビュー機能 ---
    st.subheader(f"👁️ リアルタイムプレビュー (P.{preview_page})")
    col_prev1, col_prev2 = st.columns([1, 1])
    
    with col_prev1:
        st.caption("処理前")
        # プレビュー用に1枚だけ変換
        if preview_page > max_pages: preview_page = max_pages
        preview_img = convert_from_bytes(pdf_bytes, first_page=preview_page, last_page=preview_page)[0]
        st.image(preview_img, use_container_width=True)

    with col_prev2:
        st.caption("処理イメージ（赤枠＝消える場所 / 青枠＝認識された図）")
        # OpenCVで加工プレビューを表示
        cv_prev = np.array(preview_img)
        cv_prev = cv2.cvtColor(cv_prev, cv2.COLOR_RGB2BGR)
        h, w = cv_prev.shape[:2]
        
        # ロゴ消しエリア（赤塗りつぶし）
        if use_erase:
            overlay = cv_prev.copy()
            cv2.rectangle(overlay, (w - erase_width, h - erase_height), (w, h), (0, 0, 255), -1)
            cv2.addWeighted(overlay, 0.4, cv_prev, 0.6, 0, cv_prev)
            
        # 図形認識エリア（青枠）
        if mode.startswith("分解"):
            objects = extract_objects(cv_prev, min_area_size)
            for (x, y, ow, oh) in objects:
                cv2.rectangle(cv_prev, (x, y), (x+ow, y+oh), (255, 0, 0), 5)
        
        st.image(cv2.cvtColor(cv_prev, cv2.COLOR_BGR2RGB), use_container_width=True)

    # --- 変換実行エリア ---
    st.divider()
    page_range = st.slider("変換ページ範囲", 1, max_pages, (1, min(max_pages, 5)))
    start_p, end_p = page_range
    
    if st.button("🔥 設定を確定して変換スタート", type="primary"):
        process_count = end_p - start_p + 1
        progress_bar = st.progress(0)
        status = st.empty()
        
        # テンプレート読み込み
        if template_file:
            prs = Presentation(template_file)
        else:
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
        
        blank_layout = prs.slide_layouts[len(prs.slide_layouts)-1]

        # 本番変換処理
        images = convert_from_bytes(pdf_bytes, dpi=300, first_page=start_p, last_page=end_p)
        
        for i, image in enumerate(images):
            current_num = start_p + i
            status.text(f"処理中: {i+1}/{process_count} (P.{current_num}) - 画像解析...")
            
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
            h_orig, w_orig = cv_img.shape[:2]
            scale_x = prs.slide_width / w_orig
            scale_y = prs.slide_height / h_orig
            
            slide = prs.slides.add_slide(blank_layout)
            object_rects = []

            # 1. ロゴ消し適用
            if use_erase:
                mask = np.zeros((h_orig, w_orig), np.uint8)
                cv2.rectangle(mask, (w_orig - erase_width, h_orig - erase_height), (w_orig, h_orig), 255, -1)
                cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

            # 2. 画像/図表配置
            if mode.startswith("分解"):
                objects = extract_objects(cv_img, min_area_size)
                for (ox, oy, ow, oh) in objects:
                    crop = cv_img[oy:oy+oh, ox:ox+ow]
                    stream = io.BytesIO()
                    cv2.imencode(".jpg", crop, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_quality])[1].tofile(stream)
                    slide.shapes.add_picture(stream, int(ox*scale_x), int(oy*scale_y), width=int(ow*scale_x), height=int(oh*scale_y))
                    object_rects.append((ox, oy, ow, oh))
            else:
                stream = io.BytesIO()
                cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_quality])[1].tofile(stream)
                slide.shapes.add_picture(stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

            # 3. テキスト処理
            status.text(f"処理中: {i+1}/{process_count} (P.{current_num}) - OCRと段落結合...")
            ocr_img, _, _ = preprocess_image_for_ocr(cv_img, 2.0)
            d = pytesseract.image_to_data(ocr_img, lang='jpn', output_type=Output.DICT)
            
            raw_blocks = {}
            for j in range(len(d['text'])):
                text = d['text'][j].strip()
                if int(d['conf'][j]) > 40 and text != "" and not (len(text)==1 and text in ".,-_|"):
                    b_id = d['block_num'][j]
                    if b_id not in raw_blocks: raw_blocks[b_id] = {'text': [], 'left': [], 'top': [], 'width': [], 'height': []}
                    raw_blocks[b_id]['text'].append(text)
                    raw_blocks[b_id]['left'].append(d['left'][j])
                    raw_blocks[b_id]['top'].append(d['top'][j])
                    raw_blocks[b_id]['width'].append(d['width'][j])
                    raw_blocks[b_id]['height'].append(d['height'][j])
            
            # ★段落結合AIの発動★
            if merge_lines:
                # 結合前にOCR座標(拡大版)から元画像座標へ戻す必要があるが、
                # merge_text_blocksは相対距離で判定するのでOCR座標のままで結合してから縮尺計算する
                final_blocks = merge_text_blocks(raw_blocks, spacing_limit=int(line_spacing_limit * 2.0)) # 2.0はOCR拡大分
            else:
                final_blocks = raw_blocks

            # テキスト配置
            for b_id, b_data in final_blocks.items():
                text_content = "".join(b_data['text']) if not merge_lines else "".join([t if t=="\n" else t for t in b_data['text']])
                text_content = text_content.replace("\n", "\n") # 改行コード正規化

                # 座標計算（全要素の包含矩形）
                ocr_x = min(b_data['left'])
                ocr_y = min(b_data['top'])
                ocr_w = max([l+w for l,w in zip(b_data['left'], b_data['width'])]) - ocr_x
                ocr_h = max([t+h for t,h in zip(b_data['top'], b_data['height'])]) - ocr_y
                
                orig_x, orig_y = int(ocr_x/2.0), int(ocr_y/2.0)
                orig_w, orig_h = int(ocr_w/2.0), int(ocr_h/2.0)

                # 重なり判定
                if mode.startswith("分解"):
                    cx, cy = orig_x + orig_w/2, orig_y + orig_h/2
                    if any(ox < cx < ox+ow and oy < cy < oy+oh for (ox,oy,ow,oh) in object_rects): continue

                # 配置
                pp_x, pp_y = int(orig_x * scale_x), int(orig_y * scale_y)
                pp_w = int(orig_w * scale_x)
                
                if pp_w > Inches(0.2):
                    try:
                        # フォントサイズ推定
                        avg_h = (sum(b_data['height'])/len(b_data['height']))/2.0
                        font_pt = max(9, min((prs.slide_height.inches*72)*(avg_h/h_orig)*0.8, 80))
                        
                        # 高さ自動調整（段落結合時は高さを長めに）
                        box_h = Inches(font_pt/72 * 1.5 * (text_content.count('\n') + 1))
                        
                        txBox = slide.shapes.add_textbox(pp_x, pp_y, pp_w, box_h)
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        p = tf.paragraphs[0]
                        p.text = text_content
                        p.font.size = Pt(font_pt)
                        
                        # 色検出
                        if True: # 常に色検出を試みる
                            y1,y2 = max(0,orig_y), min(h_orig,orig_y+orig_h)
                            x1,x2 = max(0,orig_x), min(w_orig,orig_x+orig_w)
                            roi = cv_img[y1:y2, x1:x2]
                            if roi.size > 0: 
                                r,g,b = get_dominant_color(roi)
                                p.font.color.rgb = RGBColor(r,g,b)
                    except: pass
            
            progress_bar.progress((i + 1) / process_count)

        # 完了
        status.empty()
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        st.success("🎉 変換完了！")
        st.download_button(f"📥 ダウンロード", out, f"{uploaded_file.name}_GodMode.pptx")
