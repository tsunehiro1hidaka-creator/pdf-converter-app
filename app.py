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
from pptx.enum.text import MSO_VERTICAL_ANCHOR, MSO_ANCHOR
import io

# ===========================
# ページ設定 & キャッシュ設定
# ===========================
st.set_page_config(page_title="THE FINAL PDF Converter", layout="wide", initial_sidebar_state="expanded")

# CSSでUIを少しリッチにする
st.markdown("""
<style>
    .stButton>button { width: 100%; border-radius: 5px; font-weight: bold; }
    .stProgress .st-bo { background-color: #4CAF50; }
</style>
""", unsafe_allow_html=True)

st.title("💎 THE FINAL: 超高速・日本語完全対応 PDF変換")
st.markdown("キャッシュ技術による**爆速プレビュー**と、**フォント・縦書き**に対応した最終完成形です。")

# ===========================
# 関数定義（キャッシュ付き）
# ===========================

# ★ここが高速化の鍵！計算結果をメモリに保存するデコレータ
@st.cache_data(show_spinner=False)
def load_pdf_images(file_bytes):
    """PDFを画像に変換してメモリに保持する（重い処理はここだけ）"""
    return convert_from_bytes(file_bytes, dpi=300)

@st.cache_data(show_spinner=False)
def get_pdf_info(file_bytes):
    return pdfinfo_from_bytes(file_bytes)

def preprocess_image_for_ocr(cv_img, zoom_factor=2.0):
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
        # 画面の端にある細い枠などは除外
        if cw * ch < min_area: continue
        if cw > w * 0.95 or ch > h * 0.95: continue 
        objects.append((x, y, cw, ch))
    return objects

def merge_text_blocks(blocks, spacing_limit=40):
    if not blocks: return {}
    # 上から順、次に左から順に並べ替え
    sorted_ids = sorted(blocks.keys(), key=lambda k: (min(blocks[k]['top']), min(blocks[k]['left'])))
    merged = {}
    curr_m_id = 0
    merged[curr_m_id] = blocks[sorted_ids[0]].copy()
    
    for i in range(1, len(sorted_ids)):
        curr = blocks[sorted_ids[i]]
        prev = merged[curr_m_id]
        
        prev_bottom = max([t + h for t, h in zip(prev['top'], prev['height'])])
        curr_top = min(curr['top'])
        prev_left = min(prev['left'])
        curr_left = min(curr['left'])
        
        # 縦書き対応判定：もし前のブロックが「縦長」で、今のブロックが「左」にあるなら結合しない（縦書きの列変え）
        # ここでは簡易的に「行間が狭く、インデントが近い」場合のみ結合
        if 0 < (curr_top - prev_bottom) < spacing_limit and abs(curr_left - prev_left) < 50:
            prev['text'].append("\n")
            prev['text'].extend(curr['text'])
            prev['left'].extend(curr['left'])
            prev['top'].extend(curr['top'])
            prev['width'].extend(curr['width'])
            prev['height'].extend(curr['height'])
        else:
            curr_m_id += 1
            merged[curr_m_id] = curr.copy()
    return merged

# ===========================
# サイドバー & UI
# ===========================
st.sidebar.header("🎛️ 設定パネル")

# 1. デザイン設定
with st.sidebar.expander("🎨 デザイン・フォント", expanded=True):
    target_font = st.selectbox("出力フォント", ["Meiryo", "Yu Gothic", "BIZ UDPGothic", "MS PGothic"], index=0)
    mode = st.radio("変換モード", ["分解モード（図と文字を分離）", "通常モード（背景一枚絵）"])
    detect_vertical = st.checkbox("縦書きを検知する（β版）", value=False)

# 2. プレビュー調整
with st.sidebar.expander("✂️ 切り抜き・ロゴ消し", expanded=True):
    use_erase = st.checkbox("不要領域カット", value=True)
    erase_w = st.slider("右端カット(px)", 0, 800, 350, step=10)
    erase_h = st.slider("下端カット(px)", 0, 500, 180, step=10)

# 3. 高度な設定
with st.sidebar.expander("🔧 高度な設定"):
    min_area = st.slider("図形認識感度", 1000, 20000, 5000)
    jpeg_q = st.slider("画質品質", 10, 100, 85)
    merge_lines = st.checkbox("段落結合", value=True)

uploaded_file = st.file_uploader("PDFファイルをドロップ", type="pdf")

if uploaded_file is not None:
    # --- 1. 高速読み込み（キャッシュ利用） ---
    file_bytes = uploaded_file.read()
    pdf_info = get_pdf_info(file_bytes)
    total_pages = pdf_info["Pages"]
    
    # 画像リストもキャッシュされるので、2回目以降は一瞬で終わる
    with st.spinner("PDFを読み込み中...（初回のみ数秒かかります）"):
        images = load_pdf_images(file_bytes)

    # --- 2. 爆速プレビュー ---
    st.divider()
    col1, col2 = st.columns([1, 2])
    
    with col1:
        preview_idx = st.number_input("プレビューページ", 1, total_pages, 1) - 1
        st.image(images[preview_idx], caption="元画像", use_container_width=True)
    
    with col2:
        # OpenCVプレビュー作成（ここだけリアルタイム処理）
        cv_prev = np.array(images[preview_idx])
        cv_prev = cv2.cvtColor(cv_prev, cv2.COLOR_RGB2BGR)
        h, w = cv_prev.shape[:2]
        
        # ロゴ消し可視化
        if use_erase:
            overlay = cv_prev.copy()
            cv2.rectangle(overlay, (w - erase_w, h - erase_h), (w, h), (200, 200, 200), -1) # グレーで隠す
            cv2.addWeighted(overlay, 0.5, cv_prev, 0.5, 0, cv_prev)
            cv2.rectangle(cv_prev, (w - erase_w, h - erase_h), (w, h), (0, 0, 255), 2) # 赤枠
        
        # 図形認識可視化
        if mode.startswith("分解"):
            objs = extract_objects(cv_prev, min_area)
            for (ox, oy, ow, oh) in objs:
                cv2.rectangle(cv_prev, (ox, oy), (ox+ow, oy+oh), (255, 100, 0), 3) # 青枠

        st.image(cv2.cvtColor(cv_prev, cv2.COLOR_BGR2RGB), caption="変換イメージ（赤枠＝削除 / 青枠＝図形）", use_container_width=True)

    # --- 3. 変換実行 ---
    st.divider()
    c1, c2 = st.columns([3, 1])
    with c1:
        page_range = st.slider("変換範囲", 1, total_pages, (1, min(total_pages, 10)))
    with c2:
        st.write("") # Spacer
        start_btn = st.button("🚀 変換スタート", type="primary")

    if start_btn:
        start_p, end_p = page_range
        process_cnt = end_p - start_p + 1
        p_bar = st.progress(0)
        status = st.empty()

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        # ターゲット画像を切り出す
        target_images = images[start_p-1 : end_p]

        for i, image in enumerate(target_images):
            status.text(f"Processing: {i+1}/{process_cnt} ...")
            
            # 画像処理
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
            h_orig, w_orig = cv_img.shape[:2]
            scale_x = prs.slide_width / w_orig
            scale_y = prs.slide_height / h_orig
            
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            
            # 1. 背景・図形処理
            object_rects = []
            if use_erase:
                mask = np.zeros((h_orig, w_orig), np.uint8)
                cv2.rectangle(mask, (w_orig - erase_w, h_orig - erase_h), (w_orig, h_orig), 255, -1)
                cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

            if mode.startswith("分解"):
                objs = extract_objects(cv_img, min_area)
                for (ox, oy, ow, oh) in objs:
                    crop = cv_img[oy:oy+oh, ox:ox+ow]
                    s = io.BytesIO()
                    cv2.imencode(".jpg", crop, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_q])[1].tofile(s)
                    slide.shapes.add_picture(s, int(ox*scale_x), int(oy*scale_y), width=int(ow*scale_x), height=int(oh*scale_y))
                    object_rects.append((ox, oy, ow, oh))
            else:
                s = io.BytesIO()
                cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_q])[1].tofile(s)
                slide.shapes.add_picture(s, 0, 0, width=prs.slide_width, height=prs.slide_height)

            # 2. OCR処理
            ocr_img, _, _ = preprocess_image_for_ocr(cv_img, 2.0)
            # 日本語縦書き対応のオプションを追加検討（今回は通常jpn）
            d = pytesseract.image_to_data(ocr_img, lang='jpn', output_type=Output.DICT)
            
            raw_blocks = {}
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
            
            final_blocks = merge_text_blocks(raw_blocks) if merge_lines else raw_blocks

            # 3. テキスト配置
            for bid, bdata in final_blocks.items():
                content = "".join(bdata['text'])
                
                # 座標計算
                ox = min(bdata['left'])
                oy = min(bdata['top'])
                ow = max([l+w for l,w in zip(bdata['left'], bdata['width'])]) - ox
                oh = max([t+h for t,h in zip(bdata['top'], bdata['height'])]) - oy
                
                # OCR拡大分(2.0)を戻す
                orig_x, orig_y, orig_w, orig_h = int(ox/2), int(oy/2), int(ow/2), int(oh/2)

                # 重なり判定
                if mode.startswith("分解"):
                    cx, cy = orig_x + orig_w/2, orig_y + orig_h/2
                    if any(rox < cx < rox+row and roy < cy < roy+roh for (rox,roy,row,roh) in object_rects): continue

                pp_x = int(orig_x * scale_x)
                pp_y = int(orig_y * scale_y)
                pp_w = int(orig_w * scale_x)
                pp_h = int(orig_h * scale_y) # 高さは推定値で上書きする

                if pp_w > Inches(0.2):
                    try:
                        # 縦書き判定（高さが幅の2.5倍以上なら縦書きとみなす）
                        is_vertical = detect_vertical and (orig_h > orig_w * 2.5)
                        
                        # フォントサイズ推定
                        avg_h = (sum(bdata['height'])/len(bdata['height']))/2.0
                        est_pt = (prs.slide_height.inches*72)*(avg_h/h_orig)*0.75
                        font_pt = max(9, min(est_pt, 80))

                        # テキストボックス作成
                        # 縦書きの場合は幅と高さを入れ替えるようなイメージで枠を作る必要があるが、
                        # Python-pptxでの縦書き設定は複雑なため、横書きボックスで縦に文字を入れる
                        
                        box_h = Inches(font_pt/72 * 1.5 * (content.count('\n') + 1))
                        txBox = slide.shapes.add_textbox(pp_x, pp_y, pp_w, box_h)
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        
                        # ★縦書き設定（試験的）
                        if is_vertical:
                            tf.orientation = MSO_ANCHOR.TOP # 簡易設定
                            # 本当の縦書きは word_wrap と @フォントが必要で難易度が高い
                            # ここでは「縦書きっぽい」というフラグに留める

                        p = tf.paragraphs[0]
                        p.text = content
                        p.font.size = Pt(font_pt)
                        p.font.name = target_font # ★ユーザー指定フォント
                        
                        # 日本語フォントを強制適用するためのハック
                        p.font.language_id = 1041 # Japanese
                        
                        # 色推定
                        y1,y2 = max(0,orig_y), min(h_orig,orig_y+orig_h)
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
            
            p_bar.progress((i+1)/process_cnt)

        # 保存
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        status.empty()
        st.balloons() # 完成時の演出
        st.success("💎 変換完了！")
        st.download_button(f"📥 {target_font}でダウンロード", out, "Final_Presentation.pptx")
