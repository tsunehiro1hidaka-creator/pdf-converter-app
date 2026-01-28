import streamlit as st
import os
import cv2
import numpy as np
import pytesseract
from pytesseract import Output
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

# === ページ設定 ===
st.set_page_config(page_title="PDF完全再現ツール", layout="wide")

st.title("🎨 色も配置も再現！究極のPDF変換ツール")
st.markdown("""
**特徴:**
1. **スライダー**でロゴ消し範囲を自由に調整可能。
2. **文字の色**を自動でスポイトして再現。
3. 画像拡大とノイズ除去で**OCR精度**を限界まで向上。
""")

# --- サイドバー：設定エリア ---
st.sidebar.header("🔧 調整パネル")

# 1. ロゴ消し設定
use_erase = st.sidebar.checkbox("ロゴ/ページ番号を消す", value=True)
erase_width = st.sidebar.slider("消す幅 (横)", 0, 800, 350, step=10, help="右端から内側へ何ピクセル消すか")
erase_height = st.sidebar.slider("消す高さ (縦)", 0, 500, 180, step=10, help="下端から上側へ何ピクセル消すか")

# 2. テキスト設定
bg_fill = st.sidebar.checkbox("文字の背景を白くする", value=True, help="元の文字を隠すために白塗りします")
detect_color = st.sidebar.checkbox("文字色を自動検出する", value=True, help="元の文字色（赤や青など）を再現します")
font_scale = st.sidebar.slider("フォントサイズ微調整", 0.5, 1.5, 0.75, step=0.05)

# 3. PDFアップロード
uploaded_file = st.file_uploader("ここにPDFをドラッグ＆ドロップ", type="pdf")

# --- 関数エリア ---

def get_dominant_color(img_crop):
    """画像の一部から主要な色（文字色）を抽出する"""
    # 背景（白）を除外したいので、少し暗い画素の平均を取る簡易ロジック
    # グレーっぽくならないよう、彩度があるものを優先
    try:
        # 画素をフラットにする
        pixels = np.float32(img_crop.reshape(-1, 3))
        
        # 白に近い画素(200以上)を除外して、文字部分だけの色を狙う
        mask = np.all(pixels < 240, axis=1)
        target_pixels = pixels[mask]
        
        if len(target_pixels) > 0:
            # 平均色を計算
            avg_color = np.mean(target_pixels, axis=0)
            return int(avg_color[0]), int(avg_color[1]), int(avg_color[2])
        else:
            return 0, 0, 0 # 黒
    except:
        return 0, 0, 0

def preprocess_image_for_ocr(cv_img, zoom_factor=2.0):
    """OCR用画像処理"""
    h, w = cv_img.shape[:2]
    new_w, new_h = int(w * zoom_factor), int(h * zoom_factor)
    resized = cv2.resize(cv_img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary, new_w, new_h

# --- メイン処理 ---
if uploaded_file is not None:
    st.info(f"📄 {uploaded_file.name} を読み込みました。左のパネルで設定を調整して「変換」を押してください。")
    
    if st.button("変換スタート"):
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        images = convert_from_bytes(uploaded_file.read(), dpi=300)
        total_pages = len(images)
        
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        progress_text.text("解析を開始します...")

        for i, image in enumerate(images):
            # 画像準備
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR) # OpenCVはBGR
            h_orig, w_orig = cv_img.shape[:2]

            # 1. OCR用画像作成（2倍拡大）
            OCR_ZOOM = 2.0
            ocr_img, w_ocr, h_ocr = preprocess_image_for_ocr(cv_img, OCR_ZOOM)

            # 2. 解析実行
            d = pytesseract.image_to_data(ocr_img, lang='jpn', output_type=Output.DICT)
            
            # 3. ブロック生成
            blocks = {}
            n_boxes = len(d['text'])
            
            for j in range(n_boxes):
                # 【改良】ゴミ除去：信頼度が低い、または空白、または1文字の記号などはスキップ
                text = d['text'][j].strip()
                conf = int(d['conf'][j])
                
                # 信頼度50以下、かつ文字数が少ないゴミは無視
                if conf > 40 and text != "":
                     # 記号だけのゴミを除去する簡易フィルター
                    if len(text) == 1 and text in ".,-~_`'":
                        continue

                    b_id = d['block_num'][j]
                    if b_id not in blocks:
                        blocks[b_id] = {'text': [], 'left': [], 'top': [], 'width': [], 'height': []}
                    
                    blocks[b_id]['text'].append(text)
                    blocks[b_id]['left'].append(d['left'][j])
                    blocks[b_id]['top'].append(d['top'][j])
                    blocks[b_id]['width'].append(d['width'][j])
                    blocks[b_id]['height'].append(d['height'][j])

            # 4. ロゴ消し処理（サイドバーの設定を使用）
            if use_erase:
                mask = np.zeros((h_orig, w_orig), np.uint8)
                cv2.rectangle(mask, (w_orig - erase_width, h_orig - erase_height), (w_orig, h_orig), 255, -1)
                cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

            # 背景画像セット
            image_stream = io.BytesIO()
            is_success, buffer = cv2.imencode(".jpg", cv_img)
            image_stream.write(buffer)
            
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(image_stream, 0
