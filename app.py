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
        
        if 0 < vertical_dist < spacing_
