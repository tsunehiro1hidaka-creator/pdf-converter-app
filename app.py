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
    jpeg_q = st.slider("画質品質", 10, 10
