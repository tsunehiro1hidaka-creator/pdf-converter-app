import streamlit as st
import sys
import os
import io

# === 1. ページ設定 ===
st.set_page_config(page_title="Biz PDF Converter Ultimate", layout="centered")

# UIスタイル
st.markdown("""
<style>
    .stButton>button { border-radius: 5px; font-weight: bold; width: 100%; }
    .stProgress .st-bo { background-color: #4CAF50; }
    img { border: 1px solid #ddd; border-radius: 5px; }
    /* ガイド部分を見やすく */
    .streamlit-expanderHeader { font-weight: bold; font-size: 1.1em; color: #2E7D32; }
</style>
""", unsafe_allow_html=True)

# === 2. ライブラリ読み込み ===
try:
    import fitz  # PyMuPDF
    import cv2
    import numpy as np
    import pytesseract
    from PIL import Image
    from pptx import Presentation
    from pptx.util import Inches, Emu, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.dml.fill import MSO_FILL
except ImportError as e:
    st.error(f"ライブラリ不足: {e}")
    st.stop()

st.title("🏆 Biz PDF Converter Ultimate")
st.caption("PDFをパワーポイントに変換する、ビジネス専用ツールです。")

# ===========================
# ★ここが追加ポイント：親切な使い方ガイド
# ===========================
with st.expander("🔰 初めての方へ（使い方の手順）", expanded=False):
    st.markdown("""
    **ようこそ！ このツールは、PDF資料を「編集できるパワーポイント」に変換します。**
    以下の手順で操作してください。
    
    ---
    
    ### 1️⃣ ファイルをアップロード
    下の「PDFファイルをアップロード」と書かれた場所に、PDFファイルを置いてください。（複数を一度に選んでもOKです）
    
    ### 2️⃣ 設定を調整（左のメニュー）
    画面左側のメニューで、細かい調整ができます。
    * **スライドサイズ:** 基本は「PDFに合わせる」のままでOKです。
    * **画質補正:** 文字が薄いときは「コントラスト」を右にずらすとクッキリします。
    * **ロゴ消し:** 資料の下にある不要なページ番号などを消せます。「プレビュー」を見ながら赤枠を調整してください。
    
    ### 3️⃣ プレビューで確認
    画面の中央に「仕上がりプレビュー」が表示されます。
    赤枠の部分が消える範囲です。問題なければ次に進みます。
    
    ### 4️⃣ 変換スタート
    「変換スタート」ボタンを押してください。処理が終わると、ダウンロードボタンが出てきます。
    
    ---
    **💡 便利な機能：文字の修正について**
    変換後のパワーポイントには、右上に**「修正用パッチ（白い箱）」**を用意しています。
    文字化けしている箇所や直したい文字の上に、この白い箱をマウスで移動させて、上から正しい文字を入力してください。
    """)

# ===========================
# 関数定義
# ===========================

def load_pdf_doc(file_bytes):
    try:
        return fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        return None

def preprocess_image_for_ocr(cv_img):
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary

def adjust_image(cv_img, brightness=0, contrast=0):
    if brightness == 0 and contrast == 0:
        return cv_img
    alpha = (contrast + 100.0) / 100.0 
    beta = brightness
    adjusted = cv2.convertScaleAbs(cv_img, alpha=alpha, beta=beta)
    return adjusted

def add_watermark(cv_img, text="CONFIDENTIAL"):
    h, w = cv_img.shape[:2]
    overlay = cv_img.copy()
    font = cv2.FONT_HERSHEY_SIMPLEX
    scale = w / 1000.0 * 2.0
    thickness = int(scale * 2)
    (text_w, text_h), _ = cv2.getTextSize(text, font, scale, thickness)
    x = int((w - text_w) /
