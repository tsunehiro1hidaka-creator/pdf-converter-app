import streamlit as st
import sys
import os
import io

# === 1. ページ設定 ===
st.set_page_config(page_title="Simple PDF Converter", layout="centered")

# UIスタイル
st.markdown("""
<style>
    .stButton>button { border-radius: 5px; font-weight: bold; width: 100%; }
    .stProgress .st-bo { background-color: #4CAF50; }
</style>
""", unsafe_allow_html=True)

# === 2. ライブラリ読み込みチェック ===
try:
    import fitz  # PyMuPDF
    import cv2
    import numpy as np
    import pytesseract
    from PIL import Image
    from pptx import Presentation
    from pptx.util import Inches
except ImportError as e:
    st.error(f"ライブラリ不足: {e}")
    st.stop()

st.title("📄 PDF → パワポ変換 (安定版)")
st.markdown("レイアウト崩れなし。画像として貼り付け、テキストはノートに抽出します。")

# ===========================
# 関数定義
# ===========================

@st.cache_data(show_spinner=False)
def load_pdf_doc(file_bytes):
    """PDFをメモリに読み込む"""
    try:
        return fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        st.error(f"PDF読み込みエラー: {e}")
        return None

def preprocess_image_for_ocr(cv_img):
    """OCR精度向上のための画像加工"""
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    # ノイズ除去と二値化
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary

# ===========================
# サイドバー設定
# ===========================
st.sidebar.header("設定")

use_erase = st.sidebar.checkbox("ロゴ/不要領域を隠す", value=True)
erase_w = st.sidebar.slider("右端カット (px)", 0, 800, 350)
erase_h = st.sidebar.slider("下端カット (px)", 0, 500, 180)

ocr_enabled = st.sidebar.checkbox("テキスト抽出（ノートへ）", value=True)

# ===========================
# メイン処理
# ===========================
uploaded_file = st.file_uploader("PDFファイルをアップロード", type="pdf")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    doc = load_pdf_doc(file_bytes)
    
    if doc:
        total_pages = len(doc)
        st.success(f"読み込み完了: 全 {total_pages} ページ")

        # --- プレビュー機能 ---
        st.divider()
        st.subheader("プレビュー (ロゴ消し確認)")
        
        # 1ページ目を画像化して表示
        page1 = doc[0]
        pix = page1.get_pixmap(matrix=fitz.Matrix(1.0, 1.0)) # プレビューは標準画質
        img_prev = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        cv_prev = cv2.cvtColor(np.array(img_prev), cv2.COLOR_RGB2BGR)
        h, w = cv_prev.shape[:2]

        # ロゴ消しエリアの可視化
        if use_erase:
            # 赤枠で表示
            cv2.rectangle(cv_prev, (w - erase_w, h - erase_h), (w, h), (0, 0, 255), 3)
            # 塗りつぶしイメージ
            overlay = cv_prev.copy()
            cv2.rectangle(overlay, (w - erase_w, h - erase_h), (w, h), (255, 255, 255), -1)
            cv2.addWeighted(overlay, 0.7, cv_prev, 0.3, 0, cv_prev)

        st.image(cv2.cvtColor(cv_prev, cv2.COLOR_BGR2RGB), caption="赤枠の部分が白塗りされます", use_container_width=True)

        # --- 変換実行 ---
        st.divider()
        if st.button("変換スタート", type="primary"):
            p_bar = st.progress(0)
            status_area = st.empty()
            
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
            
            for i, page in enumerate(doc):
                status_area.text(f"処理中... {i+1}/{total_pages} ページ")
                
                # 1. 高画質で画像化 (zoom=2.0)
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                cv_img = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
                h_orig, w_orig = cv_img.shape[:2]
                
                # 2. ロゴ消し処理
                if use_erase:
                    # 単純な白塗りで隠す（一番高速で確実）
                    cv2.rectangle(cv_img, (w_orig - erase_w, h_orig - erase_h), (w_orig, h_orig), (255, 255, 255), -1)
                
                # 3. パワポに貼り付け
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                # 画像をバイト列に変換して貼り付け
                img_bytes = cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), 90])[1].tobytes()
                image_stream = io.BytesIO(img_bytes)
                slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
                
                # 4. OCR (ノートへ書き込み)
                if ocr_enabled:
                    try:
                        # OCR用の画像準備（少し縮小しても十分読めるので速度優先）
                        ocr_img = preprocess_image_for_ocr(cv_img)
                        text = pytesseract.image_to_string(ocr_img, lang='jpn+eng')
                        slide.notes_slide.notes_text_frame.text = text
                    except:
                        slide.notes_slide.notes_text_frame.text = "(文字読み取りに失敗しました)"
                
                p_bar.progress((i + 1) / total_pages)
            
            status_area.success("完了しました！")
            
            # 保存とダウンロード
            out_ppt = io.BytesIO()
            prs.save(out_ppt)
            out_ppt.seek(0)
            
            st.download_button(
                label="📥 パワーポイントをダウンロード",
                data=out_ppt,
                file_name=f"{uploaded_file.name}_slide.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
