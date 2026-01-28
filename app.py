import streamlit as st
import sys
import os
import io

# === 1. ページ設定 ===
st.set_page_config(page_title="Smart PDF Converter", layout="centered")

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
    from pptx.util import Inches, Pt, Emu
except ImportError as e:
    st.error(f"ライブラリ不足: {e}")
    st.stop()

st.title("📄 Smart PDF Converter")
st.markdown("サイズ自動調整・軽量化に対応した、実務用・決定版ツールです。")

# ===========================
# 関数定義
# ===========================

@st.cache_data(show_spinner=False)
def load_pdf_doc(file_bytes):
    try:
        return fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        st.error(f"PDF読み込みエラー: {e}")
        return None

def preprocess_image_for_ocr(cv_img):
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary

# ===========================
# サイドバー設定
# ===========================
st.sidebar.header("🎨 デザイン・画質")

# 1. サイズ設定
slide_sizing = st.sidebar.radio(
    "スライドサイズ",
    ["PDFに合わせる (推奨)", "16:9 (ワイド)", "4:3 (標準)"],
    help="元のPDFがA4縦などの場合、「PDFに合わせる」を選ぶとレイアウトが崩れません。"
)

# 2. 軽量化設定
quality_mode = st.sidebar.select_slider(
    "画質とファイルサイズ",
    options=["軽量 (メール用)", "標準", "高画質 (プレゼン用)"],
    value="標準"
)

# 画質パラメータ変換
if quality_mode == "軽量 (メール用)":
    zoom_factor = 1.0   # 72dpi相当
    jpeg_quality = 70
elif quality_mode == "標準":
    zoom_factor = 1.5   # 108dpi相当
    jpeg_quality = 80
else:
    zoom_factor = 2.0   # 144dpi相当
    jpeg_quality = 95

st.sidebar.divider()
st.sidebar.header("✂️ 加工設定")
use_erase = st.sidebar.checkbox("ロゴ/不要領域を白塗り", value=True)
erase_w = st.sidebar.slider("右端カット (px)", 0, 800, 350)
erase_h = st.sidebar.slider("下端カット (px)", 0, 500, 180)
ocr_enabled = st.sidebar.checkbox("テキスト抽出 (ノートへ)", value=True)

# ===========================
# メイン処理
# ===========================
uploaded_file = st.file_uploader("PDFファイルをアップロード", type="pdf")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    doc = load_pdf_doc(file_bytes)
    
    if doc:
        total_pages = len(doc)
        
        # --- PDF情報表示 ---
        page1 = doc[0]
        pdf_w, pdf_h = page1.rect.width, page1.rect.height
        aspect_ratio = pdf_w / pdf_h
        
        st.info(f"読み込み完了: 全 {total_pages} ページ / サイズ: {pdf_w:.0f}x{pdf_h:.0f} pt (比率 {aspect_ratio:.2f})")

        # --- プレビュー ---
        st.subheader("プレビュー (白塗り範囲の確認)")
        pix = page1.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
        img_prev = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        cv_prev = cv2.cvtColor(np.array(img_prev), cv2.COLOR_RGB2BGR)
        h, w = cv_prev.shape[:2]

        if use_erase:
            # 赤枠と半透明白塗り
            cv2.rectangle(cv_prev, (w - erase_w, h - erase_h), (w, h), (0, 0, 255), 3)
            overlay = cv_prev.copy()
            cv2.rectangle(overlay, (w - erase_w, h - erase_h), (w, h), (255, 255, 255), -1)
            cv2.addWeighted(overlay, 0.7, cv_prev, 0.3, 0, cv_prev)

        st.image(cv2.cvtColor(cv_prev, cv2.COLOR_BGR2RGB), caption="プレビュー画面", use_container_width=True)

        # --- 変換実行 ---
        st.divider()
        if st.button("変換スタート", type="primary"):
            p_bar = st.progress(0)
            status_area = st.empty()
            
            prs = Presentation()
            
            # ★ここが進化ポイント：スライドサイズの決定
            if slide_sizing == "PDFに合わせる (推奨)":
                # PDFのポイント単位(pt)をそのままパワポのEmu単位に変換
                # 1 pt = 12700 Emu
                prs.slide_width = Emu(pdf_w * 12700)
                prs.slide_height = Emu(pdf_h * 12700)
            elif slide_sizing == "16:9 (ワイド)":
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
            else: # 4:3
                prs.slide_width = Inches(10)
                prs.slide_height = Inches(7.5)

            for i, page in enumerate(doc):
                status_area.text(f"処理中... {i+1}/{total_pages} ページ")
                
                # 1. 画像化 (設定した画質で)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom_factor, zoom_factor))
                img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                cv_img = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
                h_orig, w_orig = cv_img.shape[:2]
                
                # 2. 白塗り
                if use_erase:
                    # PyMuPDFのzoomに合わせて消す範囲も調整する
                    # 画面のスライダーは「元サイズ基準」だと思われるので、倍率をかける
                    # ただしプレビューとズレると困るので、今回はスライダー値を
                    # 「表示されている画像に対するピクセル数」として簡易的に扱うため
                    # 画像右下からの固定ピクセル除去とする
                    cv2.rectangle(cv_img, (w_orig - int(erase_w * zoom_factor), h_orig - int(erase_h * zoom_factor)), 
                                  (w_orig, h_orig), (255, 255, 255), -1)
                
                # 3. 配置
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                img_bytes = cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_quality])[1].tobytes()
                image_stream = io.BytesIO(img_bytes)
                
                # 画像をスライドいっぱいに貼る
                slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
                
                # 4. OCR (ノート)
                if ocr_enabled:
                    try:
                        # OCRは重いので、軽量設定のときはスキップするか、解像度を落とす手もあるが
                        # ここではそのまま実行（エラー時はスルー）
                        ocr_img = preprocess_image_for_ocr(cv_img)
                        text = pytesseract.image_to_string(ocr_img, lang='jpn+eng')
                        slide.notes_slide.notes_text_frame.text = text
                    except:
                        slide.notes_slide.notes_text_frame.text = ""
                
                p_bar.progress((i + 1) / total_pages)
            
            status_area.success("完了しました！")
            
            out_ppt = io.BytesIO()
            prs.save(out_ppt)
            out_ppt.
