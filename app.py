import streamlit as st
import sys
import os
import io

# === 1. ページ設定 ===
st.set_page_config(page_title="Biz PDF Converter (Preview)", layout="centered")

# UIスタイル
st.markdown("""
<style>
    .stButton>button { border-radius: 5px; font-weight: bold; width: 100%; }
    .stProgress .st-bo { background-color: #4CAF50; }
    /* プレビュー画像を見やすく */
    img { border: 1px solid #ddd; border-radius: 5px; }
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
    from pptx.util import Inches, Emu
except ImportError as e:
    st.error(f"ライブラリ不足: {e}")
    st.stop()

st.title("💼 Biz PDF Converter")
st.markdown("プレビューで確認しながら、確実にスライド化できる実務ツールです。")

# ===========================
# 関数定義
# ===========================

@st.cache_data(show_spinner=False)
def load_pdf_doc(file_bytes):
    try:
        return fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        return None

def preprocess_image_for_ocr(cv_img):
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary

def add_watermark(cv_img, text="CONFIDENTIAL"):
    h, w = cv_img.shape[:2]
    overlay = cv_img.copy()
    font = cv2.FONT_HERSHEY_SIMPLEX
    scale = w / 1000.0 * 2.0
    thickness = int(scale * 2)
    (text_w, text_h), _ = cv2.getTextSize(text, font, scale, thickness)
    x = int((w - text_w) / 2)
    y = int((h + text_h) / 2)
    cv2.putText(overlay, text, (x, y), font, scale, (150, 150, 150), thickness, cv2.LINE_AA)
    cv2.addWeighted(overlay, 0.2, cv_img, 0.8, 0, cv_img)
    return cv_img

# ===========================
# サイドバー設定
# ===========================
st.sidebar.header("🎨 出力設定")

slide_sizing = st.sidebar.radio("スライドサイズ", ["PDFに合わせる (推奨)", "16:9 (ワイド)", "4:3 (標準)"])
quality_mode = st.sidebar.select_slider("画質設定", options=["軽量", "標準", "高画質"], value="標準")

if quality_mode == "軽量": zoom_factor=1.0; jpeg_quality=70
elif quality_mode == "標準": zoom_factor=1.5; jpeg_quality=80
else: zoom_factor=2.0; jpeg_quality=95

st.sidebar.divider()
st.sidebar.header("🛡️ 加工オプション")
watermark_text = st.sidebar.text_input("透かし文字", value="")
ocr_enabled = st.sidebar.checkbox("テキスト抽出 (ノートへ)", value=True)

st.sidebar.subheader("✂️ ロゴ消し調整")
use_erase = st.sidebar.checkbox("ロゴ/不要領域の削除", value=True)
erase_w = st.sidebar.slider("右端カット (px)", 0, 800, 350)
erase_h = st.sidebar.slider("下端カット (px)", 0, 500, 180)

# ===========================
# メイン処理
# ===========================
uploaded_files = st.file_uploader("PDFファイルをアップロード (複数可)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    # ファイル読み込み
    docs = []
    total_pages_all = 0
    for up_file in uploaded_files:
        file_bytes = up_file.read()
        doc = load_pdf_doc(file_bytes)
        if doc:
            docs.append((up_file.name, doc))
            total_pages_all += len(doc)
    
    if docs:
        first_doc = docs[0][1]
        
        # --- ★強化されたプレビュー機能 ---
        st.divider()
        st.subheader("👁️ プレビュー確認")
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            # ページ選択機能
            preview_page_idx = st.number_input("確認するページ番号", min_value=1, max_value=len(first_doc), value=1) - 1
            preview_mode = st.radio("表示モード", ["赤枠で範囲を確認", "仕上がりを確認 (白塗り)"])
        
        with col2:
            # プレビュー画像生成 (高速化のため標準画質)
            page = first_doc[preview_page_idx]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
            img_prev = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            cv_prev = cv2.cvtColor(np.array(img_prev), cv2.COLOR_RGB2BGR)
            h, w = cv_prev.shape[:2]

            # 透かしプレビュー
            if watermark_text:
                cv_prev = add_watermark(cv_prev, watermark_text)

            # ロゴ消しプレビュー
            if use_erase:
                if preview_mode == "赤枠で範囲を確認":
                    # 赤枠と半透明赤塗り
                    cv2.rectangle(cv_prev, (w - erase_w, h - erase_h), (w, h), (0, 0, 255), 4)
                    overlay = cv_prev.copy()
                    cv2.rectangle(overlay, (w - erase_w, h - erase_h), (w, h), (255, 200, 200), -1)
                    cv2.addWeighted(overlay, 0.5, cv_prev, 0.5, 0, cv_prev)
                else:
                    # 本番と同じ白塗り
                    cv2.rectangle(cv_prev, (w - erase_w, h - erase_h), (w, h), (255, 255, 255), -1)

            st.image(cv2.cvtColor(cv_prev, cv2.COLOR_BGR2RGB), caption=f"{docs[0][0]} - P.{preview_page_idx + 1}", use_container_width=True)

        # --- 変換実行 ---
        st.divider()
        if st.button("変換スタート", type="primary"):
            p_bar = st.progress(0)
            status_area = st.empty()
            
            prs = Presentation()
            
            # 1枚目のPDFサイズを基準にスライドサイズ決定
            page1 = first_doc[0]
            pdf_w, pdf_h = page1.rect.width, page1.rect.height
            
            if slide_sizing == "PDFに合わせる (推奨)":
                prs.slide_width = Emu(pdf_w * 12700)
                prs.slide_height = Emu(pdf_h * 12700)
            elif slide_sizing == "16:9 (ワイド)":
                prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
            else:
                prs.slide_width = Inches(10); prs.slide_height = Inches(7.5)

            current_cnt = 0
            
            for filename, doc in docs:
                status_area.text(f"処理中: {filename} ...")
                
                for i, page in enumerate(doc):
                    # 画像化
                    pix = page.get_pixmap(matrix=fitz.Matrix(zoom_factor, zoom_factor))
                    img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    cv_img = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
                    h_orig, w_orig = cv_img.shape[:2]
                    
                    if watermark_text: cv_img = add_watermark(cv_img, watermark_text)
                    
                    if use_erase:
                        # プレビューと同じ比率で削除範囲を計算
                        cv2.rectangle(cv_img, (w_orig - int(erase_w * zoom_factor), h_orig - int(erase_h * zoom_factor)), 
                                      (w_orig, h_orig), (255, 255, 255), -1)
                    
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    img_bytes = cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_quality])[1].tobytes()
                    image_stream = io.BytesIO(img_bytes)
                    slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
                    
                    if ocr_enabled:
                        try:
                            ocr_img = preprocess_image_for_ocr(cv_img)
                            header = f"[{filename} - P.{i+1}]\n"
                            text = pytesseract.image_to_string(ocr_img, lang='jpn+eng')
                            slide.notes_slide.notes_text_frame.text = header + text
                        except:
                            slide.notes_slide.notes_text_frame.text = ""
                    
                    current_cnt += 1
                    p_bar.progress(current_cnt / total_pages_all)
            
            status_area.success("完了しました！")
            out_ppt = io.BytesIO()
            prs.save(out_ppt)
            out_ppt.seek(0)
            
            dl_name = "Combined_Slides.pptx" if len(docs) > 1 else f"{os.path.splitext(docs[0][0])[0]}_slide.pptx"
            st.download_button("📥 パワポをダウンロード", out_ppt, dl_name)
