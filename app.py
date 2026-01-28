import streamlit as st
import sys
import os
import io

# === 1. ページ設定 ===
st.set_page_config(page_title="Biz PDF Converter", layout="centered")

# UIスタイル
st.markdown("""
<style>
    .stButton>button { border-radius: 5px; font-weight: bold; width: 100%; }
    .stProgress .st-bo { background-color: #4CAF50; }
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
st.markdown("結合・透かし・サイズ調整に対応した、ビジネス実務専用ツールです。")

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
    """画像に半透明の透かし文字を入れる"""
    h, w = cv_img.shape[:2]
    overlay = cv_img.copy()
    
    # 文字の設定
    font = cv2.FONT_HERSHEY_SIMPLEX
    scale = w / 1000.0 * 2.0  # 画像サイズに合わせて大きく
    thickness = int(scale * 2)
    
    # 文字サイズ計測
    (text_w, text_h), _ = cv2.getTextSize(text, font, scale, thickness)
    
    # 中心に配置
    x = int((w - text_w) / 2)
    y = int((h + text_h) / 2)
    
    # 文字を描画（グレー）
    cv2.putText(overlay, text, (x, y), font, scale, (150, 150, 150), thickness, cv2.LINE_AA)
    
    # 半透明合成 (alpha=0.2 くらいが薄くて丁度いい)
    cv2.addWeighted(overlay, 0.2, cv_img, 0.8, 0, cv_img)
    return cv_img

# ===========================
# サイドバー設定
# ===========================
st.sidebar.header("🎨 出力設定")

slide_sizing = st.sidebar.radio(
    "スライドサイズ",
    ["PDFに合わせる (推奨)", "16:9 (ワイド)", "4:3 (標準)"]
)

quality_mode = st.sidebar.select_slider(
    "画質設定",
    options=["軽量", "標準", "高画質"],
    value="標準"
)

# パラメータ設定
if quality_mode == "軽量":
    zoom_factor = 1.0; jpeg_quality = 70
elif quality_mode == "標準":
    zoom_factor = 1.5; jpeg_quality = 80
else:
    zoom_factor = 2.0; jpeg_quality = 95

st.sidebar.divider()
st.sidebar.header("🛡️ 加工オプション")

# 透かし設定
watermark_text = st.sidebar.text_input("透かし文字 (空欄でオフ)", value="")
use_erase = st.sidebar.checkbox("ロゴ/不要領域の白塗り", value=True)
erase_w = st.sidebar.slider("右端カット (px)", 0, 800, 350)
erase_h = st.sidebar.slider("下端カット (px)", 0, 500, 180)

st.sidebar.divider()
ocr_enabled = st.sidebar.checkbox("テキスト抽出 (ノートへ)", value=True)

# ===========================
# メイン処理
# ===========================
# ★新機能: 複数ファイルのアップロードを許可
uploaded_files = st.file_uploader("PDFファイルをアップロード (複数可)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    # ファイル情報を整理
    docs = []
    total_pages_all = 0
    
    for up_file in uploaded_files:
        file_bytes = up_file.read()
        doc = load_pdf_doc(file_bytes)
        if doc:
            docs.append((up_file.name, doc))
            total_pages_all += len(doc)
    
    if docs:
        # 1枚目のPDF情報でサイズ表示
        first_doc = docs[0][1]
        page1 = first_doc[0]
        pdf_w, pdf_h = page1.rect.width, page1.rect.height
        
        st.info(f"読み込み完了: 計 {len(docs)} ファイル / 全 {total_pages_all} ページ")

        # --- プレビュー (1ファイル目の1ページ目) ---
        st.subheader("プレビュー")
        pix = page1.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
        img_prev = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        cv_prev = cv2.cvtColor(np.array(img_prev), cv2.COLOR_RGB2BGR)
        h, w = cv_prev.shape[:2]

        # 透かしプレビュー
        if watermark_text:
            cv_prev = add_watermark(cv_prev, watermark_text)

        # 白塗りプレビュー
        if use_erase:
            cv2.rectangle(cv_prev, (w - erase_w, h - erase_h), (w, h), (0, 0, 255), 3)
            overlay = cv_prev.copy()
            cv2.rectangle(overlay, (w - erase_w, h - erase_h), (w, h), (255, 255, 255), -1)
            cv2.addWeighted(overlay, 0.7, cv_prev, 0.3, 0, cv_prev)

        st.image(cv2.cvtColor(cv_prev, cv2.COLOR_BGR2RGB), caption="仕上がりイメージ", use_container_width=True)

        # --- 変換実行 ---
        st.divider()
        if st.button("まとめて変換スタート", type="primary"):
            p_bar = st.progress(0)
            status_area = st.empty()
            
            prs = Presentation()
            
            # サイズ設定
            if slide_sizing == "PDFに合わせる (推奨)":
                prs.slide_width = Emu(pdf_w * 12700)
                prs.slide_height = Emu(pdf_h * 12700)
            elif slide_sizing == "16:9 (ワイド)":
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
            else:
                prs.slide_width = Inches(10)
                prs.slide_height = Inches(7.5)

            current_page_count = 0
            
            # ★複数ファイルをループ処理
            for filename, doc in docs:
                status_area.text(f"処理中: {filename} ...")
                
                for i, page in enumerate(doc):
                    # 1. 画像化
                    pix = page.get_pixmap(matrix=fitz.Matrix(zoom_factor, zoom_factor))
                    img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    cv_img = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
                    h_orig, w_orig = cv_img.shape[:2]
                    
                    # 2. 透かし (文字がある場合のみ)
                    if watermark_text:
                        cv_img = add_watermark(cv_img, watermark_text)
                    
                    # 3. 白塗り
                    if use_erase:
                        cv2.rectangle(cv_img, (w_orig - int(erase_w * zoom_factor), h_orig - int(erase_h * zoom_factor)), 
                                      (w_orig, h_orig), (255, 255, 255), -1)
                    
                    # 4. 配置
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    img_bytes = cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_quality])[1].tobytes()
                    image_stream = io.BytesIO(img_bytes)
                    slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
                    
                    # 5. OCR
                    if ocr_enabled:
                        try:
                            # ファイル名もノートに入れておくと便利
                            header = f"[{filename} - P.{i+1}]\n"
                            ocr_img = preprocess_image_for_ocr(cv_img)
                            text = pytesseract.image_to_string(ocr_img, lang='jpn+eng')
                            slide.notes_slide.notes_text_frame.text = header + text
                        except:
                            slide.notes_slide.notes_text_frame.text = ""
                    
                    current_page_count += 1
                    p_bar.progress(current_page_count / total_pages_all)
            
            status_area.success("すべて完了しました！")
            
            out_ppt = io.BytesIO()
            prs.save(out_ppt)
            out_ppt.seek(0)
            
            # ファイル名生成（複数ならCombined、単体ならその名前）
            if len(docs) > 1:
                dl_name = "Combined_Slides.pptx"
            else:
                dl_name = f"{os.path.splitext(docs[0][0])[0]}_slide.pptx"
            
            st.download_button(
                label="📥 パワーポイントをダウンロード",
                data=out_ppt,
                file_name=dl_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
