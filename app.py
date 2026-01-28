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
st.set_page_config(page_title="Pro PDF Converter Ultimate", layout="wide", initial_sidebar_state="expanded")

st.title("🚀 世界最高峰のPDF変換ツール")
st.markdown("""
PDFを解析し、レイアウト、文字、色、そして**図表オブジェクト**まで分解してPowerPoint化します。
実務で必要な「テンプレート適用」や「ページ指定」にも対応したプロ仕様モデルです。
""")

# ===========================
# サイドバー：設定エリア
# ===========================
st.sidebar.header("🎛️ プロフェッショナル設定")

# タブで設定を整理
tab1, tab2, tab3 = st.sidebar.tabs(["基本設定", "高度な調整", "出力設定"])

with tab1:
    st.subheader("変換モード")
    mode = st.radio("処理モード選択", ["分解モード（推奨）", "通常モード"], 
                    help="分解モードは図と文字を分離します。通常モードは背景を一枚絵として扱います。")
    
    st.subheader("テンプレート (任意)")
    template_file = st.file_uploader("会社のPPTXテンプレートを適用", type="pptx", help="アップロードすると、そのスライドマスターデザインが適用されます。")

with tab2:
    st.subheader("ロゴ/不要領域の削除")
    use_erase = st.checkbox("有効にする", value=True)
    col1, col2 = st.columns(2)
    with col1:
        erase_width = st.slider("右端カット(px)", 0, 800, 350, step=10)
    with col2:
        erase_height = st.slider("下端カット(px)", 0, 500, 180, step=10)
    
    st.subheader("分解・認識感度")
    min_area_size = st.slider("最小図形サイズ", 1000, 20000, 5000, help="これより小さい塊は図として認識しません")
    detect_color = st.checkbox("文字色の再現", value=True)

with tab3:
    st.subheader("ファイルサイズと画質")
    # JPEG品質 (0-100)
    jpeg_quality = st.slider("画像圧縮品質 (低画質← →高画質)", 10, 100, 80, help="値を下げるとファイルサイズが軽くなります")

# メインエリア：PDFアップロード
st.header("📄 PDFファイルをアップロード")
uploaded_file = st.file_uploader("", type="pdf", label_visibility="collapsed")

# ===========================
# 関数定義エリア
# ===========================
def preprocess_image_for_ocr(cv_img, zoom_factor=2.0):
    h, w = cv_img.shape[:2]
    new_w, new_h = int(w * zoom_factor), int(h * zoom_factor)
    resized = cv2.resize(cv_img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
    # コントラスト強調（CLAHE）を追加してさらに精度向上
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    enhanced = clahe.apply(gray)
    _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary, new_w, new_h

def get_dominant_color(img_crop):
    try:
        pixels = np.float32(img_crop.reshape(-1, 3))
        mask = np.all(pixels < 230, axis=1) # 白に近い色を除外強化
        target_pixels = pixels[mask]
        if len(target_pixels) > 0:
            avg = np.mean(target_pixels, axis=0)
            return int(avg[0]), int(avg[1]), int(avg[2])
        return 0, 0, 0
    except:
        return 0, 0, 0

def extract_objects(cv_img, min_area=5000, dilation=3):
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (10, 10))
    dilated = cv2.dilate(thresh, kernel, iterations=dilation)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    objects = []
    h_img, w_img = cv_img.shape[:2]
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        if area < min_area: continue
        if w > w_img * 0.95 and h > h_img * 0.95: continue # ほぼ全画面の枠は無視
        objects.append((x, y, w, h))
    return objects

# ===========================
# メイン処理
# ===========================
if uploaded_file is not None:
    # PDF情報の取得（総ページ数など）
    pdf_info = pdfinfo_from_bytes(uploaded_file.read())
    max_pages = pdf_info["Pages"]
    uploaded_file.seek(0) # ファイルポインタを戻す

    st.success(f"✅ 読み込み完了: 全 {max_pages} ページ")
    
    # ページ範囲指定スライダー（動的に最大値を設定）
    st.subheader("⚡ 処理範囲の選択")
    page_range = st.slider("変換するページ範囲を指定してください", 1, max_pages, (1, min(max_pages, 5)))
    start_page, end_page = page_range
    
    process_count = end_page - start_page + 1
    st.info(f"{start_page}ページ目 から {end_page}ページ目 まで（計 {process_count} ページ）を変換します。")

    if st.button("🔥 究極変換スタート"):
        # 進捗表示用
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("PDFを画像に変換中... (高解像度処理のため時間がかかります)")
        
        # 指定範囲のみ画像変換（first_page, last_pageを指定）
        images = convert_from_bytes(uploaded_file.read(), dpi=300, first_page=start_page, last_page=end_page)
        
        # テンプレートの適用
        if template_file:
            prs = Presentation(template_file)
            status_text.text("テンプレートファイルを適用しました。")
        else:
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

        # 共通スライドレイアウト（白紙またはテンプレートの最後のレイアウトを利用）
        blank_layout = prs.slide_layouts[len(prs.slide_layouts)-1]

        for i, image in enumerate(images):
            current_page_num = start_page + i
            status_text.markdown(f"**処理中: {i+1} / {process_count}枚目 (元PDFのP.{current_page_num})**")
            
            # OpenCV形式へ
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
            h_orig, w_orig = cv_img.shape[:2]
            
            # パワポスケール計算
            scale_ppt_x = prs.slide_width / w_orig
            scale_ppt_y = prs.slide_height / h_orig

            # スライド追加
            slide = prs.slides.add_slide(blank_layout)

            # === 図形・背景処理 ===
            object_rects = []
            status_text.text(f"P.{current_page_num}: 図表オブジェクトを分離中...")

            if mode == "通常モード":
                # 背景を一枚絵として圧縮して配置
                img_stream = io.BytesIO()
                # JPEG圧縮品質を適用
                cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_quality])[1].tofile(img_stream)
                slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            else: # 分解モード
                objects = extract_objects(cv_img, min_area=min_area_size)
                for (ox, oy, ow, oh) in objects:
                    crop_img = cv_img[oy:oy+oh, ox:ox+ow]
                    img_stream = io.BytesIO()
                    # 切り抜いた図は綺麗に見せたいのでPNG(可逆圧縮)または高画質JPEG
                    cv2.imencode(".jpg", crop_img, [int(cv2.IMWRITE_JPEG_QUALITY), max(90, jpeg_quality)])[1].tofile(img_stream)
                    
                    pp_x = int(ox * scale_ppt_x)
                    pp_y = int(oy * scale_ppt_y)
                    pp_w = int(ow * scale_ppt_x)
                    pp_h = int(oh * scale_ppt_y)
                    slide.shapes.add_picture(img_stream, pp_x, pp_y, width=pp_w, height=pp_h)
                    object_rects.append((ox, oy, ow, oh))

            # === 文字
