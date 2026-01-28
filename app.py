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

# === 設定 ===
REMOVE_LOGO = True
ERASE_WIDTH = 350
ERASE_HEIGHT = 180

# OCRの前処理設定
OCR_ZOOM = 2.0  # 画像を何倍に拡大してOCRするか（2.0〜3.0が推奨）

st.set_page_config(page_title="PDF配置＆フォント復元ツール", layout="wide")

st.title("📏 フォントサイズも再現！高精度PDF変換ツール")
st.markdown("""
**機能強化ポイント：**
1. 画像を拡大・鮮明化してから読み取ることで**OCR精度**を向上。
2. 文字の大きさを解析し、**フォントサイズを自動調整**します。
""")

# PDFアップロード
uploaded_file = st.file_uploader("変換したいPDFファイルをアップロード", type="pdf")

def preprocess_image_for_ocr(cv_img, zoom_factor=2.0):
    """OCR精度を上げるための画像前処理関数"""
    # 1. 画像を拡大（小さい文字をつぶれないようにする）
    h, w = cv_img.shape[:2]
    new_w = int(w * zoom_factor)
    new_h = int(h * zoom_factor)
    resized = cv2.resize(cv_img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
    
    # 2. グレースケール化
    gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
    
    # 3. 二値化（白黒はっきりさせる・ノイズ除去）
    # 大津の二値化を使って自動で閾値を決定
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    return binary, new_w, new_h

if uploaded_file is not None:
    st.info(f"📄 {uploaded_file.name} を読み込みました。")
    
    bg_fill = st.checkbox("テキストボックスの背景を白く塗りつぶす", value=True)
    
    if st.button("高精度変換スタート"):
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        # PDFを画像変換
        # dpi=300 に上げることで元の画質を確保
        images = convert_from_bytes(uploaded_file.read(), dpi=300)
        total_pages = len(images)
        
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        progress_text.text("AIが画像を解析中...（拡大処理が入るため少し時間がかかります）")

        for i, image in enumerate(images):
            # 元画像の準備 (OpenCV形式)
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
            h_orig, w_orig = cv_img.shape[:2]

            # --- 1. OCR用の高画質画像を作成 ---
            # ここで拡大・二値化した画像を作ります
            ocr_img, w_ocr, h_ocr = preprocess_image_for_ocr(cv_img, OCR_ZOOM)

            # --- 2. 文字位置解析 (OCR) ---
            # 前処理した画像(ocr_img)を使って解析
            d = pytesseract.image_to_data(ocr_img, lang='jpn', output_type=Output.DICT)
            
            # --- 3. ブロック作成 ---
            blocks = {}
            n_boxes = len(d['text'])
            
            for j in range(n_boxes):
                if int(d['conf'][j]) > 0 and d['text'][j].strip() != "":
                    b_id = d['block_num'][j]
                    
                    if b_id not in blocks:
                        blocks[b_id] = {
                            'text': [],
                            'left': [], 'top': [], 'width': [], 'height': []
                        }
                    
                    # 座標データを蓄積
                    blocks[b_id]['text'].append(d['text'][j])
                    blocks[b_id]['left'].append(d['left'][j])
                    blocks[b_id]['top'].append(d['top'][j])
                    blocks[b_id]['width'].append(d['width'][j])
                    blocks[b_id]['height'].append(d['height'][j])

            # --- ロゴ消し (元画像に対して処理) ---
            if REMOVE_LOGO:
                mask = np.zeros((h_orig, w_orig), np.uint8)
                cv2.rectangle(mask, (w_orig - ERASE_WIDTH, h_orig - ERASE_HEIGHT), (w_orig, h_orig), 255, -1)
                cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

            # 背景画像保存
            image_stream = io.BytesIO()
            is_success, buffer = cv2.imencode(".jpg", cv_img)
            image_stream.write(buffer)
            
            # --- スライド作成 ---
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            # --- スライド座標計算用スケール ---
            # 元画像のサイズに対するパワポの倍率
            scale_ppt_x = prs.slide_width / w_orig
            scale_ppt_y = prs.slide_height / h_orig

            # --- 4. テキストボックス配置とフォントサイズ計算 ---
            for b_id, b_data in blocks.items():
                text_content = "".join(b_data['text'])
                
                # --- 座標の計算（重要：OCR画像は拡大されているので縮尺を戻す）---
                # 1. OCR画像上の座標を取得
                ocr_x = min(b_data['left'])
                ocr_y = min(b_data['top'])
                ocr_right = max([l + w for l, w in zip(b_data['left'], b_data['width'])])
                ocr_bottom = max([t + h for t, h in zip(b_data['top'], b_data['height'])])
                ocr_w = ocr_right - ocr_x
                ocr_h = ocr_bottom - ocr_y

                # 2. 元画像（拡大前）のサイズに戻す
                orig_x = ocr_x / OCR_ZOOM
                orig_y = ocr_y / OCR_ZOOM
                orig_w = ocr_w / OCR_ZOOM
                orig_h = ocr_h / OCR_ZOOM

                # --- フォントサイズの推定ロジック ---
                # ブロック内の「各行の高さ」の平均を計算するのが一番自然
                # ここでは簡易的に「ブロック全体の高さ / 行数」などを考慮したいが、
                # 単純に「文字の高さの平均」を計算してみます
                
                # 個々の文字の高さの平均値(px)を計算（拡大前のサイズで）
                avg_char_height_px = (sum(b_data['height']) / len(b_data['height'])) / OCR_ZOOM
                
                # ピクセルをポイント(pt)に変換する魔法の計算式
                # PPTの高さ(インチ) * 72(dpi) * (文字高さpx / 画像高さpx)
                estimated_pt = (prs.slide_height.inches * 72) * (avg_char_height_px / h_orig)
                
                # 補正係数（OCRの枠は文字より少し大きいため、0.7〜0.8倍すると適正に見える）
                font_size_pt = estimated_pt * 0.75
                
                # 安全装置：極端に小さい/大きいフォントを防ぐ
                if font_size_pt < 8: font_size_pt = 8
                if font_size_pt > 100: font_size_pt = 100

                # 3. パワポ上の座標に変換
                pp_x = int(orig_x * scale_ppt_x)
                pp_y = int(orig_y * scale_ppt_y)
                pp_w = int(orig_w * scale_ppt_x)
                pp_h = int(orig_h * scale_ppt_y)
                
                # ゴミ除去
                if pp_w > Inches(0.2) and pp_h > Inches(0.1):
                    try:
                        txBox = slide.shapes.add_textbox(pp_x, pp_y, pp_w, pp_h)
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        p = tf.paragraphs[0]
                        p.text = text_content
                        
                        # ★ここで推定したフォントサイズを適用！
                        p.font.size = Pt(font_size_pt)
                        
                        if bg_fill:
                            fill = txBox.fill
                            fill.solid()
                            fill.fore_color.rgb = RGBColor(255, 255, 255)
                    except:
                        pass

            progress_bar.progress((i + 1) / total_pages)
            progress_text.text(f"{i+1} / {total_pages} ページ完了...")

        output_pptx = io.BytesIO()
        prs.save(output_pptx)
        output_pptx.seek(0)
        
        st.success("✨ 高精度変換完了！")
        
        new_filename = uploaded_file.name.replace(".pdf", "_AI解析版.pptx")
        st.download_button(
            label="パワーポイントをダウンロード 📥",
            data=output_pptx,
            file_name=new_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
