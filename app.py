import streamlit as st
import os
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
import io

# === 設定 ===
REMOVE_LOGO = True
ERASE_WIDTH = 350
ERASE_HEIGHT = 180

st.set_page_config(page_title="PDF変換＆編集ツール", layout="wide")

st.title("🏢 業務効率化：PDF → パワポ変換（編集モード付）")
st.markdown("""
PDFを読み込み、**編集可能なテキストボックス**を乗せた状態でPowerPointにします。
外部SaaSを使わず、Pythonのみで処理するためセキュアです。
""")

# PDFアップロード
uploaded_file = st.file_uploader("変換したいPDFファイルをアップロード", type="pdf")

if uploaded_file is not None:
    st.info(f"📄 {uploaded_file.name} を読み込みました。")
    
    # オプション選択
    add_text_layer = st.checkbox("文字をテキストボックスとして上に重ねる", value=True)
    
    if st.button("変換スタート"):
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        # PDFを画像変換
        images = convert_from_bytes(uploaded_file.read())
        total_pages = len(images)
        
        # スライド準備
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        progress_text.text("処理を開始します...")

        for i, image in enumerate(images):
            # 画像処理 (OpenCV)
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
            h, w = cv_img.shape[:2]

            # 文字読み取り (OCR)
            extracted_text = ""
            try:
                extracted_text = pytesseract.image_to_string(cv_img, lang='jpn')
            except:
                extracted_text = "（文字読み取りに失敗しました）"

            # ロゴ消し
            if REMOVE_LOGO:
                mask = np.zeros((h, w), np.uint8)
                cv2.rectangle(mask, (w - ERASE_WIDTH, h - ERASE_HEIGHT), (w, h), 255, -1)
                cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

            # 画像をストリームに変換
            image_stream = io.BytesIO()
            is_success, buffer = cv2.imencode(".jpg", cv_img)
            image_stream.write(buffer)
            
            # --- スライド作成 ---
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # 白紙スライド
            
            # 1. 背景画像を貼る
            slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            # 2. テキストボックスを重ねる（ここがポイント！）
            if add_text_layer and extracted_text.strip():
                # 左上から、少し余白を空けてボックスを作成
                left = Inches(0.5)
                top = Inches(0.5)
                width = Inches(12.3)
                height = Inches(6.5)
                
                txBox = slide.shapes.add_textbox(left, top, width, height)
                tf = txBox.text_frame
                tf.word_wrap = True # 折り返しあり
                
                # ボックスの中身を設定
                p = tf.paragraphs[0]
                p.text = extracted_text
                p.font.size = Pt(14) # フォントサイズ
                p.font.color.rgb = RGBColor(0, 0, 0) # 文字色：黒
                
                # 【視認性向上】テキストボックスの背景を「半透明の白」にする
                # ※python-pptxで完全な半透明は難しいので、塗りつぶし設定を行います
                fill = txBox.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(255, 255, 255) # 白
                # ※透過度はPowerPoint上で調整が必要ですが、一旦白背景で視認性を確保します

            # ノートにも念のため残す
            slide.notes_slide.notes_text_frame.text = f"【抽出テキスト】\n{extracted_text}"

            # 進捗更新
            progress_bar.progress((i + 1) / total_pages)
            progress_text.text(f"{i+1} / {total_pages} ページ完了...")

        # 保存とダウンロード
        output_pptx = io.BytesIO()
        prs.save(output_pptx)
        output_pptx.seek(0)
        
        st.success("✨ 変換完了！以下のボタンからダウンロードしてください。")
        
        new_filename = uploaded_file.name.replace(".pdf", "_編集用.pptx")
        st.download_button(
            label="パワーポイントをダウンロード 📥",
            data=output_pptx,
            file_name=new_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
