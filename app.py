import streamlit as st
import os
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches
import io

# === 設定 ===
REMOVE_LOGO = True
ERASE_WIDTH = 350
ERASE_HEIGHT = 180

st.title("📄 PDF → パワポ変換ツール（全部入り）")
st.write("PDFをアップロードすると、ロゴを消して、文字を読み取って、パワポにします。")

# PDFアップロード
uploaded_file = st.file_uploader("ここにPDFをドラッグ＆ドロップ", type="pdf")

if uploaded_file is not None:
    st.info("ファイルを読み込みました。変換ボタンを押してください。")
    
    if st.button("変換スタート！"):
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        # PDFをメモリ上で読み込む
        images = convert_from_bytes(uploaded_file.read())
        total_pages = len(images)
        
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        progress_text.text("変換中...おみくじ：大吉🌸")

        for i, image in enumerate(images):
            # 画像処理
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
            h, w = cv_img.shape[:2]

            # 文字読み取り
            try:
                extracted_text = pytesseract.image_to_string(cv_img, lang='jpn')
            except:
                extracted_text = "読み取り失敗"

            # ロゴ消し
            if REMOVE_LOGO:
                mask = np.zeros((h, w), np.uint8)
                cv2.rectangle(mask, (w - ERASE_WIDTH, h - ERASE_HEIGHT), (w, h), 255, -1)
                cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

            # 一時保存せずにバイナリとして扱う
            image_stream = io.BytesIO()
            is_success, buffer = cv2.imencode(".jpg", cv_img)
            image_stream.write(buffer)
            
            # スライド作成
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
            slide.notes_slide.notes_text_frame.text = f"【P.{i+1} テキスト】\n{extracted_text}"

            # 進捗更新
            progress_bar.progress((i + 1) / total_pages)
            progress_text.text(f"{i+1} / {total_pages} ページ完了...")

        # PPTXをメモリに保存
        output_pptx = io.BytesIO()
        prs.save(output_pptx)
        output_pptx.seek(0)
        
        st.success("✨ 完了しました！下のボタンからダウンロードできます。")
        
        # ダウンロードボタン
        new_filename = uploaded_file.name.replace(".pdf", "_変換済.pptx")
        st.download_button(
            label="パワーポイントをダウンロード",
            data=output_pptx,
            file_name=new_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
