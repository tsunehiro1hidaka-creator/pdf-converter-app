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

# ページ設定
st.set_page_config(page_title="PDF配置復元ツール", layout="wide")

st.title("🧩 配置も復元！PDF → パワポ変換ツール")
st.markdown("""
PDFの文字位置を解析し、**元のレイアウトに近い場所にテキストボックスを自動配置**します。
画像の上に文字を重ねるため、まるで「翻訳こんにゃく」のように編集できます。
""")

# PDFアップロード
uploaded_file = st.file_uploader("変換したいPDFファイルをアップロード", type="pdf")

if uploaded_file is not None:
    st.info(f"📄 {uploaded_file.name} を読み込みました。")
    
    # オプション
    bg_fill = st.checkbox("テキストボックスの背景を白く塗りつぶす（下の文字を隠す）", value=True)
    
    if st.button("配置を解析して変換スタート"):
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        # PDFを画像変換
        images = convert_from_bytes(uploaded_file.read())
        total_pages = len(images)
        
        # スライド準備 (16:9)
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        # 座標変換のための倍率計算（画像ピクセル → パワポEMU単位）
        # ※pdf2imageのデフォルト解像度を基準に計算
        # 処理する画像のサイズを取得してから計算するためループ内で調整

        progress_text.text("レイアウト解析中...")

        for i, image in enumerate(images):
            # 画像処理 (OpenCV)
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
            h_img, w_img = cv_img.shape[:2]

            # パワポとのサイズ比率を計算
            scale_x = prs.slide_width / w_img
            scale_y = prs.slide_height / h_img

            # --- 文字位置解析 (OCR) ---
            # image_to_data で文字と座標を一度に取得
            # output_type=Output.DICT で辞書形式で受け取る
            d = pytesseract.image_to_data(cv_img, lang='jpn', output_type=Output.DICT)
            
            # --- ブロックごとに文字をまとめる処理 ---
            # バラバラの単語を「段落（ブロック）」ごとに結合します
            blocks = {}
            n_boxes = len(d['text'])
            
            for j in range(n_boxes):
                # 信頼度が低い、または空白のデータはスキップ
                if int(d['conf'][j]) > 0 and d['text'][j].strip() != "":
                    b_id = d['block_num'][j] # ブロックID
                    
                    if b_id not in blocks:
                        blocks[b_id] = {
                            'text': [],
                            'left': [], 'top': [], 'width': [], 'height': []
                        }
                    
                    blocks[b_id]['text'].append(d['text'][j])
                    blocks[b_id]['left'].append(d['left'][j])
                    blocks[b_id]['top'].append(d['top'][j])
                    blocks[b_id]['width'].append(d['width'][j])
                    blocks[b_id]['height'].append(d['height'][j])

            # --- ロゴ消し処理 ---
            if REMOVE_LOGO:
                mask = np.zeros((h_img, w_img), np.uint8)
                cv2.rectangle(mask, (w_img - ERASE_WIDTH, h_img - ERASE_HEIGHT), (w_img, h_img), 255, -1)
                cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

            # 画像をスライド背景用に保存
            image_stream = io.BytesIO()
            is_success, buffer = cv2.imencode(".jpg", cv_img)
            image_stream.write(buffer)
            
            # --- スライド作成 ---
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # 白紙
            
            # 1. 背景画像を貼る
            slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            # 2. 解析したブロックごとにテキストボックス配置
            for b_id, b_data in blocks.items():
                # そのブロックの結合テキスト
                text_content = "".join(b_data['text'])
                
                # ブロック全体の座標範囲を計算
                x = min(b_data['left'])
                y = min(b_data['top'])
                # 幅と高さは、「右端 - 左端」「下端 - 上端」で計算
                right = max([l + w for l, w in zip(b_data['left'], b_data['width'])])
                bottom = max([t + h for t, h in zip(b_data['top'], b_data['height'])])
                w = right - x
                h = bottom - y
                
                # 座標をパワポ用に変換
                pp_x = int(x * scale_x)
                pp_y = int(y * scale_y)
                pp_w = int(w * scale_x)
                pp_h = int(h * scale_y) # 高さは少し余裕を持たせてもよい
                
                # テキストボックス作成
                # あまりに小さいゴミのようなブロックは無視（幅・高さが小さすぎる場合）
                if pp_w > Inches(0.2) and pp_h > Inches(0.1):
                    try:
                        txBox = slide.shapes.add_textbox(pp_x, pp_y, pp_w, pp_h)
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        p = tf.paragraphs[0]
                        p.text = text_content
                        p.font.size = Pt(10.5) # 標準的なフォントサイズ
                        
                        # オプション：背景を白く塗る（元の画像を隠すため）
                        if bg_fill:
                            fill = txBox.fill
                            fill.solid()
                            fill.fore_color.rgb = RGBColor(255, 255, 255)
                    except:
                        # エラー時はスキップ（座標計算ミスなど）
                        pass

            # 進捗更新
            progress_bar.progress((i + 1) / total_pages)
            progress_text.text(f"{i+1} / {total_pages} ページ完了...")

        # 保存とダウンロード
        output_pptx = io.BytesIO()
        prs.save(output_pptx)
        output_pptx.seek(0)
        
        st.success("✨ 配置復元完了！ダウンロードして確認してください。")
        
        new_filename = uploaded_file.name.replace(".pdf", "_配置版.pptx")
        st.download_button(
            label="パワーポイントをダウンロード 📥",
            data=output_pptx,
            file_name=new_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
