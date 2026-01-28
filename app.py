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

# === ページ設定 ===
st.set_page_config(page_title="PDF完全分解ツール", layout="wide")

st.title("🧩 バラバラに分解！PDFオブジェクト化ツール")
st.markdown("""
**新機能：**
PDFを「一枚絵」として貼るのではなく、**「図表」と「文字」を自動で切り分けて配置**します。
グラフや写真が個別のパーツになるので、配置換えが可能です。
""")

# --- サイドバー設定 ---
st.sidebar.header("🔧 分解設定")
mode = st.sidebar.radio("変換モード", ["通常モード（背景画像＋文字）", "分解モード（図表切り抜き＋文字）"])

# 分解モード用の設定
st.sidebar.subheader("分解感度")
min_area_size = st.sidebar.slider("最小オブジェクトサイズ", 1000, 20000, 5000, help="小さいゴミを無視する基準")
dilation_iter = st.sidebar.slider("結合強度", 1, 10, 3, help="バラバラの線を1つの図としてまとめる強さ")

# 共通設定
st.sidebar.subheader("共通設定")
detect_color = st.sidebar.checkbox("文字色を自動検出", value=True)
bg_fill = st.sidebar.checkbox("文字背景を白くする", value=False) # 分解モードならFalse推奨

uploaded_file = st.file_uploader("PDFをアップロード", type="pdf")

# --- 画像処理関数群 ---

def preprocess_image_for_ocr(cv_img, zoom_factor=2.0):
    h, w = cv_img.shape[:2]
    new_w, new_h = int(w * zoom_factor), int(h * zoom_factor)
    resized = cv2.resize(cv_img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary, new_w, new_h

def get_dominant_color(img_crop):
    try:
        pixels = np.float32(img_crop.reshape(-1, 3))
        mask = np.all(pixels < 240, axis=1) # 白以外
        target_pixels = pixels[mask]
        if len(target_pixels) > 0:
            avg = np.mean(target_pixels, axis=0)
            return int(avg[0]), int(avg[1]), int(avg[2])
        return 0, 0, 0
    except:
        return 0, 0, 0

def extract_objects(cv_img, min_area=5000, dilation=3):
    """
    画像から「図表っぽい部分」の座標リストを返す関数
    """
    # 1. グレースケール＆二値化
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    # 文字や線をくっきりさせる（反転させる：背景黒、物体白）
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    # 2. 膨張処理（バラバラの文字や線をくっつけて「かたまり」にする）
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (10, 10))
    dilated = cv2.dilate(thresh, kernel, iterations=dilation)

    # 3. 輪郭抽出
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    objects = []
    h_img, w_img = cv_img.shape[:2]

    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        
        # 指定サイズより小さいゴミは無視
        if area < min_area:
            continue
            
        # 画面全体を覆うような枠（ただの外枠）は無視
        if w > w_img * 0.9 and h > h_img * 0.9:
            continue

        objects.append((x, y, w, h))
    
    return objects

# --- メイン処理 ---
if uploaded_file is not None:
    st.info(f"📄 {uploaded_file.name} を読み込みました。")
    
    if st.button("変換スタート"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        images = convert_from_bytes(uploaded_file.read(), dpi=300)
        total_pages = len(images)
        
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        for i, image in enumerate(images):
            status_text.text(f"{i+1}/{total_pages} ページ目を解析中...")
            
            # 画像準備
            cv_img = np.array(image)
            cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
            h_orig, w_orig = cv_img.shape[:2]
            scale_ppt_x = prs.slide_width / w_orig
            scale_ppt_y = prs.slide_height / h_orig

            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # === モード分岐 ===
            if mode == "通常モード（背景画像＋文字）":
                # 今まで通りのやり方（一枚絵を貼る）
                image_stream = io.BytesIO()
                is_success, buffer = cv2.imencode(".jpg", cv_img)
                image_stream.write(buffer)
                slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            else:
                # ★分解モード（背景は白、図だけ切り抜いて貼る）
                # 1. オブジェクト（図表）の検出と貼り付け
                objects = extract_objects(cv_img, min_area=min_area_size, dilation=dilation_iter)
                
                # オブジェクト領域を記憶（あとで文字が重なっているか判定するため）
                object_rects = [] 

                for (ox, oy, ow, oh) in objects:
                    # 切り抜き
                    crop_img = cv_img[oy:oy+oh, ox:ox+ow]
                    
                    # PPT配置
                    img_stream = io.BytesIO()
                    cv2.imencode(".png", crop_img)[1].tofile(img_stream)
                    
                    pp_x = int(ox * scale_ppt_x)
                    pp_y = int(oy * scale_ppt_y)
                    pp_w = int(ow * scale_ppt_x)
                    pp_h = int(oh * scale_ppt_y)
                    
                    slide.shapes.add_picture(img_stream, pp_x, pp_y, width=pp_w, height=pp_h)
                    object_rects.append((ox, oy, ow, oh))

            # === 文字の配置（共通処理） ===
            OCR_ZOOM = 2.0
            ocr_img, w_ocr, h_ocr = preprocess_image_for_ocr(cv_img, OCR_ZOOM)
            d = pytesseract.image_to_data(ocr_img, lang='jpn', output_type=Output.DICT)
            
            blocks = {}
            n_boxes = len(d['text'])
            for j in range(n_boxes):
                text = d['text'][j].strip()
                conf = int(d['conf'][j])
                if conf > 40 and text != "" and not (len(text)==1 and text in ".,-_"):
                    b_id = d['block_num'][j]
                    if b_id not in blocks:
                        blocks[b_id] = {'text': [], 'left': [], 'top': [], 'width': [], 'height': []}
                    blocks[b_id]['text'].append(text)
                    blocks[b_id]['left'].append(d['left'][j])
                    blocks[b_id]['top'].append(d['top'][j])
                    blocks[b_id]['width'].append(d['width'][j])
                    blocks[b_id]['height'].append(d['height'][j])

            for b_id, b_data in blocks.items():
                text_content = "".join(b_data['text'])
                
                ocr_x = min(b_data['left'])
                ocr_y = min(b_data['top'])
                ocr_w = max([l+w for l,w in zip(b_data['left'], b_data['width'])]) - ocr_x
                ocr_h = max([t+h for t,h in zip(b_data['top'], b_data['height'])]) - ocr_y
                
                orig_x, orig_y = int(ocr_x/OCR_ZOOM), int(ocr_y/OCR_ZOOM)
                orig_w, orig_h = int(ocr_w/OCR_ZOOM), int(ocr_h/OCR_ZOOM)

                # 分解モードの場合、図表エリアと重なっている文字は「図の中の文字」として無視する
                # （二重に表示されるのを防ぐため）
                is_inside_object = False
                if mode == "分解モード（図表切り抜き＋文字）":
                    center_x = orig_x + orig_w/2
                    center_y = orig_y + orig_h/2
                    for (ox, oy, ow, oh) in object_rects:
                        if ox < center_x < ox+ow and oy < center_y < oy+oh:
                            is_inside_object = True
                            break
                
                if is_inside_object:
                    continue

                avg_h = (sum(b_data['height'])/len(b_data['height']))/OCR_ZOOM
                est_pt = (prs.slide_height.inches * 72) * (avg_h/h_orig) * 0.75
                font_pt = max(8, min(est_pt * 1.0, 100)) # 倍率は調整可

                font_rgb = (0,0,0)
                if detect_color:
                    y1, y2 = max(0, orig_y), min(h_orig, orig_y + orig_h)
                    x1, x2 = max(0, orig_x), min(w_orig, orig_x + orig_w)
                    roi = cv_img[y1:y2, x1:x2]
                    if roi.size > 0:
                        r, g, b = get_dominant_color(roi)
                        font_rgb = (r, g, b)

                pp_x = int(orig_x * scale_ppt_x)
                pp_y = int(orig_y * scale_ppt_y)
                pp_w = int(orig_w * scale_ppt_x)

                if pp_w > Inches(0.2):
                    try:
                        txBox = slide.shapes.add_textbox(pp_x, pp_y, pp_w, Inches(font_pt/72*1.5))
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        p = tf.paragraphs[0]
                        p.text = text_content
                        p.font.size = Pt(font_pt)
                        p.font.color.rgb = RGBColor(*font_rgb)
                        
                        if bg_fill:
                            fill = txBox.fill
                            fill.solid()
                            fill.fore_color.rgb = RGBColor(255, 255, 255)
                    except: pass

            progress_bar.progress((i + 1) / total_pages)

        output_pptx = io.BytesIO()
        prs.save(output_pptx)
        output_pptx.seek(0)
        
        st.success("🎉 分解完了！")
        st.download_button("ダウンロード 📥", output_pptx, uploaded_file.name.replace(".pdf", "_分解版.pptx"))
