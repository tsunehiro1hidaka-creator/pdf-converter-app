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
from pptx.enum.text import MSO_ANCHOR
import io
import zipfile
# ★新機能：画像比較スライダー
from streamlit_image_comparison import image_comparison

# ===========================
# ページ設定
# ===========================
st.set_page_config(page_title="Universe PDF Converter", layout="wide", initial_sidebar_state="expanded")

# UIカスタマイズ
st.markdown("""
<style>
    .stButton>button { border-radius: 8px; font-weight: bold; border: 2px solid #4CAF50; }
    .stProgress .st-bo { background-color: #4CAF50; }
    h1 { color: #2E7D32; }
</style>
""", unsafe_allow_html=True)

st.title("🪐 Universe PDF Converter (宇宙一エディション)")
st.markdown("""
**人類未踏の機能:**
1. **比較スライダー**: 処理前後の違いを一目瞭然に。
2. **スマート・カラム検知**: 2段組みの論文レイアウトも正しく読み取ります。
3. **素材抽出**: 切り抜いた図表だけをZIPで持ち帰れます。
""")

# ===========================
# 関数定義（キャッシュ & ロジック）
# ===========================

@st.cache_data(show_spinner=False)
def load_pdf_images(file_bytes):
    return convert_from_bytes(file_bytes, dpi=300)

@st.cache_data(show_spinner=False)
def get_pdf_info(file_bytes):
    return pdfinfo_from_bytes(file_bytes)

def preprocess_image_for_ocr(cv_img, zoom_factor=2.0):
    h, w = cv_img.shape[:2]
    new_w, new_h = int(w * zoom_factor), int(h * zoom_factor)
    resized = cv2.resize(cv_img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    enhanced = clahe.apply(gray)
    _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary, new_w, new_h

def extract_objects(cv_img, min_area=5000):
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (10, 10))
    dilated = cv2.dilate(thresh, kernel, iterations=3)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    objects = []
    h, w = cv_img.shape[:2]
    for cnt in contours:
        x, y, cw, ch = cv2.boundingRect(cnt)
        if cw * ch < min_area: continue
        if cw > w * 0.95 or ch > h * 0.95: continue 
        objects.append((x, y, cw, ch))
    return objects

def sort_blocks_smart(blocks):
    """
    ★宇宙一機能：スマートソート
    2段組みレイアウトに対応するため、単純な「上から順」ではなく、
    「左カラムの上から下」→「右カラムの上から下」へ並べ替える
    """
    if not blocks: return []
    
    # 全ブロックのリスト化
    block_list = []
    for bid, data in blocks.items():
        data['id'] = bid
        data['avg_x'] = min(data['left']) # 左端座標
        data['avg_y'] = min(data['top'])  # 上端座標
        block_list.append(data)
    
    # ページの中央線（X座標）を計算
    if not block_list: return []
    max_x = max([b['avg_x'] + max(b['width']) for b in block_list])
    center_x = max_x / 2

    # カラム分け（簡易判定：中央より左か右か）
    left_col = []
    right_col = []
    
    for b in block_list:
        if b['avg_x'] < center_x: left_col.append(b)
        else: right_col.append(b)
        
    # 各カラム内で、Y座標（上から順）にソート
    left_col.sort(key=lambda x: x['avg_y'])
    right_col.sort(key=lambda x: x['avg_y'])
    
    # 結合して返す（左カラム全部 → 右カラム全部）
    # ※もし1段組みならleft_colに全部入るので問題なし
    return left_col + right_col

def merge_text_blocks_ordered(sorted_blocks, spacing_limit=50):
    """ソート済みのリストを受け取って結合する"""
    if not sorted_blocks: return {}
    
    merged = {}
    curr_m_id = 0
    # 辞書のコピーを作成して初期化
    merged[curr_m_id] = {k: v[:] if isinstance(v, list) else v for k, v in sorted_blocks[0].items()}
    
    for i in range(1, len(sorted_blocks)):
        curr = sorted_blocks[i]
        prev = merged[curr_m_id]
        
        prev_bottom = max([t + h for t, h in zip(prev['top'], prev['height'])])
        curr_top = min(curr['top'])
        prev_left = min(prev['left'])
        curr_left = min(curr['left'])
        
        # 結合判定：同じカラム（左端が近い）かつ、行間が狭い
        if 0 < (curr_top - prev_bottom) < spacing_limit and abs(curr_left - prev_left) < 60:
            prev['text'].append("\n")
            prev['text'].extend(curr['text'])
            prev['left'].extend(curr['left'])
            prev['top'].extend(curr['top'])
            prev['width'].extend(curr['width'])
            prev['height'].extend(curr['height'])
        else:
            curr_m_id += 1
            merged[curr_m_id] = {k: v[:] if isinstance(v, list) else v for k, v in curr.items()}
            
    return merged

# ===========================
# サイドバー設定
# ===========================
st.sidebar.header("🎛️ コックピット")

with st.sidebar.expander("🛠️ 基本設定", expanded=True):
    target_font = st.selectbox("出力フォント", ["Meiryo", "Yu Gothic", "BIZ UDPGothic", "MS PGothic"])
    mode = st.radio("モード選択", ["分解モード（図形分離）", "通常モード（一枚絵）"])
    smart_sort = st.checkbox("2段組みレイアウト補正", value=True, help="論文などの2列文章を正しく読み取ります")

with st.sidebar.expander("✂️ 加工設定", expanded=True):
    use_erase = st.checkbox("ロゴ消し有効", value=True)
    erase_w = st.slider("右端カット (px)", 0, 800, 350, step=10)
    erase_h = st.slider("下端カット (px)", 0, 500, 180, step=10)

with st.sidebar.expander("⚙️ 詳細設定"):
    min_area = st.slider("図形認識感度", 1000, 20000, 5000)
    jpeg_q = st.slider("画質", 10, 100, 85)
    merge_lines = st.checkbox("段落結合", value=True)

# ===========================
# メイン処理
# ===========================
uploaded_file = st.file_uploader("PDFファイルをドロップ", type="pdf")

if uploaded_file is not None:
    # 1. ロード
    file_bytes = uploaded_file.read()
    pdf_info = get_pdf_info(file_bytes)
    total_pages = pdf_info["Pages"]
    
    with st.spinner("解析エンジン起動中..."):
        images = load_pdf_images(file_bytes)

    # 2. ★宇宙一機能：ビフォーアフター比較スライダー
    st.divider()
    st.subheader("👁️ 処理結果プレビュー")
    
    preview_idx = st.number_input("確認ページ", 1, total_pages, 1) - 1
    
    # 比較用画像の作成
    img_original = np.array(images[preview_idx]) # 元画像
    img_processed = np.array(images[preview_idx]) # 処理後画像
    img_processed = cv2.cvtColor(img_processed, cv2.COLOR_RGB2BGR) # OpenCV用
    h, w = img_processed.shape[:2]

    # 加工処理（プレビュー用）
    if use_erase:
        mask = np.zeros((h, w), np.uint8)
        cv2.rectangle(mask, (w - erase_w, h - erase_h), (w, h), 255, -1)
        img_processed = cv2.inpaint(img_processed, mask, 3, cv2.INPAINT_TELEA)

    # 図形枠の描画
    if mode.startswith("分解"):
        objs = extract_objects(img_processed, min_area)
        for (ox, oy, ow, oh) in objs:
            cv2.rectangle(img_processed, (ox, oy), (ox+ow, oy+oh), (255, 120, 0), 4) # 青枠
    
    # RGBに戻す
    img_processed = cv2.cvtColor(img_processed, cv2.COLOR_BGR2RGB)

    # ★スライダー表示
    image_comparison(
        img1=img_original,
        img2=img_processed,
        label1="元のPDF",
        label2="変換イメージ (赤:消去 / 青:図形)",
        width=700,
        starting_position=50,
        show_labels=True,
        make_responsive=True,
        in_memory=True
    )

    # 3. 変換実行
    st.divider()
    col_a, col_b = st.columns([2, 1])
    with col_a:
        page_range = st.slider("変換範囲", 1, total_pages, (1, min(total_pages, 10)))
    with col_b:
        st.write("")
        start_btn = st.button("🚀 変換エンジン点火", type="primary", use_container_width=True)

    if start_btn:
        start_p, end_p = page_range
        process_cnt = end_p - start_p + 1
        p_bar = st.progress(0)
        status = st.empty()

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        # ★素材抽出用のZIPファイル準備
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            
            target_images = images[start_p-1 : end_p]
            
            for i, image in enumerate(target_images):
                p_num = start_p + i
                status.text(f"Processing P.{p_num} ...")
                
                # 画像準備
                cv_img = np.array(image)
                cv_img = cv2.cvtColor(cv_img, cv2.COLOR_RGB2BGR)
                h_orig, w_orig = cv_img.shape[:2]
                scale_x = prs.slide_width / w_orig
                scale_y = prs.slide_height / h_orig
                
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                
                # --- ロゴ消し ---
                if use_erase:
                    mask = np.zeros((h_orig, w_orig), np.uint8)
                    cv2.rectangle(mask, (w_orig - erase_w, h_orig - erase_h), (w_orig, h_orig), 255, -1)
                    cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

                object_rects = []
                
                # --- 図形処理 & 素材ZIPへの保存 ---
                if mode.startswith("分解"):
                    objs = extract_objects(cv_img, min_area)
                    for idx, (ox, oy, ow, oh) in enumerate(objs):
                        crop = cv_img[oy:oy+oh, ox:ox+ow]
                        # ZIPに保存
                        crop_bytes = cv2.imencode(".jpg", crop)[1].tobytes()
                        zf.writestr(f"images/P{p_num}_img{idx+1}.jpg", crop_bytes)
                        
                        # PPT配置
                        s = io.BytesIO(crop_bytes)
                        slide.shapes.add_picture(s, int(ox*scale_x), int(oy*scale_y), width=int(ow*scale_x), height=int(oh*scale_y))
                        object_rects.append((ox, oy, ow, oh))
                else:
                    # 通常モード
                    full_bytes = cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), jpeg_q])[1].tobytes()
                    s = io.BytesIO(full_bytes)
                    slide.shapes.add_picture(s, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # --- OCR処理 ---
                ocr_img, _, _ = preprocess_image_for_ocr(cv_img, 2.0)
                d = pytesseract.image_to_data(ocr_img, lang='jpn', output_type=Output.DICT)
                
                raw_blocks = {}
                for j in range(len(d['text'])):
                    txt = d['text'][j].strip()
                    if int(d['conf'][j]) > 40 and txt != "" and not (len(txt)==1 and txt in ".,-_|"):
                        bid = d['block_num'][j]
                        if bid not in raw_blocks: raw_blocks[bid] = {'text':[], 'left':[], 'top':[], 'width':[], 'height':[]}
                        raw_blocks[bid]['text'].append(txt)
                        raw_blocks[bid]['left'].append(d['left'][j])
                        raw_blocks[bid]['top'].append(d['top'][j])
                        raw_blocks[bid]['width'].append(d['width'][j])
                        raw_blocks[bid]['height'].append(d['height'][j])
                
                # ★宇宙一ロジック：スマートソート＆結合
                if smart_sort:
                    sorted_list = sort_blocks_smart(raw_blocks) # 左カラム→右カラムの順に並び替え
                    final_blocks = merge_text_blocks_ordered(sorted_list) if merge_lines else {i:v for i,v in enumerate(sorted_list)}
                else:
                    # 従来の結合（上から順）
                    # 辞書をリスト化してソートしてから結合関数へ
                    sorted_list = sorted(raw_blocks.values(), key=lambda x: min(x['top']))
                    final_blocks = merge_text_blocks_ordered(sorted_list) if merge_lines else raw_blocks

                # --- テキスト配置 ---
                for bid, bdata in final_blocks.items():
                    content = "".join(bdata['text'])
                    
                    ox = min(bdata['left'])
                    oy = min(bdata['top'])
                    ow = max([l+w for l,w in zip(bdata['left'], bdata['width'])]) - ox
                    oh = max([t+h for t,h in zip(bdata['top'], bdata['height'])]) - oy
                    
                    orig_x, orig_y, orig_w, orig_h = int(ox/2), int(oy/2), int(ow/2), int(oh/2)

                    if mode.startswith("分解"):
                        cx, cy = orig_x + orig_w/2, orig_y + orig_h/2
                        if any(rox < cx < rox+row and roy < cy < roy+roh for (rox,roy,row,roh) in object_rects): continue

                    pp_x = int(orig_x * scale_x)
                    pp_y = int(orig_y * scale_y)
                    pp_w = int(orig_w * scale_x)

                    if pp_w > Inches(0.2):
                        try:
                            avg_h = (sum(bdata['height'])/len(bdata['height']))/2.0
                            font_pt = max(9, min((prs.slide_height.inches*72)*(avg_h/h_orig)*0.75, 80))
                            
                            box_h = Inches(font_pt/72 * 1.5 * (content.count('\n') + 1))
                            txBox = slide.shapes.add_textbox(pp_x, pp_y, pp_w, box_h)
                            tf = txBox.text_frame
                            tf.word_wrap = True
                            p = tf.paragraphs[0]
                            p.text = content
                            p.font.size = Pt(font_pt)
                            p.font.name = target_font
                            p.font.language_id = 1041 # JP
                            
                            # 色検出
                            y1,y2 = max(0,orig_y), min(h_orig,orig_y+orig_h)
                            x1,x2 = max(0,orig_x), min(w_orig,orig_x+orig_w)
                            roi = cv_img[y1:y2, x1:x2]
                            if roi.size > 0:
                                pixels = np.float32(roi.reshape(-1, 3))
                                mask = np.all(pixels < 230, axis=1)
                                t_pix = pixels[mask]
                                if len(t_pix)>0:
                                    r,g,b = np.mean(t_pix, axis=0)
                                    p.font.color.rgb = RGBColor(int(r), int(g), int(b))
                        except: pass
                
                p_bar.progress((i+1)/process_cnt)

        # 完了
        status.empty()
        st.balloons()
        st.success("🪐 宇宙一の変換が完了しました！")
        
        col_d1, col_d2 = st.columns(2)
        
        # PPTXダウンロード
        out_ppt = io.BytesIO()
        prs.save(out_ppt)
        out_ppt.seek(0)
        col_d1.download_button(f"📥 パワポをダウンロード", out_ppt, "Universe_Presentation.pptx", type="primary", use_container_width=True)

        # ZIPダウンロード（素材）
        if mode.startswith("分解"):
            zip_buffer.seek(0)
            col_d2.download_button(f"🗂️ 画像素材ZIPをダウンロード", zip_buffer, "images.zip", use_container_width=True)
        else:
            col_d2.info("※分解モードにすると画像素材をDLできます")
