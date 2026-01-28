# === 変更ここから ===
    if btn:
        start_p, end_p = page_range
        process_cnt = end_p - start_p + 1
        
        # ★修正1: status.empty() ではなく、プログレスバーだけ先に作る
        p_bar = st.progress(0)
        
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        zip_buffer = io.BytesIO()

        # ★修正2: 処理状況を表示する専用の場所を作る（markdownで固定）
        status_area = st.empty()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            target_images = images[start_p-1 : end_p]
            
            for i, image in enumerate(target_images):
                p_num = start_p + i
                
                # ★修正3: ここで .text() や .empty() を混ぜず、常に markdown で上書きする
                status_area.markdown(f"**⏳ 処理中... {i+1}/{process_cnt} 枚目 (P.{p_num})**")
                
                # --- (中略：画像処理やOCRのロジックはそのまま) ---
                # ※ここは長いので元のコードのまま、いじらなくてOKです
                
                cv_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                h_orig, w_orig = cv_img.shape[:2]
                scale_x = prs.slide_width / w_orig
                scale_y = prs.slide_height / h_orig
                
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # 1. 画像配置
                if use_erase:
                    mask = np.zeros((h_orig, w_orig), np.uint8)
                    cv2.rectangle(mask, (w_orig - erase_w, h_orig - erase_h), (w_orig, h_orig), 255, -1)
                    cv_img = cv2.inpaint(cv_img, mask, 3, cv2.INPAINT_TELEA)

                object_rects = []
                if mode.startswith("分解"):
                    objs = extract_objects(cv_img, min_area)
                    for idx, (ox, oy, ow, oh) in enumerate(objs):
                        crop = cv_img[oy:oy+oh, ox:ox+ow]
                        zf.writestr(f"assets/P{p_num}_fig{idx}.jpg", cv2.imencode(".jpg", crop)[1].tobytes())
                        s = io.BytesIO(cv2.imencode(".jpg", crop)[1].tobytes())
                        slide.shapes.add_picture(s, int(ox*scale_x), int(oy*scale_y), width=int(ow*scale_x), height=int(oh*scale_y))
                        object_rects.append((ox, oy, ow, oh))
                else:
                    s = io.BytesIO(cv2.imencode(".jpg", cv_img, [int(cv2.IMWRITE_JPEG_QUALITY), 80])[1].tofile(s))
                    slide.shapes.add_picture(s, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 2. OCR & 翻訳
                ocr_img, _, _ = preprocess_image_for_ocr(cv_img, 2.0)
                try:
                    d = pytesseract.image_to_data(ocr_img, lang=ocr_lang, output_type=Output.DICT)
                except Exception:
                    d = {'text': [], 'conf': []} # エラー回避

                raw_blocks = {}
                if 'text' in d:
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
                
                sorted_list = sort_blocks_smart(raw_blocks) if smart_sort else sorted(raw_blocks.values(), key=lambda x: min(x['top']))
                final_blocks = merge_text_blocks_ordered(sorted_list) if merge_lines else {i:v for i,v in enumerate(sorted_list)}

                # 3. テキスト配置
                notes_content = []
                for bid, bdata in final_blocks.items():
                    original_text = "".join(bdata['text'])
                    display_text = original_text
                    if do_translate:
                        translated = translate_text(original_text, target_lang_code)
                        display_text = translated
                        notes_content.append(f"[原] {original_text}\n[訳] {translated}\n")
                    else:
                        notes_content.append(original_text + "\n")

                    ox = min(bdata['left'])
                    oy = min(bdata['top'])
                    ow = max([l+w for l,w in zip(bdata['left'], bdata['width'])]) - ox
                    orig_x, orig_y, orig_w = int(ox/2), int(oy/2), int(ow/2)

                    if mode.startswith("分解"):
                        cx, cy = orig_x + orig_w/2, orig_y + int(max(bdata['height'])/2)
                        if any(rox < cx < rox+row and roy < cy < roy+roh for (rox,roy,row,roh) in object_rects): continue

                    pp_x = int(orig_x * scale_x)
                    pp_y = int(orig_y * scale_y)
                    pp_w = int(orig_w * scale_x)

                    if pp_w > Inches(0.2):
                        try:
                            is_title = detect_title and (orig_y < h_orig * 0.2)
                            avg_h = (sum(bdata['height'])/len(bdata['height']))/2.0
                            est_pt = (prs.slide_height.inches*72)*(avg_h/h_orig)*0.75
                            font_pt = max(9, min(est_pt, 80))
                            if is_title: font_pt *= 1.2
                            
                            box_h = Inches(font_pt/72 * 1.5 * (display_text.count('\n') + 1))
                            txBox = slide.shapes.add_textbox(pp_x, pp_y, pp_w, box_h)
                            tf = txBox.text_frame
                            tf.word_wrap = True
                            p = tf.paragraphs[0]
                            p.text = display_text
                            p.font.size = Pt(font_pt)
                            p.font.name = target_font
                            p.font.language_id = 1041
                            
                            if is_title:
                                p.font.bold = True
                                p.font.color.rgb = RGBColor(0, 51, 102)
                            else:
                                y1,y2 = max(0,orig_y), min(h_orig,orig_y+int(max(bdata['height'])))
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
                
                slide.notes_slide.notes_text_frame.text = "\n".join(notes_content)
                p_bar.progress((i+1)/process_cnt)
                
                # --- ループ内の修正ここまで ---

        # ★修正4: 最後に status.empty() を呼ばず、成功メッセージで上書きする
        status_area.success("🌌 シンギュラリティ到達！全処理完了。")
        st.balloons()
        
        col_d1, col_d2 = st.columns(2)
        
        out_ppt = io.BytesIO()
        prs.save(out_ppt)
        out_ppt.seek(0)
        col_d1.download_button(f"📥 スライド ({target_lang})", out_ppt, "Singularity_Slides.pptx", type="primary", use_container_width=True)

        if mode.startswith("分解"):
            zip_buffer.seek(0)
            col_d2.download_button(f"🗂️ 画像素材", zip_buffer, "assets.zip", use_container_width=True)
    # === 変更ここまで ===
