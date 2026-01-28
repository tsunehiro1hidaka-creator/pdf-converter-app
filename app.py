import streamlit as st
import sys
import os

# === 1. ページ設定 ===
st.set_page_config(page_title="PDF Converter (Debug Mode)", layout="wide")

st.title("🚑 PDF読み込み診断モード")
st.warning("エラー詳細を表示するモードです。")

# === 2. ライブラリ読み込みテスト ===
try:
    import cv2
    import numpy as np
    import pytesseract
    from pdf2image import convert_from_bytes, pdfinfo_from_bytes
    from pptx import Presentation
    from pptx.util import Inches, Pt
    import io
    from deep_translator import GoogleTranslator
    st.success("✅ Pythonライブラリの読み込み成功")
except ImportError as e:
    st.error(f"❌ ライブラリ不足: {e}")
    st.stop()

# === 3. 外部ツール(Poppler)診断 ===
import subprocess
try:
    # pdftoppm (Popplerの一部) があるか確認
    poppler_version = subprocess.check_output(["pdftoppm", "-v"], stderr=subprocess.STDOUT).decode()
    st.success(f"✅ Poppler検出成功: {poppler_version.splitlines()[0]}")
except Exception as e:
    st.error(f"❌ Popplerが見つかりません: {e}")
    st.info("GitHubの 'packages.txt' に 'poppler-utils' が含まれているか確認してください。")
    st.stop()

# === 関数定義 ===
def load_pdf_with_debug(file_bytes):
    try:
        # タイムアウト対策などせず、シンプルに実行
        return convert_from_bytes(file_bytes, dpi=200)
    except Exception as e:
        return str(e)

# === メイン処理 ===
uploaded_file = st.file_uploader("PDFファイルをアップロード（診断用）", type="pdf")

if uploaded_file is not None:
    st.write(f"ファイルサイズ: {uploaded_file.size / 1024:.2f} KB")
    
    if st.button("読み込みテスト開始"):
        file_bytes = uploaded_file.read()
        
        with st.spinner("PDFを解析中..."):
            result = load_pdf_with_debug(file_bytes)
        
        if isinstance(result, list):
            st.success(f"✅ 読み込み成功！ {len(result)} ページ検出しました。")
            st.image(result[0], caption="1ページ目のプレビュー", use_container_width=True)
        else:
            st.error("❌ 読み込み失敗")
            st.code(result, language="text") # エラー内容をそのまま表示
