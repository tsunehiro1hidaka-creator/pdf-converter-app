import streamlit as st
import sys
import os
import io
import datetime

# === 1. ページ設定 (必ず一番最初) ===
st.set_page_config(page_title="Biz PDF Converter Pro", layout="centered")

# ===========================
# テーマ設定関数
# ===========================
def apply_theme(theme):
    base_css = """
        img { border: 1px solid #ddd; border-radius: 5px; }
        .streamlit-expanderHeader { font-weight: bold; font-size: 1.2em; background-color: #f0f2f6; border-radius: 5px; }
        .stMarkdown h3 { border-bottom: 2px solid #ddd; padding-bottom: 5px; margin-top: 20px; }
    """
    
    css = ""
    if theme == "ビジネス (通常)":
        color_primary = "#4CAF50"
        text_color = "#2E7D32"
        css = f"""
            .stButton>button {{ background-color: white; color: {text_color}; border: 2px solid {color_primary}; border-radius: 5px; font-weight: bold; width: 100%; }}
            .stButton>button:hover {{ background-color: {color_primary}; color: white; }}
            h1 {{ color: {text_color}; }}
            .stProgress .st-bo {{ background-color: {color_primary}; }}
        """
    elif theme == "箱推し (全員)":
        css = f"""
            .stApp {{ background: linear-gradient(135deg, #fff0f0 25%, #fffff0 25%, #fffff0 50%, #fff0f5 50%, #fff0f5 75%, #f8f0ff 75%); }}
            .stButton>button {{ 
                background: linear-gradient(90deg, #E60033, #FFF100, #FF69B4, #800080); 
                color: white; border: none; border-radius: 20px; font-weight: bold; width: 100%; text-shadow: 1px 1px 2px black;
            }}
            .stButton>button:hover {{ opacity: 0.9; }}
            h1 {{ 
                background: linear-gradient(90deg, #E60033, #F2C000, #FF69B4, #800080);
