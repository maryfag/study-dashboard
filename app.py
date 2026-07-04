import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random

st.set_page_config(page_title="Ultimate Study Dashboard", layout="wide") 

# --- CUSTOM CSS FOR PERFECT ALIGNMENT & LUCIDE DESIGN LAYOUT ---
st.markdown("""
    <style>
        .flex-container {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 10px;
        }
        .icon-svg {
            color: #6D28D9; /* Clean purple profile accent */
            flex-shrink: 0;
            vertical-align: middle;
        }
        .btn-flex {
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        .workspace-card {
            background-color: rgba(255, 255, 255, 0.05);
            padding: 15px;
            border-radius: 8px;
            border: 1px solid rgba(255, 255, 255,
