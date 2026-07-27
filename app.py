import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random
import json
import re
import base64

APP_NAME = "DocDigest"

st.set_page_config(page_title=APP_NAME, layout="wide", initial_sidebar_state="expanded")

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
        .memo-card {
            background-color: rgba(109, 40, 217, 0.05);
            padding: 18px;
            border-radius: 8px;
            border: 1px solid rgba(109, 40, 217, 0.15);
            margin-bottom: 15px;
            height: 100%;
        }
        section[data-testid="stSidebar"] .stButton > button {
            width: 100%;
            text-align: left;
            border-radius: 6px;
        }
        /* Lock the sidebar open — hide the collapse arrow entirely */
        [data-testid="stSidebarCollapseButton"] {
            display: none !important;
        }
        [data-testid="collapsedControl"] {
            display: none !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- STATE MEMORY CORES: Keeps sections from wiping out ---
if "generated_summary" not in st.session_state:
    st.session_state.generated_summary = None
if "generated_notes" not in st.session_state:
    st.session_state.generated_notes = None
if "generated_sections" not in st.session_state:
    st.session_state.generated_sections = None
if "generated_mode_label" not in st.session_state:
    st.session_state.generated_mode_label = None
if "generated_batch_label" not in st.session_state:
    st.session_state.generated_batch_label = None
if "generated_cbt" not in st.session_state:
    st.session_state.generated_cbt = None
if "current_cbt_batch" not in st.session_state:
    st.session_state.current_cbt_batch = None
if "generated_cheatsheet" not in st.session_state:
    st.session_state.generated_cheatsheet = None
if "active_view" not in st.session_state:
    st.session_state.active_view = None  # "analogy" | "quiz" | "cheatsheet"
if "transcribed_notes" not in st.session_state:
    st.session_state.transcribed_notes = None

# --- TITLE WITH PURE LUCIDE BOOKSTACK ---
st.markdown(f"""
    <div class="flex-container">
        <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="m16 6 4 14"/>
            <path d="M12 6v14"/>
            <path d="M8 8v12"/>
            <path d="M4 4v16"/>
        </svg>
        <h1 style="margin: 0; padding: 0; font-size: 2.2rem;">{APP_NAME}</h1>
    </div>
""", unsafe_allow_html=True)

st.write("Upload your lecture notes, slides, or PDFs, then choose how you want to conquer them.")


def extract_text(uploaded_file):
    filename = uploaded_file.name
    text = ""
    try:
        if filename.endswith('.pdf'):
            pdf_reader = pypdf.PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                try:
                    t = page.extract_text()
                    if t: text += t + "\n"
                except: continue
        elif filename.endswith('.docx'):
            doc = Document(uploaded_file)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif filename.endswith('.pptx') or filename.endswith('.pptm'):
            prs = Presentation(uploaded_file)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text += shape.text.strip() + "\n"
    except Exception as e:
        return f"Error reading file structure: {e}"
    return text


def ask_gemini(api_key, prompt_text, dynamic_mode=False):
    models = ["gemini-2.5-flash", "gemini-1.5-flash"]
    generation_config = {"temperature": 0.85 if dynamic_mode else 0.2}

    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}
        payload = {
            "contents": [{"parts": [{"text": prompt_text}]}],
            "generationConfig": generation_config
        }
        try:
            response = requests.post(url, headers=headers, json=payload)
            response_json = response.json()
            if 'candidates' in response_json and response_json['candidates']:
                return response_json['candidates'][0]['content']['parts'][0]['text']
            elif 'error' in response_json:
                error_msg = response_json['error']['message']
                if "demand" in error_msg.lower() or "quota" in error_msg.lower() or "not found" in error_msg.lower():
                    continue
                return f"Google API Error: {error_msg}"
        except: continue

    # --- PURE LUCIDE ALERT BOX ---
    return """
    <div style="display:flex; align-items:center; gap:8px; color:#DC2626; font-weight:600;">
        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
        Server lines are busy. Tap the button again!
    </div>
    """


def transcribe_images(api_key, image_files):
    """Sends photos directly to Gemini's multimodal endpoint and asks it to
    transcribe the handwritten/printed content into plain text."""
    models = ["gemini-2.5-flash", "gemini-1.5-flash"]

    parts = []
    for img_file in image_files:
        img_bytes = img_file.getvalue()
        mime_type = img_file.type or "image/jpeg"
        encoded = base64.b64encode(img_bytes).decode("utf-8")
        parts.append({"inline_data": {"mime_type": mime_type, "data": encoded}})

    parts.append({
        "text": (
            "Transcribe the handwritten and/or printed content of these images "
            "verbatim into plain text, in the exact order the images were given "
            "(treat them as consecutive pages). Preserve structure where visible "
            "(headings, bullet points, numbered lists). Do not summarize, "
            "correct, or add commentary — output ONLY the transcribed text."
        )
    })

    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}
        payload = {
            "contents": [{"parts": parts}],
            "generationConfig": {"temperature": 0.1}
        }
        try:
            response = requests.post(url, headers=headers, json=payload)
            response_json = response.json()
            if 'candidates' in response_json and response_json['candidates']:
                return response_json['candidates'][0]['content']['parts'][0]['text']
            elif 'error' in response_json:
                error_msg = response_json['error']['message']
                if "demand" in error_msg.lower() or "quota" in error_msg.lower() or "not found" in error_msg.lower():
                    continue
                return None
        except Exception:
            continue
    return None


def extract_json_block(text):
    """Parses JSON from Gemini output cleanly without string syntax bugs."""
    if not text:
        return None
    cleaned = text.replace("```json", "").replace("```", "").strip()
