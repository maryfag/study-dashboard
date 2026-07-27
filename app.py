import base64
import json
import random
import re
from docx import Document
import pypdf
from pptx import Presentation
import requests
import streamlit as st

APP_NAME = "DocDigest"

st.set_page_config(
    page_title=APP_NAME, layout="wide", initial_sidebar_state="expanded"
)

# --- CUSTOM CSS FOR PERFECT ALIGNMENT & LUCIDE DESIGN LAYOUT ---
st.markdown(
    """
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
        .memo-card {
            background-color: rgba(109, 40, 217, 0.05);
            padding: 18px;
            border-radius: 8px;
            border: 1px solid rgba(109, 40, 217, 0.15);
            margin-bottom: 15px;
        }
        section[data-testid="stSidebar"] .stButton > button {
            width: 100%;
            text-align: left;
            border-radius: 6px;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# --- STATE MEMORY CORES: Keeps sections from wiping out ---
if "generated_summary" not in st.session_state:
    st.session_state.generated_summary = None
if "generated_notes" not in st.session_state:
    st.session_state.generated_notes = None
if "generated_cbt" not in st.session_state:
    st.session_state.generated_cbt = None
if "generated_cheatsheet" not in st.session_state:
    st.session_state.generated_cheatsheet = None
if "transcribed_notes" not in st.session_state:
    st.session_state.transcribed_notes = None


# --- TITLE WITH PURE LUCIDE BOOKSTACK ---
st.markdown(
    f"""
    <div class="flex-container">
        <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="m16 6 4 14"/>
            <path d="M12 6v14"/>
            <path d="M8 8v12"/>
            <path d="M4 4v16"/>
        </svg>
        <h1 style="margin: 0; padding: 0; font-size: 2.2rem;">{APP_NAME}</h1>
    </div>
""",
    unsafe_allow_html=True,
)

st.write(
    "Upload your lecture notes, slides, or photos, then choose how you want to study them."
)


# --- HELPER FUNCTIONS ---
def extract_text(uploaded_file):
    filename = uploaded_file.name
    text = ""
    try:
        if filename.endswith(".pdf"):
            pdf_reader = pypdf.PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
        elif filename.endswith(".docx"):
            doc = Document(uploaded_file)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif filename.endswith((".pptx", ".pptm")):
            prs = Presentation(uploaded_file)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text += shape.text.strip() + "\n"
    except Exception as e:
        return None, f"Error reading file structure: {e}"
    return text, None


def ask_gemini(api_key, prompt_text, dynamic_mode=False):
    models = ["gemini-2.0-flash", "gemini-1.5-flash"]
    generation_config = {"temperature": 0.85 if dynamic_mode else 0.2}

    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        headers = {"Content-Type": "application/json"}
        payload = {
            "contents": [{"parts": [{"text": prompt_text}]}],
            "generationConfig": generation_config,
        }
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            response_json = response.json()
            if "candidates" in response_json and response_json["candidates"]:
                return response_json["candidates"][0]["content"]["parts"][0][
                    "text"
                ]
            elif "error" in response_json:
                error_msg = response_json["error"]["message"]
                if any(
                    err in error_msg.lower()
                    for err in ["demand", "quota", "not found"]
                ):
                    continue
                return f"Google API Error: {error_msg}"
        except Exception as e:
            continue

    return "ERROR: Server lines are busy or API Key is invalid. Please try again."


def transcribe_images(api_key, image_files):
    models = ["gemini-2.0-flash", "gemini-1.5-flash"]
    parts = []
    for img_file in image_files:
        img_bytes = img_file.getvalue()
        mime_type = img_file.type or "image/jpeg"
        encoded = base64.b64encode(img_bytes).decode("utf-8")
        parts.append(
            {"inline_data": {"mime_type": mime_type, "data": encoded}}
        )

    parts.append({
        "text": (
            "Transcribe the handwritten and/or printed content of these images "
            "verbatim into plain text, in the exact order given. Preserve structure "
            "(headings, bullet points, numbered lists). Output ONLY the transcribed text."
        )
    })

    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        headers = {"Content-Type": "application/json"}
        payload = {
            "contents": [{"parts": parts}],
            "generationConfig": {"temperature": 0.1},
        }
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=45)
            response_json = response.json()
            if "candidates" in response_json and response_json["candidates"]:
                return response_json["candidates"][0]["content"]["parts"][0][
                    "text"
                ]
        except Exception:
            continue
    return None


def extract_json_block(text):
    """Parses JSON from Gemini output safely."""
    if not text:
        return None
    try:
        # Match array or object JSON pattern
        json_match = re.search(r"(\[.*\]|\{.*\})", text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group(1))
    except Exception:
        pass
    return None


# --- SIDEBAR CONFIGURATION ---
with st.sidebar:
    st.header("⚙️ Configuration")
    api_key = st.text_input("Gemini API Key", type="password")
    st.markdown("---")

    st.subheader("📁 Upload Documents")
    uploaded_files = st.file_uploader(
        "Upload PDF, DOCX, PPTX or Images",
        type=["pdf", "docx", "pptx", "pptm", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )

# --- MAIN APPLICATION WORKFLOW ---
if not api_key:
    st.info("👈 Please enter your Gemini API key in the sidebar to get started.")
    st.stop()

document_text = ""

if uploaded_files:
    text_files = [
        f for f in uploaded_files if not f.name.lower().endswith((".jpg", ".jpeg", ".png"))
    ]
    image_files = [
        f for f in uploaded_files if f.name.lower().endswith((".jpg", ".jpeg", ".png"))
    ]

    # Handle text files
    for tf in text_files:
        extracted, err = extract_text(tf)
        if err:
            st.error(err)
        elif extracted:
            document_text += f"\n--- Content from {tf.name} ---\n" + extracted

    # Handle image files (Transcription)
    if image_files:
        if (
            st.session_state.transcribed_notes is None
            or st.button("Re-transcribe Uploaded Images")
        ):
            with st.spinner("Transcribing text from image files..."):
                st.session_state.transcribed_notes = transcribe_images(
                    api_key, image_files
                )

        if st.session_state.transcribed_notes:
            document_text += (
                "\n--- Transcribed Image Content ---\n"
                + st.session_state.transcribed_notes
            )

# --- STUDY TOOLS DASHBOARD ---
if document_text.strip():
    st.success(
        f"Successfully loaded study materials ({len(document_text.split())} words)."
    )

    tab1, tab2, tab3 = st.tabs(
        ["📝 Core Summary", "🎯 Practice Quiz (CBT)", "⚡ Cheat Sheet"]
    )

    # --- TAB 1: SUMMARY & NOTES ---
    with tab1:
        if st.button("Generate Summary & Study Notes", type="primary"):
            with st.spinner("Analyzing document..."):
                prompt = (
                    "Provide a comprehensive, high-yield summary of these study notes. "
                    "Use bullet points, bold key terms, and format cleanly in Markdown:\n\n"
                    + document_text[:12000]
                )
                res = ask_gemini(api_key, prompt)
                st.session_state.generated_summary = res

        if st.session_state.generated_summary:
            st.markdown(st.session_state.generated_summary)

    # --- TAB 2: CBT PRACTICE QUIZ ---
    with tab2:
        if st.button("Generate 5-Question CBT Quiz"):
            with st.spinner("Building interactive quiz..."):
                prompt = (
                    "Create 5 multiple-choice questions based on the document below. "
                    "Return ONLY a valid JSON array of objects with keys: "
                    "'question', 'options' (array of 4 strings), and 'answer' (index integer 0-3):\n\n"
                    + document_text[:12000]
                )
                raw_quiz = ask_gemini(api_key, prompt)
                parsed_quiz = extract_json_block(raw_quiz)
                st.session_state.generated_cbt = parsed_quiz

        if st.session_state.generated_cbt:
            score = 0
            with st.form("quiz_form"):
                for idx, q in enumerate(st.session_state.generated_cbt):
                    st.write(f"**Q{idx+1}: {q['question']}**")
                    user_choice = st.radio(
                        "Select answer:",
                        q["options"],
                        key=f"q_{idx}",
                        index=None,
                    )

                submitted = st.form_submit_button("Submit Quiz")
                if submitted:
                    st.subheader("Quiz Results")
                    for idx, q in enumerate(st.session_state.generated_cbt):
                        selected = st.session_state.get(f"q_{idx}")
                        correct_idx = q["answer"]
                        correct_txt = q["options"][correct_idx]

                        if selected == correct_txt:
                            st.success(f"Q{idx+1}: Correct!")
                            score += 1
                        else:
                            st.error(
                                f"Q{idx+1}: Incorrect. Correct answer was: **{correct_txt}**"
                            )
                    st.info(
                        f"Final Score: {score} / {len(st.session_state.generated_cbt)}"
                    )

    # --- TAB 3: CHEAT SHEET ---
    with tab3:
        if st.button("Generate High-Yield Cheat Sheet"):
            with st.spinner("Extracting definitions and formulas..."):
                prompt = (
                    "Extract all major definitions, formulas, key dates, and core takeaway points "
                    "into a concise 1-page Cheat Sheet format. Use Markdown tables and bold lists:\n\n"
                    + document_text[:12000]
                )
                st.session_state.generated_cheatsheet = ask_gemini(
                    api_key, prompt
                )

        if st.session_state.generated_cheatsheet:
            st.markdown(st.session_state.generated_cheatsheet)

else:
    st.info("Upload study documents via the sidebar to unlock study features.")
