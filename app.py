import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random
import json
import re
import io

APP_NAME = "DocDigest"

st.set_page_config(page_title=APP_NAME, layout="wide")

# ============================================================
# LIGHT THEME, PURPLE ACCENT (kept close to the original look)
# ============================================================
st.markdown("""
<style>
    .flex-container {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 10px;
    }
    .icon-svg {
        color: #6D28D9;
        flex-shrink: 0;
        vertical-align: middle;
    }
    .workspace-card {
        background-color: rgba(109, 40, 217, 0.05);
        padding: 18px 20px;
        border-radius: 8px;
        border: 1px solid rgba(109, 40, 217, 0.15);
        margin-bottom: 12px;
    }
    .memo-card {
        background-color: rgba(109, 40, 217, 0.05);
        padding: 18px 20px;
        border-radius: 8px;
        border: 1px solid rgba(109, 40, 217, 0.15);
        height: 100%;
    }
    .memo-card h4 {
        margin-top: 0;
        color: #6D28D9;
    }
    .subtitle {
        color: #6b7280;
        margin-top: -8px;
        margin-bottom: 1.2rem;
    }
    section[data-testid="stSidebar"] .stButton > button {
        width: 100%;
        text-align: left;
        border-radius: 6px;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# SESSION STATE
# ============================================================
defaults = {
    "generated_sections": None,
    "generated_mode_label": None,
    "generated_cbt": None,
    "cbt_user_answers": {},
    "cbt_submitted": False,
    "generated_concept_map": None,
    "chosen_api_key": None,
    "active_view": None,  # "analogy" | "quiz" | "map"
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# CORE HELPERS
# ============================================================
@st.cache_data(show_spinner=False)
def extract_text(file_bytes, filename):
    text = ""
    try:
        if filename.endswith('.pdf'):
            pdf_reader = pypdf.PdfReader(io.BytesIO(file_bytes))
            for page in pdf_reader.pages:
                try:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
                except Exception:
                    continue
        elif filename.endswith('.docx'):
            doc = Document(io.BytesIO(file_bytes))
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif filename.endswith('.pptx') or filename.endswith('.pptm'):
            prs = Presentation(io.BytesIO(file_bytes))
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
            response = requests.post(url, headers=headers, json=payload, timeout=60)
            response_json = response.json()
            if 'candidates' in response_json and response_json['candidates']:
                return response_json['candidates'][0]['content']['parts'][0]['text']
            elif 'error' in response_json:
                continue
        except Exception:
            continue
    return None


def extract_json_block(text):
    if not text:
        return None
    cleaned = re.sub(r"```json|```", "", text).strip()
    match = re.search(r"(\[.*\]|\{.*\})", cleaned, re.DOTALL)
    if not match:
        return None
    try:
        return json.loads(match.group(1))
    except Exception:
        return None


STYLE_PROMPTS = {
    "buddy": "You are a relatable university peer tutor. Use simple, engaging, and funny student/campus analogies (like hostel porters or campus gates). Highlight key terms in **bold**.",
    "street": "You are a highly practical, street-smart operations mentor. Explain using crisp, real-world analogies based on everyday logic, physical logistics, or coordinating delivery logistics. Avoid corporate jargon and campus-specific terms. Highlight key terms in **bold**.",
    "corporate": "You are a clean communicator stripping away complex corporate jargon and buzzwords. Re-explain the practical mechanism using general real-world metaphors. Highlight key terms in **bold**.",
    "technical": "You are a rigorous subject-matter expert writing for advanced students. Preserve precise terminology and explain underlying mechanisms in full technical depth, without folksy analogies. Highlight key terms in **bold**.",
    "layman": "You are explaining this to a complete beginner, like explaining to a curious 5-year-old. Use short sentences, simple words, and everyday comparisons. Highlight key terms in **bold**.",
}

MODE_LABELS = {
    "buddy": "Campus Buddy",
    "street": "Street-Smart",
    "corporate": "Corporate Decoupler",
    "technical": "Deep Technical",
    "layman": "Layman / ELI5",
}


def build_document_context(chunks):
    return (
        f"[Introductory Section]\n{chunks[0][:3500]}\n\n"
        f"[Core Section A]\n{chunks[1][:3500]}\n\n"
        f"[Core Section B]\n{chunks[2][:3500]}\n\n"
        f"[Advanced / Concluding Section]\n{chunks[3][:3500]}"
    )


def generate_notes_and_analogy(api_key, chunks, mode_key):
    """Breaks the document into topic-sized sections and returns, for each section,
    the literal note AND its matching analogy — so they can be rendered as paired
    rows instead of two unrelated blocks of text."""
    doc_context = build_document_context(chunks)
    style_prompt = STYLE_PROMPTS[mode_key]
    prompt = f"""
    Break the document below into 5-9 logical topic sections (definitions, rules,
    processes, configurations — whatever the natural divisions are).

    For EACH section, produce a matched pair: the literal factual note, and the
    analogy explanation of that SAME piece of content. They must correspond to
    each other one-to-one — do not summarize the whole document twice separately.

    Respond with ONLY a raw JSON array, no markdown fences, no commentary, in this
    exact schema:
    [
      {{
        "topic": "short section title, e.g. 'Definition of X' or 'Rule for Y'",
        "literal_note": "Strict, literal, factual explanation of this section only: exact definitions, configurations, or rules. No analogies. 1-3 sentences or a short bullet list.",
        "analogy": "The SAME section re-explained using the style below. 1-3 sentences."
      }}
    ]

    Style for "analogy": {style_prompt}

    Document Text:
    {doc_context}
    """
    result = ask_gemini(api_key, prompt, dynamic_mode=True)
    return extract_json_block(result)


def generate_quiz(api_key, chunks, num_questions):
    doc_context = build_document_context(chunks)
    prompt = f"""
    Based strictly on the document text below, write exactly {num_questions} multiple-choice
    objective questions suitable for exam practice.

    Respond with ONLY a raw JSON array, no markdown fences, no commentary, in this exact schema:
    [
      {{
        "question": "string",
        "options": ["string", "string", "string", "string"],
        "correct_index": 0,
        "explanation": "string"
      }}
    ]

    Document Text:
    {doc_context}
    """
    result = ask_gemini(api_key, prompt, dynamic_mode=False)
    return extract_json_block(result)


def generate_concept_map(api_key, chunks):
    doc_context = build_document_context(chunks)
    prompt = f"""
    Based strictly on the document text below, identify the key concepts a student should know.

    Respond with ONLY a raw JSON array, no markdown fences, no commentary, in this exact schema:
    [
      {{
        "concept": "string",
        "definition": "string (1-2 sentences)",
        "related_to": "string (name of another concept in this list it connects to, or 'None')"
      }}
    ]

    Aim for 6-12 concepts.

    Document Text:
    {doc_context}
    """
    result = ask_gemini(api_key, prompt, dynamic_mode=False)
    return extract_json_block(result)


# ============================================================
# LEFT SIDEBAR — everything is an instruction/control, clustered together
# ============================================================
with st.sidebar:
    st.markdown(f"""
        <div class="flex-container">
            <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                <path d="m16 6 4 14"/><path d="M12 6v14"/><path d="M8 8v12"/><path d="M4 4v16"/>
            </svg>
            <h2 style="margin:0; font-size:1.4rem;">{APP_NAME}</h2>
        </div>
    """, unsafe_allow_html=True)
    st.caption("Upload a document, then run an action below.")

    uploaded_file = st.file_uploader("Upload document", type=["pdf", "docx", "pptx", "pptm"])

    # API key setup (chosen once per session)
    api_keys = []
    if "GEMINI_API_KEY_1" in st.secrets and st.secrets["GEMINI_API_KEY_1"]:
        api_keys.append(st.secrets["GEMINI_API_KEY_1"])
    if "GEMINI_API_KEY_2" in st.secrets and st.secrets["GEMINI_API_KEY_2"]:
        api_keys.append(st.secrets["GEMINI_API_KEY_2"])
    if "GEMINI_API_KEY_3" in st.secrets and st.secrets["GEMINI_API_KEY_3"]:
        api_keys.append(st.secrets["GEMINI_API_KEY_3"])

    manual_key = st.text_input("Backup API Key (optional)", type="password")
    if manual_key:
        api_key = manual_key
    else:
        if st.session_state.chosen_api_key is None and api_keys:
            st.session_state.chosen_api_key = random.choice(api_keys)
        api_key = st.session_state.chosen_api_key

    raw_text = ""
    chunks = ["", "", "", ""]

    if uploaded_file:
        file_bytes = uploaded_file.getvalue()
        raw_text = extract_text(file_bytes, uploaded_file.name)
        total_length = len(raw_text) if raw_text else 0
        chunk_size = max(1, total_length // 4)
        chunks = [
            raw_text[0:chunk_size] if total_length > 0 else "",
            raw_text[chunk_size:chunk_size * 2] if total_length > 0 else "",
            raw_text[chunk_size * 2:chunk_size * 3] if total_length > 0 else "",
            raw_text[chunk_size * 3:] if total_length > 0 else "",
        ]

        st.markdown("---")
        st.markdown("**Quiz & Concepts**")

        if st.button("📝  Quiz me"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Writing quiz questions..."):
                    parsed = generate_quiz(api_key, chunks, 5)
                    if parsed:
                        st.session_state.generated_cbt = parsed
                        st.session_state.cbt_user_answers = {}
                        st.session_state.cbt_submitted = False
                        st.session_state.active_view = "quiz"
                    else:
                        st.error("Couldn't generate questions — try again.")

        if st.button("🧠  Concept map"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Mapping key concepts..."):
                    parsed = generate_concept_map(api_key, chunks)
                    if parsed:
                        st.session_state.generated_concept_map = parsed
                        st.session_state.active_view = "map"
                    else:
                        st.error("Couldn't generate a concept map — try again.")

        st.markdown("---")
        st.markdown("**Analogy** — pick a style")

        # A row of analogy-style options, all visible together (no dropdown)
        mode_row_1 = st.columns(2)
        mode_row_2 = st.columns(2)
        mode_row_3 = st.columns(1)
        mode_buttons = [
            (mode_row_1[0], "buddy"),
            (mode_row_1[1], "street"),
            (mode_row_2[0], "corporate"),
            (mode_row_2[1], "technical"),
            (mode_row_3[0], "layman"),
        ]
        for col, mode_key in mode_buttons:
            with col:
                if st.button(MODE_LABELS[mode_key], key=f"mode_{mode_key}"):
                    if not api_key:
                        st.error("Missing API Key!")
                    elif not raw_text.strip():
                        st.error("Could not extract any text from this document.")
                    else:
                        with st.spinner("Building notes and analogy..."):
                            parsed = generate_notes_and_analogy(api_key, chunks, mode_key)
                            if parsed:
                                st.session_state.generated_sections = parsed
                                st.session_state.generated_mode_label = MODE_LABELS[mode_key]
                                st.session_state.active_view = "analogy"
                            else:
                                st.error("Couldn't generate a summary — try again.")

# ============================================================
# MAIN AREA — pure presentation, nothing else
# ============================================================
if not uploaded_file:
    st.markdown(f"## {APP_NAME}")
    st.markdown('<div class="subtitle">Upload a document on the left to get started.</div>', unsafe_allow_html=True)
else:
    view = st.session_state.active_view

    if view is None:
        st.markdown(f"## {APP_NAME}")
        st.markdown(
            '<div class="subtitle">Document loaded. Choose an action on the left — Quiz me, Concept map, or an Analogy style.</div>',
            unsafe_allow_html=True
        )

    elif view == "analogy":
        st.markdown(f"## 💡 {st.session_state.generated_mode_label} Analogy")
        sections = st.session_state.generated_sections or []

        # Column headers, shown once
        h1, h2 = st.columns(2)
        with h1:
            st.markdown("**📌 Literal Notes**")
        with h2:
            st.markdown("**💡 Analogy**")

        for sec in sections:
            topic = sec.get("topic", "").strip()
            note = sec.get("literal_note", "").strip()
            analogy = sec.get("analogy", "").strip()

            if topic:
                st.markdown(f"#### {topic}")

            c1, c2 = st.columns(2)
            with c1:
                st.markdown(
                    f'<div class="memo-card">{note or "—"}</div>',
                    unsafe_allow_html=True
                )
            with c2:
                st.markdown(
                    f'<div class="memo-card">{analogy or "—"}</div>',
                    unsafe_allow_html=True
                )
            st.markdown("")  # small spacer between sections

    elif view == "quiz":
        st.markdown("## 📝 Practice Quiz")
        with st.form("cbt_quiz_form"):
            for i, qq in enumerate(st.session_state.generated_cbt):
                st.markdown(f"**Q{i + 1}. {qq.get('question', '')}**")
                choice = st.radio(
                    label=f"question_{i}",
                    options=list(range(len(qq.get("options", [])))),
                    format_func=lambda idx, opts=qq.get("options", []): opts[idx],
                    key=f"cbt_radio_{i}",
                    label_visibility="collapsed"
                )
                st.session_state.cbt_user_answers[i] = choice
            submitted = st.form_submit_button("Submit Answers")
            if submitted:
                st.session_state.cbt_submitted = True

        if st.session_state.cbt_submitted:
            score = 0
            for i, qq in enumerate(st.session_state.generated_cbt):
                user_choice = st.session_state.cbt_user_answers.get(i)
                correct_index = qq.get("correct_index")
                if user_choice == correct_index:
                    score += 1
                    st.success(f"Q{i + 1}: Correct!")
                else:
                    options = qq.get("options", [])
                    correct_text = options[correct_index] if correct_index is not None and correct_index < len(options) else "N/A"
                    st.error(f"Q{i + 1}: Incorrect. Correct answer: {correct_text}")
                if qq.get("explanation"):
                    st.caption(qq["explanation"])
            st.markdown(f"### Score: {score} / {len(st.session_state.generated_cbt)}")

    elif view == "map":
        st.markdown("## 🧠 Concept Map")
        st.markdown('<div class="workspace-card">', unsafe_allow_html=True)
        st.table(st.session_state.generated_concept_map)
        st.markdown('</div>', unsafe_allow_html=True)
