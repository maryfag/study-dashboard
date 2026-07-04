import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random
import json
import re
import io

APP_NAME = "PrepDeck"  # rename here if you want a different app name

st.set_page_config(page_title=APP_NAME, layout="wide")

# ============================================================
# DARK "NOTION / LINEAR" THEME
# ============================================================
st.markdown("""
<style>
    :root {
        --bg: #0E0F13;
        --surface: #17191F;
        --surface-hover: #1D2028;
        --border: rgba(255,255,255,0.08);
        --text: #E6E8EC;
        --muted: #9AA0AC;
        --accent: #818CF8;
        --accent-soft: rgba(129,140,248,0.15);
    }

    .stApp {
        background-color: var(--bg);
        color: var(--text);
    }

    section.main > div {
        padding-top: 1.5rem;
    }

    h1, h2, h3, h4, h5, p, span, label, div {
        color: var(--text);
    }

    .subtitle {
        color: var(--muted);
        font-size: 0.95rem;
        margin-top: -6px;
        margin-bottom: 1.2rem;
    }

    .flex-container {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 4px;
    }
    .icon-svg {
        color: var(--accent);
        flex-shrink: 0;
        vertical-align: middle;
    }

    .workspace-card {
        background-color: var(--surface);
        padding: 16px 18px;
        border-radius: 10px;
        border: 1px solid var(--border);
        margin-bottom: 12px;
    }

    .memo-card {
        background-color: var(--surface);
        padding: 18px 20px;
        border-radius: 10px;
        border: 1px solid var(--border);
        height: 100%;
    }
    .memo-card h4 {
        margin-top: 0;
        color: var(--accent);
        font-size: 0.95rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }

    .palette-wrap {
        background-color: var(--surface);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 14px 16px 6px 16px;
        margin-bottom: 16px;
    }

    /* Streamlit widget overrides */
    .stTextInput > div > div > input,
    .stTextArea textarea {
        background-color: var(--surface-hover) !important;
        color: var(--text) !important;
        border: 1px solid var(--border) !important;
        border-radius: 8px !important;
    }

    .stButton > button {
        background-color: var(--accent-soft);
        color: var(--accent);
        border: 1px solid rgba(129,140,248,0.35);
        border-radius: 8px;
        font-weight: 500;
    }
    .stButton > button:hover {
        background-color: var(--accent);
        color: #0E0F13;
        border: 1px solid var(--accent);
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        border-bottom: 1px solid var(--border);
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent;
        color: var(--muted);
        border-radius: 8px 8px 0 0;
    }
    .stTabs [aria-selected="true"] {
        color: var(--accent) !important;
        border-bottom: 2px solid var(--accent) !important;
    }

    .stRadio label, .stSelectbox label, .stSlider label {
        color: var(--muted) !important;
    }

    hr {
        border-color: var(--border) !important;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# SESSION STATE
# ============================================================
defaults = {
    "generated_notes": None,
    "generated_summary": None,
    "generated_cbt": None,
    "cbt_user_answers": {},
    "cbt_submitted": False,
    "generated_concept_map": None,
    "omni_bar_response": None,
    "last_processed_query": None,
    "chosen_api_key": None,
    "palette_result": None,
    "palette_result_type": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# HEADER
# ============================================================
st.markdown(f"""
    <div class="flex-container">
        <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="m16 6 4 14"/>
            <path d="M12 6v14"/>
            <path d="M8 8v12"/>
            <path d="M4 4v16"/>
        </svg>
        <h1 style="margin: 0; padding: 0; font-size: 2rem;">{APP_NAME}</h1>
    </div>
    <div class="subtitle">Upload a document. Search a command. Study faster.</div>
""", unsafe_allow_html=True)

# ============================================================
# API KEY SETUP (chosen once per session, not re-rolled every rerun)
# ============================================================
api_keys = []
if "GEMINI_API_KEY_1" in st.secrets and st.secrets["GEMINI_API_KEY_1"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_1"])
if "GEMINI_API_KEY_2" in st.secrets and st.secrets["GEMINI_API_KEY_2"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_2"])
if "GEMINI_API_KEY_3" in st.secrets and st.secrets["GEMINI_API_KEY_3"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_3"])

manual_key = st.sidebar.text_input("Backup API Key Entry (Optional)", type="password")

if manual_key:
    api_key = manual_key
else:
    if st.session_state.chosen_api_key is None and api_keys:
        st.session_state.chosen_api_key = random.choice(api_keys)
    api_key = st.session_state.chosen_api_key


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
    "buddy": "Campus Buddy Mode",
    "street": "Street-Smart Analogy Mode",
    "corporate": "Corporate Decoupler Mode",
    "technical": "Deep Technical Mode",
    "layman": "Layman Mode",
}


def build_document_context(chunks):
    return (
        f"[Introductory Section]\n{chunks[0][:3500]}\n\n"
        f"[Core Section A]\n{chunks[1][:3500]}\n\n"
        f"[Core Section B]\n{chunks[2][:3500]}\n\n"
        f"[Advanced / Concluding Section]\n{chunks[3][:3500]}"
    )


def generate_notes_and_analogy(api_key, chunks, mode_key):
    """One API call that returns literal notes + styled analogy as JSON,
    so the two side-by-side memory cards always come from the same source pass."""
    doc_context = build_document_context(chunks)
    style_prompt = STYLE_PROMPTS[mode_key]
    prompt = f"""
    You will produce two things from the document text below, and return them as ONLY a raw JSON object
    (no markdown fences, no commentary), in this exact schema:

    {{
      "literal_notes": "A strict, literal, factual bullet-style summary of the document: exact definitions, configurations, rules, and technical terms. No analogies, no creative language. Use markdown bullet points.",
      "styled_explanation": "The same content re-explained in a different style, described below."
    }}

    Style for "styled_explanation": {style_prompt}

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
# FILE UPLOAD
# ============================================================
uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX, PPTM)", type=["pdf", "docx", "pptx", "pptm"])

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

    # ============================================================
    # COMMAND PALETTE — type to filter preset macros
    # ============================================================
    st.markdown('<div class="palette-wrap">', unsafe_allow_html=True)
    st.markdown("**⌘ Command Palette** — type to filter actions (e.g. 'quiz', 'map', 'buddy', 'technical')")

    palette_query = st.text_input(
        "command_palette",
        placeholder="Search a command...",
        label_visibility="collapsed",
        key="palette_input"
    )

    PRESETS = [
        {"id": "quiz", "label": "📝 Quiz me (CBT practice)", "keywords": ["quiz", "cbt", "test", "practice", "question"]},
        {"id": "map", "label": "🧠 Build a concept map", "keywords": ["map", "concept", "relationship"]},
        {"id": "buddy", "label": "🎓 Explain it — Campus Buddy", "keywords": ["buddy", "campus", "student"]},
        {"id": "street", "label": "🏙️ Explain it — Street-Smart", "keywords": ["street", "practical", "real world"]},
        {"id": "corporate", "label": "🏢 Explain it — Corporate Decoupler", "keywords": ["corporate", "jargon", "business"]},
        {"id": "technical", "label": "🔬 Explain it — Deep Technical", "keywords": ["technical", "deep", "expert", "rigor"]},
        {"id": "layman", "label": "🍼 Explain it — Layman / ELI5", "keywords": ["layman", "eli5", "simple", "beginner"]},
    ]

    q = palette_query.strip().lower()
    if q:
        matches = [p for p in PRESETS if q in p["label"].lower() or any(q in kw for kw in p["keywords"])]
    else:
        matches = PRESETS

    cols_per_row = 4
    for row_start in range(0, len(matches), cols_per_row):
        row_items = matches[row_start:row_start + cols_per_row]
        cols = st.columns(cols_per_row)
        for col, preset in zip(cols, row_items):
            with col:
                if st.button(preset["label"], key=f"preset_{preset['id']}"):
                    if not api_key:
                        st.error("Missing API Key!")
                    elif not raw_text.strip():
                        st.error("Could not extract any text from this document.")
                    else:
                        with st.spinner("Running command..."):
                            if preset["id"] == "quiz":
                                parsed = generate_quiz(api_key, chunks, 5)
                                if parsed:
                                    st.session_state.generated_cbt = parsed
                                    st.session_state.cbt_user_answers = {}
                                    st.session_state.cbt_submitted = False
                                    st.session_state.palette_result = parsed
                                    st.session_state.palette_result_type = "quiz"
                                else:
                                    st.error("Couldn't generate questions — try again.")
                            elif preset["id"] == "map":
                                parsed = generate_concept_map(api_key, chunks)
                                if parsed:
                                    st.session_state.generated_concept_map = parsed
                                    st.session_state.palette_result = parsed
                                    st.session_state.palette_result_type = "map"
                                else:
                                    st.error("Couldn't generate a concept map — try again.")
                            else:
                                parsed = generate_notes_and_analogy(api_key, chunks, preset["id"])
                                if parsed:
                                    st.session_state.generated_notes = parsed.get("literal_notes")
                                    st.session_state.generated_summary = parsed.get("styled_explanation")
                                    st.session_state.palette_result = parsed
                                    st.session_state.palette_result_type = "notes"
                                else:
                                    st.error("Couldn't generate a summary — try again.")

    if not matches:
        st.caption("No commands match that search.")

    st.markdown('</div>', unsafe_allow_html=True)

    # Quick inline preview right under the palette so results are visible without switching tabs
    if st.session_state.palette_result_type == "notes" and st.session_state.generated_notes:
        pc1, pc2 = st.columns(2)
        with pc1:
            st.markdown(f'<div class="memo-card"><h4>Literal Notes</h4>{st.session_state.generated_notes}</div>', unsafe_allow_html=True)
        with pc2:
            st.markdown(f'<div class="memo-card"><h4>Analogy</h4>{st.session_state.generated_summary}</div>', unsafe_allow_html=True)
    elif st.session_state.palette_result_type == "map" and st.session_state.generated_concept_map:
        st.caption("Concept map generated — see the 'Concept Map Table' tab below.")
    elif st.session_state.palette_result_type == "quiz" and st.session_state.generated_cbt:
        st.caption("Quiz generated — see the 'CBT Objective Practice' tab below.")

    # ============================================================
    # FREE-FORM OMNI-BAR (ask anything, not a preset)
    # ============================================================
    with st.expander("💬 Ask anything (free-form question about this document)"):
        shortcut_input = st.text_input(
            "Ask Gemini anything about this document...",
            placeholder="e.g. 'Extract the major risks', 'Give me a list of terms'",
            key="omni_input"
        )
        if shortcut_input:
            cleaned_query = shortcut_input.strip()
            if cleaned_query != st.session_state.last_processed_query:
                if not api_key:
                    st.error("Missing API Key!")
                else:
                    with st.spinner("Gemini is analyzing the document..."):
                        context_sample = f"[Document Excerpt]\n{chunks[0][:5000]}\n{chunks[1][:3000]}"
                        omni_prompt = f"""
                        You are a real-time smart search engine assistant for this document.
                        Answer the following custom user query based strictly on the document context provided below.

                        User Query: {cleaned_query}

                        Document Context:
                        {context_sample}
                        """
                        result = ask_gemini(api_key, omni_prompt, dynamic_mode=True)
                        if result is None:
                            st.error("Server lines are busy. Try again in a moment!")
                        else:
                            st.session_state.omni_bar_response = result
                            st.session_state.last_processed_query = cleaned_query

        if st.session_state.omni_bar_response:
            st.markdown('<div class="workspace-card">', unsafe_allow_html=True)
            st.markdown(st.session_state.omni_bar_response)
            st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ============================================================
    # WORKSPACE TABS
    # ============================================================
    tab1, tab2, tab3 = st.tabs([
        "Custom Explanation Summary",
        "CBT Objective Practice",
        "Concept Map Table"
    ])

    # --- TAB 1: side-by-side memory cards ---
    with tab1:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
                <h3 style="margin:0; font-size: 1.15rem;">Tailored Multi-Mode Explanation Engine</h3>
            </div>
        """, unsafe_allow_html=True)

        mode_choice = st.selectbox(
            "Choose Your Desired Explanation Persona:",
            list(MODE_LABELS.values())
        )
        mode_key = [k for k, v in MODE_LABELS.items() if v == mode_choice][0]

        if st.button("Generate Notes & Analogy (side by side)", key="btn_summary"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Building your literal notes and analogy..."):
                    parsed = generate_notes_and_analogy(api_key, chunks, mode_key)
                    if parsed:
                        st.session_state.generated_notes = parsed.get("literal_notes")
                        st.session_state.generated_summary = parsed.get("styled_explanation")
                    else:
                        st.error("Couldn't generate a summary this time — please try again.")

        if st.session_state.generated_notes or st.session_state.generated_summary:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(
                    f'<div class="memo-card"><h4>📌 Literal Notes</h4>{st.session_state.generated_notes or "—"}</div>',
                    unsafe_allow_html=True
                )
            with c2:
                st.markdown(
                    f'<div class="memo-card"><h4>💡 Analogy</h4>{st.session_state.generated_summary or "—"}</div>',
                    unsafe_allow_html=True
                )

    # --- TAB 2: CBT quiz ---
    with tab2:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
                <h3 style="margin:0; font-size: 1.15rem;">CBT-Style Objective Practice</h3>
            </div>
        """, unsafe_allow_html=True)

        num_questions = st.slider("How many questions?", min_value=3, max_value=15, value=5, key="num_q")

        if st.button("Generate Practice Questions", key="btn_cbt"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Writing your CBT practice questions..."):
                    parsed = generate_quiz(api_key, chunks, num_questions)
                    if not parsed:
                        st.error("Couldn't generate valid questions this time — please try again.")
                    else:
                        st.session_state.generated_cbt = parsed
                        st.session_state.cbt_user_answers = {}
                        st.session_state.cbt_submitted = False

        if st.session_state.generated_cbt:
            st.markdown('<div class="workspace-card"><h4>📝 Practice Quiz</h4></div>', unsafe_allow_html=True)
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

    # --- TAB 3: concept map ---
    with tab3:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="5" r="2"/><circle cx="5" cy="19" r="2"/><circle cx="19" cy="19" r="2"/><path d="M12 7v6M12 13l-6 4M12 13l6 4"/></svg>
                <h3 style="margin:0; font-size: 1.15rem;">Concept Map Table</h3>
            </div>
        """, unsafe_allow_html=True)

        if st.button("Generate Concept Map", key="btn_concept_map"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Mapping out the key concepts and how they connect..."):
                    parsed = generate_concept_map(api_key, chunks)
                    if not parsed:
                        st.error("Couldn't generate a concept map this time — please try again.")
                    else:
                        st.session_state.generated_concept_map = parsed

        if st.session_state.generated_concept_map:
            st.markdown('<div class="workspace-card"><h4>🧠 Key Concepts</h4></div>', unsafe_allow_html=True)
            st.table(st.session_state.generated_concept_map)

else:
    st.info("Upload a document above to unlock the command palette and all three workspace tabs.")
