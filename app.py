import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random
import json
import re

st.set_page_config(page_title="Ultimate Study Dashboard", layout="wide")

# --- CUSTOM CSS FOR PERFECT ALIGNMENT ---
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
        background-color: rgba(255, 255, 255, 0.05);
        padding: 15px;
        border-radius: 8px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# --- STATE MEMORY CORES: Total protection against blank screen re-runs ---
if "generated_summary" not in st.session_state:
    st.session_state.generated_summary = None
if "generated_cbt" not in st.session_state:
    st.session_state.generated_cbt = None
if "cbt_user_answers" not in st.session_state:
    st.session_state.cbt_user_answers = {}
if "cbt_submitted" not in st.session_state:
    st.session_state.cbt_submitted = False
if "generated_concept_map" not in st.session_state:
    st.session_state.generated_concept_map = None
if "omni_bar_response" not in st.session_state:
    st.session_state.omni_bar_response = None
if "last_processed_query" not in st.session_state:
    st.session_state.last_processed_query = None
if "chosen_api_key" not in st.session_state:
    st.session_state.chosen_api_key = None

# --- TITLE WITH PURE LUCIDE BOOKSTACK ---
st.markdown("""
    <div class="flex-container">
        <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="m16 6 4 14"/>
            <path d="M12 6v14"/>
            <path d="M8 8v12"/>
            <path d="M4 4v16"/>
        </svg>
        <h1 style="margin: 0; padding: 0; font-size: 2.2rem;">Your Ultimate Exam Survival Dashboard</h1>
    </div>
""", unsafe_allow_html=True)

st.write("Upload your lecture notes, slides, or PDFs, then choose how you want to conquer them.")

# --- SMART API KEY ROTATION SETUP (fixed: chosen once per session, not re-rolled every rerun) ---
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


# --- TEXT EXTRACTION (cached so it doesn't re-run the PDF/DOCX/PPTX parser on every click) ---
@st.cache_data(show_spinner=False)
def extract_text(file_bytes, filename):
    import io
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

    return None  # None = real failure, used to decide whether to cache/store the result


def extract_json_block(text):
    """Pulls a JSON array/object out of a model response, stripping markdown fences."""
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


uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX, PPTM)", type=["pdf", "docx", "pptx", "pptm"])

if uploaded_file:
    file_bytes = uploaded_file.getvalue()
    raw_text = extract_text(file_bytes, uploaded_file.name)
    total_length = len(raw_text) if raw_text else 0
    chunk_size = max(1, total_length // 4)

    chunk_1 = raw_text[0:chunk_size] if total_length > 0 else ""
    chunk_2 = raw_text[chunk_size:chunk_size * 2] if total_length > 0 else ""
    chunk_3 = raw_text[chunk_size * 2:chunk_size * 3] if total_length > 0 else ""
    chunk_4 = raw_text[chunk_size * 3:] if total_length > 0 else ""

    # --- TRUE AI OMNI-BAR COMMAND PALETTE ---
    st.markdown("### ⚡ AI Omni-Bar Command Palette")
    shortcut_input = st.text_input(
        "Ask Gemini anything about this document (e.g., 'Extract the major risks', 'Give me a list of terms')...",
        placeholder="Type your dynamic query and press Enter..."
    )

    if shortcut_input:
        cleaned_query = shortcut_input.strip()
        if cleaned_query != st.session_state.last_processed_query:
            if not api_key:
                st.error("Missing API Key!")
            else:
                with st.spinner("Gemini is searching and analyzing the document layout..."):
                    context_sample = f"[Document Excerpt Section]\n{chunk_1[:5000]}\n{chunk_2[:3000]}"
                    omni_prompt = f"""
                    You are a real-time smart search engine assistant for this document.
                    Answer the following custom user query based strictly on the document context provided below.

                    User Query: {cleaned_query}

                    Document Context:
                    {context_sample}
                    """
                    result = ask_gemini(api_key, omni_prompt, dynamic_mode=True)
                    if result is None:
                        st.error("Server lines are busy. Tap Enter again in a moment!")
                        # Note: last_processed_query is intentionally NOT updated here,
                        # so retrying the same query will actually retry the request.
                    else:
                        st.session_state.omni_bar_response = result
                        st.session_state.last_processed_query = cleaned_query

    if st.session_state.omni_bar_response:
        st.markdown('<div class="workspace-card"><h4>🔍 AI Omni-Bar Search Result</h4></div>', unsafe_allow_html=True)
        st.markdown(st.session_state.omni_bar_response)

    st.markdown("---")

    # --- WORKSPACE TABS ---
    tab1, tab2, tab3 = st.tabs([
        "Custom Explanation Summary",
        "CBT Objective Practice",
        "Concept Map Table"
    ])

    # ==========================================================
    # TAB 1 — Tailored Multi-Mode Explanation Engine
    # ==========================================================
    with tab1:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
                <h3 style="margin:0;">Tailored Multi-Mode Explanation Engine</h3>
            </div>
        """, unsafe_allow_html=True)

        explanation_mode = st.selectbox(
            "Choose Your Desired Explanation Persona:",
            ["Campus Buddy Mode (Student & Campus Analogies)",
             "Street-Smart Analogy Mode (Practical, Everyday Logic & Logistics)",
             "Corporate Decoupler Mode (Stripping Out Complex Enterprise Buzzwords)",
             "Deep Technical Mode (Upper-Level Technical Rigor)",
             "Layman Mode (Explain Like I'm 5 Style)"]
        )

        if st.button("Generate Tailored Summary", key="btn_summary"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Processing custom persona analytics across the full document..."):
                    safe_combined_text = (
                        f"[Introductory Section]\n{chunk_1[:3500]}\n\n"
                        f"[Core Section A]\n{chunk_2[:3500]}\n\n"
                        f"[Core Section B]\n{chunk_3[:3500]}\n\n"
                        f"[Advanced / Concluding Section]\n{chunk_4[:3500]}"
                    )

                    if "Campus Buddy" in explanation_mode:
                        style_prompt = "You are a relatable university peer tutor. Use simple, engaging, and funny student/campus analogies (like hostel porters or campus gates) to explain everything simply. Highlight key terms in **bold**."
                    elif "Street-Smart" in explanation_mode:
                        style_prompt = "You are a highly practical, street-smart operations mentor. Explain the concepts using crisp, real-world analogies based on everyday logic, physical logistics, spotting counterfeits, managing daily physical operations, or coordinating delivery logistics. Avoid corporate boardroom slang and avoid campus-specific university terms. Make it punchy and clear for anyone living in the real world. Highlight key terms in **bold**."
                    elif "Corporate Decoupler" in explanation_mode:
                        style_prompt = "You are a clean communicator stripping away complex corporate jargon, buzzwords, or intense contractual terminology. Re-explain the document's practical mechanism cleanly using general real-world metaphors. Highlight key terms in **bold**."
                    elif "Deep Technical" in explanation_mode:
                        style_prompt = "You are a rigorous subject-matter expert writing for advanced students. Preserve precise terminology, cite exact definitions and configurations from the source text, and explain underlying mechanisms in full technical depth. Do not simplify with folksy analogies. Highlight key terms in **bold**."
                    else:  # Layman Mode
                        style_prompt = "You are explaining this to a complete beginner with no background knowledge, like explaining to a curious 5-year-old. Use short sentences, simple words, and everyday comparisons. Highlight key terms in **bold**."

                    final_prompt = f"""
                    {style_prompt}

                    Base your explanation strictly on the factual content, technical configurations, definitions, and rules
                    contained in the document text below. Do not invent facts that are not present in the text.

                    Document Text:
                    {safe_combined_text}
                    """

                    result = ask_gemini(api_key, final_prompt, dynamic_mode=True)
                    if result is None:
                        st.error("Server lines are busy. Tap the button again!")
                    else:
                        st.session_state.generated_summary = result

        if st.session_state.generated_summary:
            st.markdown('<div class="workspace-card"><h4>📘 Tailored Summary</h4></div>', unsafe_allow_html=True)
            st.markdown(st.session_state.generated_summary)

    # ==========================================================
    # TAB 2 — CBT Objective Practice
    # ==========================================================
    with tab2:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
                <h3 style="margin:0;">CBT-Style Objective Practice</h3>
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
                    safe_combined_text = f"{chunk_1[:3500]}\n{chunk_2[:3500]}\n{chunk_3[:3500]}\n{chunk_4[:3500]}"
                    cbt_prompt = f"""
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
                    {safe_combined_text}
                    """
                    result = ask_gemini(api_key, cbt_prompt, dynamic_mode=False)
                    parsed = extract_json_block(result) if result else None
                    if not parsed:
                        st.error("Couldn't generate valid questions this time — please try again.")
                    else:
                        st.session_state.generated_cbt = parsed
                        st.session_state.cbt_user_answers = {}
                        st.session_state.cbt_submitted = False

        if st.session_state.generated_cbt:
            st.markdown('<div class="workspace-card"><h4>📝 Practice Quiz</h4></div>', unsafe_allow_html=True)

            with st.form("cbt_quiz_form"):
                for i, q in enumerate(st.session_state.generated_cbt):
                    st.markdown(f"**Q{i + 1}. {q.get('question', '')}**")
                    choice = st.radio(
                        label=f"question_{i}",
                        options=list(range(len(q.get("options", [])))),
                        format_func=lambda idx, opts=q.get("options", []): opts[idx],
                        key=f"cbt_radio_{i}",
                        label_visibility="collapsed"
                    )
                    st.session_state.cbt_user_answers[i] = choice
                    st.markdown("")
                submitted = st.form_submit_button("Submit Answers")
                if submitted:
                    st.session_state.cbt_submitted = True

            if st.session_state.cbt_submitted:
                score = 0
                for i, q in enumerate(st.session_state.generated_cbt):
                    user_choice = st.session_state.cbt_user_answers.get(i)
                    correct_index = q.get("correct_index")
                    is_correct = user_choice == correct_index
                    if is_correct:
                        score += 1
                        st.success(f"Q{i + 1}: Correct!")
                    else:
                        options = q.get("options", [])
                        correct_text = options[correct_index] if correct_index is not None and correct_index < len(options) else "N/A"
                        st.error(f"Q{i + 1}: Incorrect. Correct answer: {correct_text}")
                    if q.get("explanation"):
                        st.caption(q["explanation"])
                st.markdown(f"### Score: {score} / {len(st.session_state.generated_cbt)}")

    # ==========================================================
    # TAB 3 — Concept Map Table
    # ==========================================================
    with tab3:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="5" r="2"/><circle cx="5" cy="19" r="2"/><circle cx="19" cy="19" r="2"/><path d="M12 7v6M12 13l-6 4M12 13l6 4"/></svg>
                <h3 style="margin:0;">Concept Map Table</h3>
            </div>
        """, unsafe_allow_html=True)

        if st.button("Generate Concept Map", key="btn_concept_map"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Mapping out the key concepts and how they connect..."):
                    safe_combined_text = f"{chunk_1[:3500]}\n{chunk_2[:3500]}\n{chunk_3[:3500]}\n{chunk_4[:3500]}"
                    concept_prompt = f"""
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
                    {safe_combined_text}
                    """
                    result = ask_gemini(api_key, concept_prompt, dynamic_mode=False)
                    parsed = extract_json_block(result) if result else None
                    if not parsed:
                        st.error("Couldn't generate a concept map this time — please try again.")
                    else:
                        st.session_state.generated_concept_map = parsed

        if st.session_state.generated_concept_map:
            st.markdown('<div class="workspace-card"><h4>🧠 Key Concepts</h4></div>', unsafe_allow_html=True)
            st.table(st.session_state.generated_concept_map)

else:
    st.info("Upload a document above to unlock the Omni-Bar and all three workspace tabs.")
