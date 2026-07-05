import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random
import json
import re

APP_NAME = "DocDigest"

st.set_page_config(page_title=APP_NAME, layout="centered", initial_sidebar_state="expanded")

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
            padding: 16px 18px;
            border-radius: 8px;
            border: 1px solid rgba(109, 40, 217, 0.15);
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
if "generated_sections" not in st.session_state:
    st.session_state.generated_sections = None
if "generated_mode_label" not in st.session_state:
    st.session_state.generated_mode_label = None
if "generated_cbt" not in st.session_state:
    st.session_state.generated_cbt = None
if "current_cbt_batch" not in st.session_state:
    st.session_state.current_cbt_batch = None
if "generated_cheatsheet" not in st.session_state:
    st.session_state.generated_cheatsheet = None
if "active_view" not in st.session_state:
    st.session_state.active_view = None  # "analogy" | "quiz" | "cheatsheet"

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


def extract_json_block(text):
    """Only used for the paired notes/analogy feature, since that one needs
    structured per-item output to render as aligned rows."""
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


# --- SMART API KEY ROTATION SETUP (kept in the sidebar, same as original) ---
api_keys = []
if "GEMINI_API_KEY_1" in st.secrets and st.secrets["GEMINI_API_KEY_1"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_1"])
if "GEMINI_API_KEY_2" in st.secrets and st.secrets["GEMINI_API_KEY_2"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_2"])
if "GEMINI_API_KEY_3" in st.secrets and st.secrets["GEMINI_API_KEY_3"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_3"])

manual_key = st.sidebar.text_input("Backup API Key Entry (Optional)", type="password")
if manual_key:
    api_keys.append(manual_key)

api_key = random.choice(api_keys) if api_keys else None

# --- FILE UPLOAD STAYS IN THE MAIN AREA ---
uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX, PPTM)", type=["pdf", "docx", "pptx", "pptm"])

if uploaded_file:
    raw_text = extract_text(uploaded_file)
    total_length = len(raw_text) if raw_text else 0
    chunk_size = max(1, total_length // 4)

    chunk_1 = raw_text[0:chunk_size] if total_length > 0 else ""
    chunk_2 = raw_text[chunk_size:chunk_size*2] if total_length > 0 else ""
    chunk_3 = raw_text[chunk_size*2:chunk_size*3] if total_length > 0 else ""
    chunk_4 = raw_text[chunk_size*3:] if total_length > 0 else ""

    safe_combined_text = (
        f"[Introductory Section]\n{chunk_1[:3500]}\n\n"
        f"[Core Section A]\n{chunk_2[:3500]}\n\n"
        f"[Core Section B]\n{chunk_3[:3500]}\n\n"
        f"[Advanced / Concluding Section]\n{chunk_4[:3500]}"
    )

    # ============================================================
    # SIDEBAR — every control clustered here, locked open
    # ============================================================
    with st.sidebar:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
                <h3 style="margin:0;">Tailored Multi-Mode Explanation Engine</h3>
            </div>
        """, unsafe_allow_html=True)

        st.caption("Pick a style — each one breaks the document down item by item, note next to analogy.")

        MODE_OPTIONS = {
            "buddy": ("Campus Buddy", "You are a relatable university peer tutor. Use simple, engaging, and funny student/campus analogies (like hostel porters or campus gates) to explain everything simply. Highlight key terms in **bold**."),
            "street": ("Street-Smart", "You are a highly practical, street-smart operations mentor. Explain the concepts using crisp, real-world analogies based on everyday logic, physical logistics, spotting counterfeits, managing daily physical operations, or coordinating delivery logistics. Avoid corporate boardroom slangs and avoid campus-specific university terms. Make it punchy and clear for anyone living in the real world. Highlight key terms in **bold**."),
            "technical": ("Deep Technical", "You are a senior technical enterprise architect. Break down the content using exact, rigorous academic definitions, engineering mechanics, and precise technical infrastructure logic. Highlight key terms in **bold**."),
            "layman": ("Layman", "You are explaining this to a complete novice. Use ultra-simplified, everyday non-tech visual elements. Avoid any advanced technical concepts or assumptions of prior engineering knowledge. Highlight key terms in **bold**."),
        }

        for mode_id, (label, style_prompt) in MODE_OPTIONS.items():
            if st.button(label, key=f"mode_{mode_id}"):
                if not api_key:
                    st.error("Missing API Key!")
                elif not raw_text.strip():
                    st.error("Could not extract any text from this document.")
                else:
                    with st.spinner("Processing custom persona analytics across the full document..."):
                        prompt = f"""
                        Go through the document below and pull out EVERY distinct rule, policy, defined
                        term, or named concept as its OWN separate entry. Do not group multiple terms
                        into one entry, and do not merge several rules into a single paragraph.

                        For example, if the document mentions two sub-types of the same category,
                        they must be TWO separate entries, not one entry about the category in general.

                        There is no fixed number of entries — extract as many as the document actually
                        contains. Err on the side of splitting things apart rather than combining them.

                        For EACH entry, produce a matched pair: the literal factual note, and the
                        analogy explanation of that SAME single item, using the persona below.

                        Respond with ONLY a raw JSON array, no markdown fences, no commentary, in this
                        exact schema:
                        [
                          {{
                            "topic": "the specific rule/term name only",
                            "literal_note": "Strict, literal, factual explanation of THIS ONE item only. No analogies. 1-2 sentences.",
                            "analogy": "The SAME single item re-explained using the persona below. 1-2 sentences."
                          }}
                        ]

                        Persona: {style_prompt}

                        Study Text Sections:
                        {safe_combined_text}
                        """
                        result = ask_gemini(api_key, prompt, dynamic_mode=True)
                        parsed = extract_json_block(result)
                        if parsed:
                            st.session_state.generated_sections = parsed
                            st.session_state.generated_mode_label = label
                            st.session_state.active_view = "analogy"
                        else:
                            st.error("Couldn't generate a summary this time — please try again.")

        st.markdown("---")
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5a3 3 0 1 0-5.997.125 4 4 0 0 0-2.526 5.77 4 4 0 0 0 .556 6.588A4 4 0 1 0 12 18Z"/><path d="M12 5a3 3 0 1 1 5.997.125 4 4 0 0 1 2.526 5.77 4 4 0 0 1-.556 6.588A4 4 0 1 1 12 18Z"/><path d="M12 5v14"/><path d="M12 12h6"/><path d="M12 12H6"/><path d="M12 7h5"/><path d="M12 16h5"/><path d="M12 7H7"/><path d="M12 16H7"/></svg>
                <h3 style="margin:0;">Theory-to-CBT Objective Drill</h3>
            </div>
        """, unsafe_allow_html=True)

        st.caption("Select which block of the syllabus notes you want to generate questions from:")

        batch_selection = st.selectbox(
            "Choose Target Study Block:",
            ["Batch 1: Introduction & Foundation Concepts",
             "Batch 2: Core Methodologies & Process Details",
             "Batch 3: Deep Technical Content",
             "Batch 4: Advanced Scenarios & Conclusions"]
        )

        if batch_selection.startswith("Batch 1"):
            selected_text = chunk_1[:9000]
            start_num = 1
        elif batch_selection.startswith("Batch 2"):
            selected_text = chunk_2[:9000]
            start_num = 8
        elif batch_selection.startswith("Batch 3"):
            selected_text = chunk_3[:9000]
            start_num = 15
        else:
            selected_text = chunk_4[:9000]
            start_num = 22

        if st.button("Convert Selected Block to Questions", key="btn_cbt"):
            if not api_key:
                st.error("Missing API Key!")
            elif not selected_text.strip():
                st.error("This segment of the document doesn't contain enough text to extract data.")
            else:
                with st.spinner(f"Mining questions starting from Q{start_num} for this specific text block..."):
                    prompt = f"""
                    You are a strict Computer Based Test (CBT) examiner.
                    Analyze the uploaded text slice and transform the theoretical concepts into highly practical multiple-choice objective questions.
                    You must output exactly 7 to 8 distinct multiple-choice questions.
                    You MUST start numbering your output starting strictly from Question {start_num}.

                    STRICT FORMATTING MANDATE: Every single question must have exactly 4 options (A, B, C, D).
                    Follow this exact structure for every item:

                    **Question X: [Insert question here]**
                    A) [Option A]
                    B) [Option B]
                    C) [Option C]
                    D) [Option D]
                    * 👉 **Correct Answer:** || [State correct letter, followed by 1-sentence explanation] ||

                    Study Text Section:
                    {selected_text}
                    """
                    st.session_state.generated_cbt = ask_gemini(api_key, prompt, dynamic_mode=False)
                    st.session_state.current_cbt_batch = batch_selection
                    st.session_state.active_view = "quiz"

        st.markdown("---")
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v18"/><path d="M3 12h18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <h3 style="margin:0;">Key Acronyms, Definitions & Formulas</h3>
            </div>
        """, unsafe_allow_html=True)

        if st.button("Generate Cheat Sheet Table", key="btn_table"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Extracting terms and formulas from the entire document..."):
                    full_syllabus_context = (
                        f"[Section 1: Foundations]\n{chunk_1[:3000]}\n\n"
                        f"[Section 2: Core Details]\n{chunk_2[:3000]}\n\n"
                        f"[Section 3: Deep Technical]\n{chunk_3[:3000]}\n\n"
                        f"[Section 4: Advanced/Conclusion]\n{chunk_4[:3000]}"
                    )

                    prompt = f"""
                    Analyze the following study material and act as a professional summary engine.
                    Extract every single critical network protocol, acronym, core cybersecurity definition, concept, and formula into a clean, comprehensive summary grid.
                    You MUST format your entire output as a valid Markdown table with exactly two columns. Output ONLY the table itself.

                    | Term / Concept / Acronym | High-Yield Summary & Core Meaning |
                    | :--- | :--- |

                    Extract at least 12-18 foundational and advanced elements:
                    {full_syllabus_context}
                    """
                    st.session_state.generated_cheatsheet = ask_gemini(api_key, prompt, dynamic_mode=False)
                    st.session_state.active_view = "cheatsheet"

    # ============================================================
    # MAIN AREA — presentation of whichever result is active
    # ============================================================
    view = st.session_state.active_view

    if view is None:
        st.info("Document loaded. Choose an action from the left panel to get started.")

    elif view == "analogy":
        st.markdown(f"### 💡 {st.session_state.generated_mode_label} — Notes & Analogy")
        sections = st.session_state.generated_sections or []

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
                st.markdown(f'<div class="memo-card">{note or "—"}</div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="memo-card">{analogy or "—"}</div>', unsafe_allow_html=True)
            st.markdown("")

    elif view == "quiz":
        st.markdown(f"### 🧩 CBT Practice — {st.session_state.current_cbt_batch}")
        if st.session_state.generated_cbt:
            st.markdown(st.session_state.generated_cbt, unsafe_allow_html=True)

    elif view == "cheatsheet":
        st.markdown("### 🗂️ Cheat Sheet")
        if st.session_state.generated_cheatsheet:
            st.markdown(st.session_state.generated_cheatsheet, unsafe_allow_html=True)

else:
    st.info(f"Upload a document above to unlock {APP_NAME}'s tools.")
