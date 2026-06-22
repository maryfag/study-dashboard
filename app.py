import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random

st.set_page_config(page_title="Ultimate Study Dashboard", layout="centered")

# --- LUCIDE ICONS ENGINE SETUP ---
# This injects the official Lucide script so the browser can render the beautiful minimalist icons
st.markdown("""
    <script src="https://unpkg.com/lucide@latest"></script>
    <script>
        // Tell Lucide to look at the page and convert all icon tags instantly
        setTimeout(() => { lucide.createIcons(); }, 500);
        // Keep checking if new tabs are clicked so icons load dynamically
        setInterval(() => { lucide.createIcons(); }, 1500);
    </script>
    <style>
        .lucide-inline {
            vertical-align: middle;
            margin-right: 8px;
            display: inline-block;
            width: 20px;
            height: 20px;
        }
        div.stButton > button {
            display: flex;
            align-items: center;
            justify-content: center;
        }
    </style>
""", unsafe_allow_html=True)

# --- STATE MEMORY CORES: Keeps tabs from wiping out ---
if "generated_summary" not in st.session_state:
    st.session_state.generated_summary = None
if "generated_cbt" not in st.session_state:
    st.session_state.generated_cbt = None
if "current_cbt_batch" not in st.session_state:
    st.session_state.current_cbt_batch = None

# --- TITLE WITH TRUE LUCIDE BOOK STACK ---
st.markdown("""
    <h1 style='display: flex; align-items: center;'>
        <svg class='lucide-inline' xmlns='http://www.w3.org/2000/svg' width='40' height='40' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round' data-lucide='library'></svg>
        Your Ultimate Exam Survival Dashboard
    </h1>
""", unsafe_allow_html=True)

st.write("Upload your lecture notes, slides, or PDFs, then choose how you want to conquer them.")

# --- SMART API KEY ROTATION SETUP ---
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
    return "🚨 Server lines are busy. Tap the button again!"

uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX, PPTM)", type=["pdf", "docx", "pptx", "pptm"])

if uploaded_file:
    tab1, tab2, tab3 = st.tabs([
        "📖 Custom Explanation Summary", 
        "🧩 CBT Objective Practice", 
        "📊 Concept Map Table"
    ])
    
    raw_text = extract_text(uploaded_file)
    total_length = len(raw_text) if raw_text else 0
    chunk_size = max(1, total_length // 4)
    
    chunk_1 = raw_text[0:chunk_size] if total_length > 0 else ""
    chunk_2 = raw_text[chunk_size:chunk_size*2] if total_length > 0 else ""
    chunk_3 = raw_text[chunk_size*2:chunk_size*3] if total_length > 0 else ""
    chunk_4 = raw_text[chunk_size*3:] if total_length > 0 else ""

    with tab1:
        st.subheader("Tailored Multi-Mode Explanation Engine")
        
        explanation_mode = st.selectbox(
            "Choose Your Desired Explanation Persona:",
            ["✨ Campus Buddy Mode (Student & Campus Analogies)",
             "🛣️ Street-Smart Analogy Mode (Practical, Everyday Logic & Logistics)", 
             "🧠 Deep Technical Mode (Upper-Level Technical Rigor)", 
             "🧸 Layman Mode (Explain Like I'm 5 Style)"]
        )
        
        if st.button("🚀 Generate Tailored Summary"):
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
                        style_prompt = "You are a highly practical, street-smart operations mentor. Explain the concepts using crisp, real-world analogies based on everyday logic, physical logistics, spotting counterfeits, managing daily physical operations, or coordinating delivery logistics. Avoid corporate boardroom slangs and avoid campus-specific university terms. Make it punchy and clear for anyone living in the real world. Highlight key terms in **bold**."
                    elif "Deep Technical" in explanation_mode:
                        style_prompt = "You are a senior technical enterprise architect. Break down the content using exact, rigorous academic definitions, engineering mechanics, and precise technical infrastructure logic. Highlight key terms in **bold**."
                    else:
                        style_prompt = "You are explaining this to a complete novice. Use ultra-simplified, everyday non-tech visual elements. Avoid any advanced technical concepts or assumptions of prior engineering knowledge. Highlight key terms in **bold**."
                    
                    prompt = f"""
                    You are an expert personalized summary processor. 
                    {style_prompt}
                    Ensure your explanation covers the entire document provided below sequentially from start to finish.
                    
                    Study Text Sections:
                    {safe_combined_text}
                    """
                    st.session_state.generated_summary = ask_gemini(api_key, prompt, dynamic_mode=True)

        if st.session_state.generated_summary:
            st.markdown(st.session_state.generated_summary)

    with tab2:
        st.subheader("Theory-to-CBT Objective Drill")
        st.write("Select which block of the syllabus notes you want to generate questions from:")
        
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

        if st.button("🚀 Convert Selected Block to Questions"):
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

        if st.session_state.generated_cbt and st.session_state.current_cbt_batch == batch_selection:
            st.markdown(st.session_state.generated_cbt)

    with tab3:
        st.subheader("Key Acronyms, Definitions & Formulas")
        if st.button("🚀 Generate Cheat Sheet Table"):
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
                    st.markdown(ask_gemini(api_key, prompt, dynamic_mode=False))
