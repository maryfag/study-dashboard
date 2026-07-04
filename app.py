import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random

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
if "generated_analogy" not in st.session_state:
    st.session_state.generated_analogy = None
if "generated_cbt" not in st.session_state:
    st.session_state.generated_cbt = None
if "current_cbt_batch" not in st.session_state:
    st.session_state.current_cbt_batch = None
if "omni_bar_response" not in st.session_state:
    st.session_state.omni_bar_response = None
if "last_processed_query" not in st.session_state:
    st.session_state.last_processed_query = None

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
                continue
        except: continue
    
    return "Server lines are busy. Tap the button again!"

uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX, PPTM)", type=["pdf", "docx", "pptx", "pptm"])

if uploaded_file:
    raw_text = extract_text(uploaded_file)
    total_length = len(raw_text) if raw_text else 0
    chunk_size = max(1, total_length // 4)
    
    chunk_1 = raw_text[0:chunk_size] if total_length > 0 else ""
    chunk_2 = raw_text[chunk_size:chunk_size*2] if total_length > 0 else ""
    chunk_3 = raw_text[chunk_size*2:chunk_size*3] if total_length > 0 else ""
    chunk_4 = raw_text[chunk_size*3:] if total_length > 0 else ""

    # --- TRUE AI OMNI-BAR COMMAND PALETTE ---
    st.markdown("### ⚡ AI Omni-Bar Command Palette")
    shortcut_input = st.text_input("Ask Gemini anything about this document (e.g., 'Extract the major risks', 'Give me a list of terms')...", placeholder="Type your dynamic query and press Enter...")
    
    if shortcut_input:
        cleaned_query = shortcut_input.strip()
        # Fire API request only if the query has changed to avoid double execution on re-runs
        if cleaned_query != st.session_state.last_processed_query:
            if not api_key:
                st.error("Missing API Key!")
            else:
                with st.spinner("Gemini is searching and analyzing the document layout..."):
                    context_sample = f"[Document Excerpt Section]\n{chunk_1[:5000]}\n{chunk_2[:3000]}"
                    omni_prompt = f"""
                    You are a real-time smart search engine assistant for this document. 
                    Answering the following custom user query based strictly on the document context provided below.
                    
                    User Query: {cleaned_query}
                    
                    Document Context:
                    {context_sample}
                    """
                    st.session_state.omni_bar_response = ask_gemini(api_key, omni_prompt, dynamic_mode=True)
                    st.session_state.last_processed_query = cleaned_query

    # Display the live AI Omni-Bar results persistently
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
                    
                    ground_truth_prompt = f"You are an exact, literal translation processor. Extract and summarize the strict facts, core technical configurations, definitions, and actual text rules present in this document. Do not use creative shortcuts. Text:\n{safe_combined_text}"
                    
                    if "Campus Buddy" in explanation_mode:
                        style_prompt = "You are a relatable university peer tutor. Use simple, engaging, and funny student/campus analogies (like hostel porters or campus gates) to explain everything simply. Highlight key terms in **bold**."
                    elif "Street-Smart" in explanation_mode:
                        style_prompt = "You are a highly practical, street-smart operations mentor. Explain the concepts using crisp, real-world analogies based on everyday logic, physical logistics, spotting counterfeits, managing daily physical operations, or coordinating delivery logistics. Avoid corporate boardroom slangs and avoid campus-specific university terms. Make it punchy and clear for anyone living in the real world. Highlight key terms in **bold**."
                    elif "Corporate Decoupler" in explanation_mode:
                        style_prompt = "You are a clean communicator stripping away complex corporate autopilot jargon, buzzwords, or intense contractual terminology. Re-evaluate the document and write out its practical mechanism cleanly using general real-world metaphors. Highlight key terms in **bold**."
                    elif "Deep Technical" in explanation_mode:
                        style_
