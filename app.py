import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random

# --- CONFIGURATION & BRANDING ---
st.set_page_config(page_title="DocDigest | Professional Intelligence Workspace", layout="wide")

# --- EXECUTIVE UI STYLING ---
st.markdown("""
    <style>
        .flex-container {
            display: flex;
            align-items: center;
            gap: 14px;
            margin-bottom: 12px;
        }
        .icon-svg {
            color: #6D28D9; /* Executive Purple Accent */
            flex-shrink: 0;
            vertical-align: middle;
        }
        .workspace-card {
            background-color: rgba(255, 255, 255, 0.05);
            padding: 20px;
            border-radius: 10px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            margin-bottom: 15px;
        }
    </style>
""", unsafe_allow_html=True)

# --- STATE MEMORY CORES ---
if "generated_summary" not in st.session_state:
    st.session_state.generated_summary = None
if "generated_analogy" not in st.session_state:
    st.session_state.generated_analogy = None
if "generated_cbt" not in st.session_state:
    st.session_state.generated_cbt = None
if "current_cbt_batch" not in st.session_state:
    st.session_state.current_cbt_batch = None
if "last_action" not in st.session_state:
    st.session_state.last_action = None

# --- HEADER ASSEMBLY ---
st.markdown("""
    <div class="flex-container">
        <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="38" height="38" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z"/>
            <path d="M12 6v6l4 2"/>
        </svg>
        <h1 style="margin: 0; padding: 0; font-size: 2.4rem; font-weight: 700;">DocDigest Workspace</h1>
    </div>
""", unsafe_allow_html=True)
st.caption("Deconstruct complex technical manuals, corporate documentation, and dense academic text into intuitive models.")

# --- API ROTATION CORE ---
api_keys = []
for key_name in ["GEMINI_API_KEY_1", "GEMINI_API_KEY_2", "GEMINI_API_KEY_3"]:
    if key_name in st.secrets and st.secrets[key_name]:
        api_keys.append(st.secrets[key_name])

manual_key = st.sidebar.text_input("Backup API Key Entry (Optional)", type="password")
if manual_key:
    api_keys.append(manual_key)

api_key = random.choice(api_keys) if api_keys else None

# --- PARSING ENGINE ---
def extract_text(uploaded_file):
    filename = uploaded_file.name
    text = ""
    try:
        if filename.endswith('.pdf'):
            pdf_reader = pypdf.PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                t = page.extract_text()
                if t: text += t + "\n"
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
    
    return "⚠️ Context processing line is busy. Please re-trigger the action."

# --- FILE INPUT LAYER ---
uploaded_file = st.file_uploader("Upload operational documents (PDF, DOCX, PPTX)", type=["pdf", "docx", "pptx", "pptm"])

if uploaded_file:
    raw_text = extract_text(uploaded_file)
    total_length = len(raw_text) if raw_text else 0
    chunk_size = max(1, total_length // 4)
    
    chunk_1 = raw_text[0:chunk_size] if total_length > 0 else ""
    chunk_2 = raw_text[chunk_size:chunk_size*2] if total_length > 0 else ""
    chunk_3 = raw_text[chunk_size*2:chunk_size*3] if total_length > 0 else ""
    chunk_4 = raw_text[chunk_size*3:] if total_length > 0 else ""

    # --- DYNAMIC COMMAND INTERFACE (THE SHORTCUT SUGGESTER) ---
    st.markdown("### ⚡ Command Palette")
    user_query = st.text_input("Type a shortcut parameter letter (e.g., 'G' for Generate, 'E' for Extract, 'R' for Risk)...", placeholder="Type a letter to reveal smart commands...")
    
    # Pre-built action database filtered instantly in-memory (0 Token Cost)
    master_commands = [
        {"trigger": "G", "label": "Generate Executive Summary & Analogy Map", "type": "summary"},
        {"trigger": "E", "label": "Extract Tasks, Deadlines & Requirements", "type": "extract"},
        {"trigger": "R", "label": "Run Corporate Risk & Flaw Assessment", "type": "risk"},
        {"trigger": "C", "label": "Count Local Document Metrics", "type": "count"}
    ]
    
    if user_query:
        filtered_commands = [cmd for cmd in master_commands if cmd["trigger"].lower() == user_query.strip().lower()[:1]]
        
        if filtered_commands:
            st.write("💡 Suggested Workspace Macros:")
            for cmd in filtered_commands:
                if st.button(cmd["label"], key=f"macro_{cmd['type']}"):
                    st.session_state.last_action = cmd["type"]
        else:
            st.info("No matching macros found for that letter shortcut yet. Try 'G', 'E', 'R', or 'C'.")

    # --- EXECUTION OF MACRO SHORTCUTS ---
    if st.session_state.last_action == "count":
        st.success("Document Metric Analysis complete (Calculated Locally).")
        st.metric(label="Total Character Count", value=total_length)
        st.metric(label="Estimated Word Count", value=len(raw_text.split()))
        st.session_state.last_action = None

    elif st.session_state.last_action == "extract":
        with st.spinner("Extracting timeline requirements and tasks..."):
            prompt = f"Analyze the following text and extract clear lists of tasks, formal requirements, deadlines, and key dates. Format nicely.\n\nContext:\n{chunk_1[:8000]}"
            st.markdown(ask_gemini(api_key, prompt))
            st.session_state.last_action = None

    elif st.session_state.last_action == "risk":
        with st.spinner("Running system risk assessment..."):
            prompt = f"Identify hidden operational risks, technical vulnerabilities, or logical flaws present in this text document documentation.\n\nContext:\n{chunk_1[:8000]}"
            st.markdown(ask_gemini(api_key, prompt))
            st.session_state.last_action = None

    st.markdown("---")

    # --- WORKSPACE TABS ---
    tab1, tab2, tab3 = st.tabs([
        "Dual-Card Reframer Engine", 
        "Objective Practice Drills", 
        "Data Extraction Matrix"
    ])
    
    with tab1:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
                <h3 style="margin:0;">Interactive Dual-Reframer</h3>
            </div>
        """, unsafe_allow_html=True)
        
        explanation_mode = st.selectbox(
            "Select Translation Paradigm:",
            ["Street-Smart Logistics (Real-world Operations Frameworks)",
             "Corporate Decoupler (Translating Jargon to Plain English)", 
             "Deep Technical Architecture (Rigor & Specification Frameworks)"]
        )
        
        if st.button("Execute Framework Realignment", key="btn_summary") or st.session_state.last_action == "summary":
            st.session_state.last_action = None
            if not api_key:
                st.error("API Architecture Key Error.")
            elif not raw_text.strip():
                st.error("Empty Data Core.")
            else:
                with st.spinner("Processing split-plane layout analytics..."):
                    safe_combined_text = f"{chunk_1[:4000]}\n{chunk_2[:4000]}"
                    
                    # Exact Summary Ground Truth Prompt
                    summary_prompt = f"Extract and organize the literal technical facts, core configurations, definitions, and rules contained in this text. Maintain maximum precision and clear terminology:\n\n{safe_combined_text}"
                    
                    # Analogy Persona Prompt
                    if "Street-Smart" in explanation_mode:
                        style_prompt = "Explain the mechanisms in this text using crisp, real-world analogies based on physical logistics, shipping centers, traffic dynamics, or daily operations. Avoid abstract academic words."
                    elif "Corporate Decoupler" in explanation_mode:
                        style_prompt = "Strip away all corporate jargon, buzzwords, and intense legal speech. Explain the practical utility like you are speaking to a general stakeholder."
                    else:
                        style_prompt = "Act as an enterprise systems architect. Map the mechanics onto infrastructural, mechanical, or hardware engineering dependencies."
                        
                    analogy_prompt = f"{style_prompt}\n\nContext to translate:\n{safe_combined_text}"
                    
                    st.session_state.generated_summary = ask_gemini(api_key, summary_prompt, dynamic_mode=False)
                    st.session_state.generated_analogy = ask_gemini(api_key, analogy_prompt, dynamic_mode=True)

        # --- SIDE-BY-SIDE DUAL VIEW CARD IMPLEMENTATION ---
        if st.session_state.generated_summary and st.session_state.generated_analogy:
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<div class="workspace-card"><h4>📖 Grounded Truth (Exact Documentation)</h4></div>', unsafe_allow_html=True)
                st.markdown(st.session_state.generated_summary)
                
            with col2:
                st.markdown('<div class="workspace-card"><h4>💡 Vibe Check (System Analogy Mapping)</h4></div>', unsafe_allow_html=True)
                st.markdown(st.session_state.generated_analogy)

    with tab2:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5a3 3 0 1 0-5.997.125 4 4 0 0 0-2.526 5.77 4 4 0 0 0 .556 6.588A4 4 0 1 0 12 18Z"/><path d="M12 5a3 3 0 1 1 5.997.125 4 4 0 0 1 2.526 5.77 4 4 0 0 1-.556 6.588A4 4 0 1 1 12 18Z"/><path d="M12 5v14"/><path d="M12 12h6"/><path d="M12 12H6"/></svg>
                <h3 style="margin:0;">Operational Assessment Drill</h3>
            </div>
        """, unsafe_allow_html=True)
        
        batch_selection = st.selectbox(
            "Select Processing Segment Layer:",
            ["Segment Block 1: Foundations & Architecture", 
             "Segment Block 2: Methodologies & Workflows", 
             "Segment Block 3: Deep Configuration Specs"]
        )
        
        selected_text = chunk_1[:9000] if "1" in batch_selection else (chunk_2[:9000] if "2" in batch_selection else chunk_3[:9000])
        start_num = 1 if "1" in batch_selection else (8 if "2" in batch_selection else 15)

        if st.button("Generate Assessment Matrix", key="btn_cbt"):
            with st.spinner("Compiling diagnostic evaluation questions..."):
                prompt = f"""
                Act as a strict systems examiner. Create exactly 6-7 high-quality diagnostic multiple-choice questions from this text. Start numbering from {start_num}.
                Format like this:
                **Question X: [Context Statement]**
                A) [Option]
                B) [Option]
                C) [Option]
                D) [Option]
                * 👉 **Correct Answer:** || [Letter & Technical Rationale] ||
                
                Context:
                {selected_text}
                """
                st.session_state.generated_cbt = ask_gemini(api_key, prompt, dynamic_mode=False)
                st.session_state.current_cbt_batch = batch_selection

        if st.session_state.generated_cbt and st.session_state.current_cbt_batch == batch_selection:
            st.markdown(st.session_state.generated_cbt)

    with tab3:
        st.markdown("""
            <div class="flex-container">
                <svg class="icon-svg" xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v18"/><path d="M3 12h18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <h3 style="margin:0;">Operational Metrics & Terms Matrix</h3>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("Compile Reference Matrix Table", key="btn_table"):
            with st.spinner("Structuring parameters..."):
                prompt = f"""
                Extract every core system parameter, key acronym, definition, technical workflow item, or specification equation into a crisp two-column Markdown matrix table.
                
                | Core Parameter / Term / Specification Element | High-Yield Technical Translation & Meaning |
                | :--- | :--- |
                
                Context:
                {chunk_1[:4000]}\n{chunk_2[:4000]}
                """
                st.markdown(ask_gemini(api_key, prompt, dynamic_mode=False))
