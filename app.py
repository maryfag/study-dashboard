import streamlit as st
import pypdf
from docx import Document
from pptx import Presentation
import requests
import random

st.set_page_config(page_title="Ultimate Study Dashboard", layout="centered")
st.title("📚 Your Ultimate Exam Survival Dashboard")
st.write("Upload a slide, doc, or PDF, then choose how you want to conquer it.")

# --- SMART API KEY ROTATION SETUP ---
api_keys = []

if "GEMINI_API_KEY_1" in st.secrets and st.secrets["GEMINI_API_KEY_1"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_1"])
if "GEMINI_API_KEY_2" in st.secrets and st.secrets["GEMINI_API_KEY_2"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_2"])
if "GEMINI_API_KEY_3" in st.secrets and st.secrets["GEMINI_API_KEY_3"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_3"])

manual_key = st.sidebar.text_input("Backup API Key Entry (Only needed if Secrets Vault is empty)", type="password")
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
                except:
                    continue
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
        except:
            continue
    return "🚨 Server lines are busy. Tap the button again!"

uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX, PPTM)", type=["pdf", "docx", "pptx", "pptm"])

if uploaded_file:
    tab1, tab2, tab3 = st.tabs(["✨ Clear Practical Summary", "🧠 CBT Objective Practice", "📋 Concept Map Table"])
    
    # Gather raw text from file
    raw_text = extract_text(uploaded_file)
    
    # AUTOMATIC TEXT CHUNKING ENGINE (Splits document text into 4 clean parts)
    total_length = len(raw_text) if raw_text else 0
    chunk_size = max(1, total_length // 4)
    
    chunk_1 = raw_text[0:chunk_size] if total_length > 0 else ""
    chunk_2 = raw_text[chunk_size:chunk_size*2] if total_length > 0 else ""
    chunk_3 = raw_text[chunk_size*2:chunk_size*3] if total_length > 0 else ""
    chunk_4 = raw_text[chunk_size*3:] if total_length > 0 else ""

    with tab1:
        st.subheader("High-Yield Plain English Breakdown")
        st.caption("💡 This engine scans all pages of your file and replaces childish examples with real-world tech analogies.")
        if st.button("🚀 Generate Full-Syllabus Summary Now"):
            if not api_key:
                st.error("Missing API Key!")
            elif not raw_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Analyzing all document layers... This takes a few seconds extra as we are scanning the entire file..."):
                    # FULL DOCUMENT STREAMING WORKAROUND:
                    # We take safely sized slices from all four parts of the document so the AI sees everything from beginning to end
                    safe_combined_text = (
                        f"--- START OF TEXT ---\n{chunk_1[:3000]}\n\n"
                        f"--- MIDDLE SECTION A ---\n{chunk_2[:3000]}\n\n"
                        f"--- MIDDLE SECTION B ---\n{chunk_3[:3000]}\n\n"
                        f"--- CLOSING/ADVANCED SECTION ---\n{chunk_4[:3000]}"
                    )
                    
                    prompt = f"""
                    You are a world-class practical tech instructor and senior engineer. 
                    Your job is to analyze the entire technical study curriculum provided below and break it down into an intuitive, high-yield professional summary.
                    
                    CRITICAL INSTRUCTIONS:
                    1. DO NOT use childish analogies (No toys, sandboxes, babies, or playgrounds). 
                    2. Instead, use real-world professional, everyday, or technical analogies (e.g., comparing networks to highways, databases to filing cabinets, firewalls to security checkpoints).
                    3. Ensure you capture key takeaways from the introduction, the middle core components, and the final concluding sections of the material so the student gets full coverage.
                    4. Keep the tone mature, crisp, and focused on helping a student deeply master the concepts for an upper-level examination.
                    
                    Study Text Sections:
                    {safe_combined_text}
                    """
                    st.markdown(ask_gemini(api_key, prompt, dynamic_mode=True))

    with tab2:
        st.subheader("🤖 Theory-to-CBT Objective Drill")
        st.write("Select which depth of the syllabus notes you want to generate questions from:")
        
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
                st.error("This segment of the document doesn't contain enough text to extract data. Try an earlier batch!")
            else:
                with st.spinner(f"Mining questions starting from Q{start_num} for this specific text block..."):
                    prompt = f"""
                    You are a strict Computer Based Test (CBT) examiner. 
                    Analyze the uploaded text slice and transform the theoretical concepts into highly practical multiple-choice objective questions.
                    
                    You must output exactly 7 to 8 distinct multiple-choice questions.
                    You MUST start numbering your output starting strictly from Question {start_num}.
                    
                    STRICT FORMATTING MANDATE:
                    Every single question must have exactly 4 options (A, B, C, D). You are strictly forbidden from writing open-ended or long-form essay questions.
                    
                    Follow this exact structure for every item:
                    
                    **Question X: [Insert the objective/multiple-choice question here]**
                    A) [Option A]
                    B) [Option B]
                    C) [Option C]
                    D) [Option D]
                    * 👉 **Correct Answer:** || [State the correct letter option, followed by a crisp 1-sentence explanation of why it is the correct answer] ||
                    
                    Study Text Section:
                    {selected_text}
                    """
                    st.markdown(ask_gemini(api_key, prompt, dynamic_mode=False))

    with tab3:
        st.subheader("Key Acronyms, Definitions & Formulas")
        if st.button("🚀 Generate Cheat Sheet Table"):
            if not api_key:
                st.error("Missing API Key!")
            elif not chunk_1.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Extracting terms and formulas into a table..."):
                    prompt = f"""
                    Analyze the following study material and act as a professional summary engine.
                    Extract every single critical term, network protocol, acronym, core definition, and formula into a clean, comprehensive summary grid.
                    
                    CRITICAL FORMATTING INSTRUCTIONS:
                    You MUST format your entire output as a valid Markdown table with exactly two columns. Do not include any conversational intro or outro text. Output ONLY the table.
                    
                    | Term / Concept / Acronym | High-Yield Summary & Core Meaning |
                    | :--- | :--- |
                    
                    Extract at least 8-12 foundational elements from the text below:
                    
                    Study Text:
                    {chunk_1[:9000]}
                    """
                    st.markdown(ask_gemini(api_key, prompt, dynamic_mode=False))
