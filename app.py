
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
        elif filename.endswith('.pptx'):
            prs = Presentation(uploaded_file)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text += shape.text.strip() + "\n"
    except Exception as e:
        return f"Error reading file structure: {e}"
    return text

def ask_gemini(api_key, prompt_text):
    models = ["gemini-2.5-flash", "gemini-1.5-flash"]
    
    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt_text}]}]}
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
    return "🚨 Both public AI server lines are packed right now. Tap the button again in 10 seconds!"

uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX)", type=["pdf", "docx", "pptx"])

if uploaded_file:
    # Retaining your clean, working layout tabs
    tab1, tab2, tab3 = st.tabs(["✨ ELI5 Summary", "🧠 CBT Objective Practice", "📋 Concept Map Table"])
    
    raw_text = extract_text(uploaded_file)
    truncated_text = raw_text[:9000] if raw_text else ""

    with tab1:
        st.subheader("Plain English Breakdown")
        if st.button("🚀 Generate Summary Now"):
            if not api_key:
                st.error("Missing API Key! Please save your key in the Streamlit Secrets Vault or paste it into the left sidebar box.")
            elif not truncated_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Stripping out jargon..."):
                    prompt = f"Analyze this study material and give me a comprehensive 'Explain Like I'm 5' summary. Break down complex jargon into simple, memorable analogies:\n\n{truncated_text}"
                    st.markdown(ask_gemini(api_key, prompt))

    with tab2:
        st.subheader("🤖 Theory-to-CBT Objective Drill")
        if st.button("🚀 Convert to Objective Questions"):
            if not api_key:
                st.error("Missing API Key! Please save your key in the Streamlit Secrets Vault or paste it into the left sidebar box.")
            elif not truncated_text.strip():
                st.error("Could not extract any text from this document.")
            else:
                with st.spinner("Converting theoretical material into CBT objective format..."):
                    # UPGRADED PROMPT: Specifically engineered for CBT/OBJ practice
                    prompt = f"""
                    You are an expert examiner building a Computer Based Test (CBT). 
                    Analyze the uploaded text or tutorial questions, extract the most critical theoretical points, and convert them into a standard objective exam style.
                    
                    Please generate 5-8 multiple-choice or fill-in-the-blank objective questions. 
                    
                    CRITICAL FORMATTING INSTRUCTIONS:
                    Format each question explicitly using standard Markdown notation like this:
                    
                    **Question X: [Insert clear, objective question here]**
                    A) [Option A]
                    B) [Option B]
                    C) [Option C]
                    D) [Option D]
                    * 👉 **Correct Answer:** || [State the letter and a brief 1-sentence reason why it is correct] ||
                    
                    Ensure the answers stay completely hidden inside the spoiler block so users can self-test.
                    
                    Study Text / Tutorial Sheet:
                    {truncated_text}
                    """
                    st.markdown(ask_gemini(api_key, prompt))

    with tab3:
        st.subheader("Key Acronyms, Definitions & Formulas")
        if st.button("🚀 Generate Cheat Sheet Table"):
            if not api_key:
                st.error("Missing API Key! Please save your key in the Streamlit Secrets Vault or paste it into the left sidebar box.")
            elif not truncated_text.strip():
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
                    {truncated_text}
                    """
                    st.markdown(ask_gemini(api_key, prompt))
                  