import streamlit as st
import PyPDF2
from docx import Document
from pptx import Presentation
import requests
import random

st.set_page_config(page_title="Ultimate Study Dashboard", layout="centered")
st.title("📚 Your Ultimate Exam Survival Dashboard")
st.write("Upload a slide, doc, or PDF, then choose how you want to conquer it.")

# --- SMART API KEY ROTATION SETUP ---
api_keys = []

# Check the vault for up to 3 separate free keys
if "GEMINI_API_KEY_1" in st.secrets and st.secrets["GEMINI_API_KEY_1"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_1"])
if "GEMINI_API_KEY_2" in st.secrets and st.secrets["GEMINI_API_KEY_2"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_2"])
if "GEMINI_API_KEY_3" in st.secrets and st.secrets["GEMINI_API_KEY_3"]:
    api_keys.append(st.secrets["GEMINI_API_KEY_3"])

# Fallback to manual sidebar entry if the vault is completely empty
if not api_keys:
    manual_key = st.sidebar.text_input("Enter Gemini API Key", type="password")
    if manual_key:
        api_keys.append(manual_key)

# Pick a random key from the active pool for this request to share the traffic load!
api_key = random.choice(api_keys) if api_keys else None

def extract_text(uploaded_file):
    filename = uploaded_file.name
    text = ""
    try:
        if filename.endswith('.pdf'):
            pdf_reader = PyPDF2.PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                t = page.extract_text()
                if t: text += t + "\n"
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
        return f"Error reading file text: {e}"
    return text

def ask_gemini(api_key, prompt_text):
    # Try the main model first; fallback to 8b if the server is crowded
    models = ["gemini-2.5-flash", "gemini-2.5-flash-8b"]
    
    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt_text}]}]}
        
        try:
            response = requests.post(url, headers=headers, json=payload)
            response_json = response.json()
            
            # Safe validation check to prevent the 'candidates' crash
            if 'candidates' in response_json and response_json['candidates']:
                return response_json['candidates'][0]['content']['parts'][0]['text']
            elif 'error' in response_json:
                error_msg = response_json['error']['message']
                # If it's a high demand or quota issue, loop to the backup model
                if "demand" in error_msg.lower() or "quota" in error_msg.lower():
                    continue
                return f"Google API Error: {error_msg}"
        except Exception as e:
            continue
            
    return "🚨 Both public AI server lines are packed right now. Tap the button again in 10 seconds to bypass the crowd!"

uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX)", type=["pdf", "docx", "pptx"])

if uploaded_file:
    if not api_key:
        st.warning("🔑 Please add your API key to the sidebar (or save it in Secrets) to unlock the system!")
    else:
        st.info("🔄 File detected and API key authenticated! Select a tool below to generate.")
        
        tab1, tab2, tab3 = st.tabs(["✨ ELI5 Summary", "🧠 Active Recall Quiz", "📋 Concept Map Table"])
        
        with st.spinner("Processing text..."):
            raw_text = extract_text(uploaded_file)
            truncated_text = raw_text[:9000] 

        with tab1:
            st.subheader("Plain English Breakdown")
            if st.button("🚀 Generate Summary Now"):
                with st.spinner("Stripping out jargon..."):
                    prompt = f"Analyze this study material and give me a comprehensive 'Explain Like I'm 5' summary. Break down complex jargon into simple, memorable analogies:\n\n{truncated_text}"
                    st.markdown(ask_gemini(api_key, prompt))

        with tab2:
            st.subheader("Test Your Knowledge")
            if st.button("🚀 Generate Active Recall Quiz"):
                with st.spinner("Building interactive flashcard questions..."):
                    prompt = f"""
                    You are an expert professor. Analyze the uploaded document thoroughly. 
                    Your job is to build an interactive Active Recall study quiz by pulling distinct conceptual questions across different sections, chapters, or slides of the material.
                    
                    CRITICAL FORMATTING INSTRUCTIONS:
                    For each question, format it EXACTLY like this using a standard Markdown text block. Do not use complex HTML. 
                    
                    **Question X: [Insert clear, challenging question from the material]**
                    * 👉 **Answer:** || [Insert the precise, detailed answer here] ||
                    
                    (Note: Placing the answer between '||' pipes creates a native spoiler tag on many platforms, or keeps it cleanly inline so students can hide it with their hands or scroll carefully). Make sure you extract 5-8 solid questions.
                    
                    Study Text:
                    {truncated_text}
                    """
                    st.markdown(ask_gemini(api_key, prompt))

        with tab3:
            st.subheader("Key Acronyms, Definitions & Formulas")
            if st.button("🚀 Generate Cheat Sheet Table"):
                with st.spinner("Extracting terms and formulas into a table..."):
                    prompt = f"""
                    Analyze the following study material and act as a professional summary engine.
                    Extract every single critical term, network protocol, acronym, core definition, and formula into a clean cheat sheet.
                    
                    CRITICAL FORMATTING INSTRUCTIONS:
                    You MUST format your entire output as a valid Markdown table with exactly two columns. Do not include any conversational intro or outro text. Output ONLY the table.
                    
                    | Term / Concept / Acronym | High-Yield Summary & Core Meaning |
                    | :--- | :--- |
                    
                    Extract at least 8-12 foundational elements from the text below to make a comprehensive summary grid:
                    
                    Study Text:
                    {truncated_text}
                    """
                    st.markdown(ask_gemini(api_key, prompt))
