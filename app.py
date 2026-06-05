import streamlit as st
import PyPDF2
from docx import Document
from pptx import Presentation
import requests

st.set_page_config(page_title="Ultimate Study Dashboard", layout="centered")
st.title("📚 Your Ultimate Exam Survival Dashboard")
st.write("Upload a slide, doc, or PDF, then choose how you want to conquer it.")

# Automatically uses your hidden key if it's in the vault
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

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
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    headers = {'Content-Type': 'application/json'}
    payload = {"contents": [{"parts": [{"text": prompt_text}]}]}
    try:
        response = requests.post(url, headers=headers, json=payload)
        return response.json()['candidates'][0]['content']['parts'][0]['text']
    except Exception as e:
        return f"Error connecting to API. Double-check your key! Details: {e}"

uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX)", type=["pdf", "docx", "pptx"])

if uploaded_file:
    st.info("🔄 File detected! Add your key in the sidebar and select a tool below when ready.")
    
    tab1, tab2, tab3 = st.tabs(["✨ ELI5 Summary", "🧠 Active Recall Quiz", "📋 Concept Map Table"])
    
    with tab1:
        st.subheader("Plain English Breakdown")
        if st.button("🚀 Generate Summary Now"):
            if not api_key:
                st.error("Please enter your Gemini API Key in the sidebar first!")
            else:
                with st.spinner("Extracting text and stripping out jargon..."):
                    file_text = extract_text(uploaded_file)
                    prompt = f"Analyze this study material and give me a comprehensive 'Explain Like I'm 5' summary. Break down complex jargon into simple, memorable analogies:\n\n{file_text[:12000]}"
                    st.write(ask_gemini(api_key, prompt))

    with tab2:
        st.subheader("Test Your Knowledge")
        if st.button("🚀 Generate Active Recall Quiz"):
            if not api_key:
                st.error("Please enter your Gemini API Key in the sidebar first!")
            else:
                with st.spinner("Building interactive flashcard questions..."):
                    file_text = extract_text(uploaded_file)
                    prompt = f"Based on this material, generate 5-10 challenging conceptual active recall questions. Under each question, provide the correct answer hidden clearly below it so I can test myself:\n\n{file_text[:12000]}"
                    st.write(ask_gemini(api_key, prompt))

    with tab3:
        st.subheader("Key Acronyms, Definitions & Formulas")
        if st.button("🚀 Generate Cheat Sheet Table"):
            if not api_key:
                st.error("Please enter your Gemini API Key in the sidebar first!")
            else:
                with st.spinner("Extracting terms and formulas into a table..."):
                    file_text = extract_text(uploaded_file)
                    prompt = f"Extract all critical terms, acronyms, definitions, and formulas from this text. Format them explicitly as a Markdown table with columns for 'Term/Concept' and 'Simple Definition/Formula':\n\n{file_text[:12000]}"
                    st.write(ask_gemini(api_key, prompt))
