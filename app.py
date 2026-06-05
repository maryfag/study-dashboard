import streamlit as st
import PyPDF2
from docx import Document
from pptx import Presentation
import requests

st.set_page_config(page_title="Ultimate Study Dashboard", layout="centered")
st.title("📚 Your Ultimate Exam Survival Dashboard")
st.write("Upload a slide, doc, or PDF, then choose how you want to conquer it.")

api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

def extract_text(uploaded_file):
    filename = uploaded_file.name
    text = ""
    if filename.endswith('.pdf'):
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
    elif filename.endswith('.docx'):
        doc = Document(uploaded_file)
        for para in doc.paragraphs:
            text += para.text + "\n"
    elif filename.endswith('.pptx'):
        prs = Presentation(uploaded_file)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + "\n"
    return text

def ask_gemini(api_key, prompt_text):
    # Direct API request to completely bypass heavy library installations
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    headers = {'Content-Type': 'application/json'}
    payload = {"contents": [{"parts": [{"text": prompt_text}]}]}
    
    try:
        response = requests.post(url, headers=headers, json=payload)
        return response.json()['candidates'][0]['content']['parts'][0]['text']
    except Exception as e:
        return f"Error connecting to API. Make sure your key is correct! Details: {e}"

uploaded_file = st.file_uploader("Drop your study document here (PDF, DOCX, PPTX)", type=["pdf", "docx", "pptx"])

if uploaded_file and api_key:
    with st.spinner("Reading your file..."):
        file_text = extract_text(uploaded_file)
    st.success("File loaded successfully!")
    
    tab1, tab2, tab3 = st.tabs(["✨ ELI5 Summary", "🧠 Active Recall Quiz", "📋 Concept Map Table"])
    truncated_text = file_text[:12000] 

    with tab1:
        st.subheader("Plain English Breakdown")
        if st.button("Generate Summary"):
            with st.spinner("Stripping out the jargon..."):
                prompt = f"Analyze this study material and give me a comprehensive 'Explain Like I'm 5' summary. Break down complex jargon into simple, memorable analogies:\n\n{truncated_text}"
                st.write(ask_gemini(api_key, prompt))

    with tab2:
        st.subheader("Test Your Knowledge")
        if st.button("Generate Quiz"):
            with st.spinner("Creating flashcard questions..."):
                prompt = f"Based on this material, generate 5-10 challenging conceptual active recall questions. Under each question, provide the correct answer hidden clearly below it so I can test myself:\n\n{truncated_text}"
                st.write(ask_gemini(api_key, prompt))

    with tab3:
        st.subheader("Key Acronyms, Definitions & Formulas")
        if st.button("Generate Concept Map"):
            with st.spinner("Extracting cheat sheet details..."):
                prompt = f"Extract all critical terms, acronyms, definitions, and formulas from this text. Format them explicitly as a Markdown table with columns for 'Term/Concept' and 'Simple Definition/Formula':\n\n{truncated_text}"
                st.write(ask_gemini(api_key, prompt))