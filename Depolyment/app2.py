import re
import subprocess
import tempfile
import textract
import docx2txt
import pdfplumber
import pandas as pd
import streamlit as st
import os
import pickle
import platform

# Windows-only imports for handling .doc
if platform.system() == 'Windows':
    import win32com.client
    import pythoncom

import en_core_web_sm
nlp = en_core_web_sm.load()

import nltk
nltk.download('stopwords')
nltk.download('punkt')
nltk.download('wordnet')
nltk.download('omw-1.4')

from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
from nltk.tokenize import RegexpTokenizer

# TITLE
st.title('RESUME CLASSIFICATION')
st.markdown('<style>h1{color: Purple;}</style>', unsafe_allow_html=True)
st.subheader('Upload Resume to Extract Skills and Resume Text')

# LOAD SKILLS
skills_csv_path = os.path.join(os.getcwd(), 'skills.csv')
if os.path.exists(skills_csv_path):
    data = pd.read_csv(skills_csv_path)
    skills_master = list(data['Skill'].values) if 'Skill' in data.columns else list(data.columns)
else:
    st.error("'skills.csv' not found.")
    st.stop()

# LOAD MODEL & VECTORIZER
try:
    with open('modelDT2.pkl', 'rb') as m:
        model = pickle.load(m)
    with open('vector2.pkl', 'rb') as v:
        vectorizer = pickle.load(v)
except Exception:
    st.error("Model or vectorizer file not found or failed to load.")
    st.stop()

# FUNCTIONS

def repair_pdf(input_path):
    """Repair a malformed PDF using Ghostscript."""
    repaired = input_path.replace('.pdf', '_repaired.pdf')
    subprocess.run([
        'gs', '-o', repaired,
        '-sDEVICE=pdfwrite',
        input_path
    ], check=True)
    return repaired

def open_pdf_safely(fp):
    try:
        return pdfplumber.open(fp)
    except Exception as e:
        if 'No /Root object' in str(e):
            fixed = repair_pdf(fp)
            return pdfplumber.open(fixed)
        else:
            raise

def extract_skills(resume_text):
    doc = nlp(resume_text)
    tokens = [t.text for t in doc if not t.is_stop]
    noun_chunks = [c.text.lower().strip() for c in doc.noun_chunks]
    found = set(t.lower() for t in tokens + noun_chunks if t.lower() in skills_master)
    return [x.capitalize() for x in found]

def getText(uploaded_file):
    typ = uploaded_file.type
    fullText = ''

    if typ == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        # Handle .docx
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name
        fullText = docx2txt.process(tmp_path)
        os.unlink(tmp_path)

    elif typ == "application/msword":
        # Handle .doc using Microsoft Word COM Automation (Windows only)
        if platform.system() == 'Windows':
            pythoncom.CoInitialize()  # ✅ Initialize COM for Streamlit threads

            with tempfile.NamedTemporaryFile(delete=False, suffix='.doc') as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(tmp_path)
                fullText = doc.Content.Text
                doc.Close(False)
            finally:
                word.Quit()
                os.unlink(tmp_path)
                pythoncom.CoUninitialize()  # ✅ Clean up COM
        else:
            st.error(".doc support is only available on Windows with Microsoft Word.")
            return ""

    elif typ == "application/pdf":
        # Handle .pdf
        tmp = os.path.join(os.getcwd(), uploaded_file.name)
        with open(tmp, 'wb') as f:
            f.write(uploaded_file.read())
        with open_pdf_safely(tmp) as pdf:
            for p in pdf.pages:
                txt = p.extract_text()
                if txt:
                    fullText += txt
        os.remove(tmp)

    else:
        st.warning(f"Unsupported file type: {typ}")

    return fullText

def preprocess(text):
    text = re.sub(r'{html}|<.*?>|http\S+|\d+', '', str(text).lower())
    tokens = RegexpTokenizer(r'\w+').tokenize(text)
    tokens = [w for w in tokens if len(w) > 2 and w not in stopwords.words('english')]
    return " ".join(WordNetLemmatizer().lemmatize(w) for w in tokens)

# UPLOAD SECTION
upload_files = st.file_uploader('Upload Resume', type=['doc', 'docx', 'pdf'], accept_multiple_files=True)
output_df = pd.DataFrame(columns=['Uploaded File', 'Skills', 'Resume Text'])

if upload_files:
    for file in upload_files:
        text = getText(file)
        skills = extract_skills(text)
        output_df = pd.concat([output_df,
            pd.DataFrame({
                'Uploaded File': [file.name],
                'Skills': [", ".join(skills)],
                'Resume Text': [text]
            })
        ], ignore_index=True)
    st.dataframe(output_df)
else:
    st.info("Upload PDF/DOC/DOCX files to extract skills and text.")
