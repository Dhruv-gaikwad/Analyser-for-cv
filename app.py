import streamlit as st
import tempfile, os
from analyzer_core import extract_pdf_text   # use your own function names

st.title("AI Resume / CV Analyzer")

files = st.file_uploader("Upload CVs", accept_multiple_files=True)

if files and st.button("Analyze"):
    for f in files:
        # save the upload to a temp file, because your code expects a file path
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f.name)[1]) as tmp:
            tmp.write(f.read())
            path = tmp.name
        text = extract_pdf_text(path)
        st.subheader(f.name)
        st.write(text[:1000])