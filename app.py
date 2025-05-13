import sys
import subprocess
import importlib
import streamlit as st
import logging
import re
import pandas as pd
from io import BytesIO
from docx import Document
from pptx import Presentation
import fitz  # PyMuPDF

# Ensure the spaCy English model is installed in user site-packages
@st.cache_resource
def ensure_spacy_model():
    import spacy
    try:
        return importlib.import_module("en_core_web_sm")
    except ModuleNotFoundError:
        # Determine spaCy version and construct matching model URL
        version = spacy.__version__
        model_url = (
            f"https://github.com/explosion/spacy-models/releases/"
            f"download/en_core_web_sm-{version}/"
            f"en_core_web_sm-{version}-py3-none-any.whl"
        )
        # Install into user site to avoid venv permissions issues
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--user", model_url
        ])
        return importlib.import_module("en_core_web_sm")

# Download/load the model once per session
ensure_spacy_model()

import spacy
nlp = spacy.load("en_core_web_sm")

class DocumentRedactor:
    def __init__(self):
        self.custom_names = []
        self.redaction_text = "[REDACTED]"
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )

    def load_names_from_csv(self, uploaded_csv):
        df = pd.read_csv(uploaded_csv)
        if "Name" not in df.columns:
            st.error("CSV must contain a 'Name' column.")
            return
        names = df["Name"].dropna().astype(str).str.strip().tolist()
        parts = [p for name in names for p in name.split() if len(p) > 2]
        seen = []
        for n in names + parts:
            if n not in seen:
                seen.append(n)
        self.custom_names = seen
        logging.info(f"Loaded {len(self.custom_names)} names for redaction")

    def apply_case(self, source, repl):
        if source.isupper():
            return repl.upper()
        if source.istitle():
            return repl.title()
        return repl

    def extract_titlecase_ngrams(self, text, min_w=2, max_w=4):
        exclude = {"University", "College", "Institute", "Department", "School"}
        tokens = re.findall(r"\b[^\s]+\b", text)
        results = set()
        for n in range(min_w, max_w + 1):
            for i in range(len(tokens) - n + 1):
                seq = " ".join(tokens[i : i + n])
                words = seq.split()
                if all(re.match(r"^[A-Z][a-zA-Z'’-]+$", w) for w in words) and not any(w in exclude for w in words):
                    results.add(seq)
        return results

    def detect_human_names(self, text):
        names = set()
        # spaCy NER
        for ent in nlp(text).ents:
            if ent.label_ == "PERSON":
                names.add(ent.text)
        # honorifics
        honor = (
            r"\b(Professor|Dr|Mr|Mrs|Ms|Miss)\.?\s+[A-Z][a-zA-Z'’-]+"
            r"(?:\s+[A-Z][a-zA-Z'’-]+)*\b"
        )
        for m in re.finditer(honor, text):
            names.add(m.group())
        # TitleCase fallback
        names |= self.extract_titlecase_ngrams(text)
        return list(names)

    def redact_text(self, text, names_list):
        redacted = text
        for name in sorted(names_list, key=len, reverse=True):
            pat = rf"\b{re.escape(name)}\b"
            redacted = re.sub(
                pat,
                lambda m: self.apply_case(m.group(), self.redaction_text),
                redacted,
                flags=re.IGNORECASE
            )
        return redacted

    def process_word(self, data, out_path):
        doc = Document(BytesIO(data))
        count = 0
        paras = doc.paragraphs if self.custom_names else doc.paragraphs[:10]
        for p in paras:
            original = p.text
            detected = set(self.detect_human_names(original))
            combined = list(detected.union(self.custom_names))
            new = self.redact_text(original, combined)
            if new != original:
                for run in p.runs:
                    run.text = new
                count += 1
        # tables, headers, footers omitted for brevity…
        doc.save(out_path)
        return count

    def process_ppt(self, data, out_path):
        prs = Presentation(BytesIO(data))
        count = 0
        slides = range(len(prs.slides)) if self.custom_names else ([0, 1] if len(prs.slides) > 1 else [0])
        for i in slides:
            slide = prs.slides[i]
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for p in shape.text_frame.paragraphs:
                        original = p.text
                        detected = set(self.detect_human_names(original))
                        combined = list(detected.union(self.custom_names))
                        new = self.redact_text(original, combined)
                        if new != original:
                            for run in p.runs:
                                run.text = new
                            count += 1
        prs.save(out_path)
        return count

    def process_pdf(self, data, out_path):
        pdf = fitz.open(stream=data, filetype="pdf")
        count = 0
        pages = range(pdf.page_count) if self.custom_names else ([0, 1] if pdf.page_count > 1 else [0])
        for i in pages:
            page = pdf[i]
            text = page.get_text()
            detected = set(self.detect_human_names(text))
            combined = list(detected.union(self.custom_names))
            for name in combined:
                for r in page.search_for(name, flags=fitz.TEXT_DEHYPHENATE):
                    page.add_redact_annot(r, fill=(0, 0, 0))
                    count += 1
        pdf.apply_redactions()
        pdf.save(out_path)
        return count

    def process(self, uploaded):
        ext = uploaded.name.rsplit(".", 1)[-1].lower()
        base = uploaded.name.rsplit(".", 1)[0]
        out = f"{base}-Redacted.{ext}"
        data = uploaded.read()
        if ext == "docx":
            cnt = self.process_word(data, out)
        elif ext in ("ppt", "pptx"):
            cnt = self.process_ppt(data, out)
        elif ext == "pdf":
            cnt = self.process_pdf(data, out)
        else:
            st.error(f"Unsupported format: {ext}")
            return None, 0
        return out, cnt

def main():
    st.set_page_config(page_title="Name Redactor", layout="wide")
    st.title("Document Name Redactor")
    redactor = DocumentRedactor()

    st.sidebar.header("Upload Names (CSV)")
    name_csv = st.sidebar.file_uploader("CSV with 'Name' column", type=["csv"])
    if name_csv:
        redactor.load_names_from_csv(name_csv)
        if redactor.custom_names:
            st.sidebar.success(f"Loaded {len(redactor.custom_names)} names")

    st.header("Upload Document")
    uploaded = st.file_uploader("Choose a file (DOCX, PPTX, PDF)", type=["docx", "ppt", "pptx", "pdf"])
    if uploaded and st.button("Redact Document"):
        out_name, cnt = redactor.process(uploaded)
        if out_name:
            with open(out_name, "rb") as f:
                st.download_button(label="Download Redacted", data=f, file_name=out_name)
            st.success(f"Redacted {cnt} item(s)")

if __name__ == "__main__":
    main()
