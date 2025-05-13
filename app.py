import streamlit as st
import logging
import re
import pandas as pd
from io import BytesIO
from docx import Document
from pptx import Presentation
import fitz  # PyMuPDF

# spaCy for robust multi-word name detection
try:
    import spacy
    nlp = spacy.load("en_core_web_sm")
except Exception:
    nlp = None

class DocumentRedactor:
    def __init__(self):
        self.custom_names = []
        self.redaction_text = "[REDACTED]"
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def load_names_from_csv(self, uploaded_csv):
        df = pd.read_csv(uploaded_csv)
        if 'Name' not in df.columns:
            st.error("CSV must contain a 'Name' column.")
            return
        names = df['Name'].dropna().astype(str).str.strip().tolist()
        variations = []
        for name in names:
            parts = name.split()
            variations += [p for p in parts if len(p) > 2]
        seen = []
        for n in names + variations:
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

    def extract_titlecase_ngrams(self, text, min_words=2, max_words=4):
        org_exclusions = ["University", "College", "Institute", "Department", "School"]
        tokens = re.findall(r"\b[^\s]+\b", text)
        seqs = set()
        for n in range(min_words, max_words+1):
            for i in range(len(tokens) - n + 1):
                seq = " ".join(tokens[i:i+n])
                words = seq.split()
                if all(re.match(r"^[A-Z][a-zA-Z'’-]+$", w) for w in words):
                    if not any(w.lower() in (org.lower() for org in org_exclusions) for w in words):
                        seqs.add(seq)
        return seqs

    def detect_human_names(self, text):
        """
        Combine spaCy NER, honorific regex, and TitleCase n-grams for name detection.
        """
        names = set()
        # 1) spaCy NER
        if nlp:
            doc = nlp(text)
            for ent in doc.ents:
                if ent.label_ == "PERSON":
                    names.add(ent.text)
        # 2) honorific-based names (multi-word)
        honorific_pattern = r"\b(Professor|Dr|Mr|Mrs|Ms|Miss)\.?\s+[A-Z][a-zA-Z'’-]+(?:\s+[A-Z][a-zA-Z'’-]+)*\b"
        for m in re.finditer(honorific_pattern, text):
            names.add(m.group())
        # 3) TitleCase sliding-window (fallback if spaCy unavailable)
        if not nlp:
            for seq in self.extract_titlecase_ngrams(text):
                names.add(seq)
        return list(names)

    def redact_text(self, text, names_list):
        flags = re.IGNORECASE
        redacted = text
        for name in sorted(names_list, key=len, reverse=True):
            pattern = rf"\b{re.escape(name)}\b"
            redacted = re.sub(
                pattern,
                lambda m: self.apply_case(m.group(), self.redaction_text),
                redacted,
                flags=flags
            )
        return redacted

    def process_word(self, file_bytes, out_path):
        doc = Document(BytesIO(file_bytes))
        count = 0
        paras = doc.paragraphs[:10] if not self.custom_names else doc.paragraphs
        for p in paras:
            original = p.text
            names = self.custom_names or self.detect_human_names(original)
            new = self.redact_text(original, names)
            if new != original:
                for run in p.runs:
                    run.text = new
                count += 1
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text
                    names = self.custom_names or self.detect_human_names(text)
                    new = self.redact_text(text, names)
                    if new != text:
                        cell.text = new
                        count += 1
        for section in doc.sections:
            for part in (section.header, section.footer):
                for p in part.paragraphs:
                    text = p.text
                    names = self.custom_names or self.detect_human_names(text)
                    new = self.redact_text(text, names)
                    if new != text:
                        p.text = new
                        count += 1
        doc.save(out_path)
        return count

    def process_ppt(self, file_bytes, out_path):
        prs = Presentation(BytesIO(file_bytes))
        count = 0
        indices = range(len(prs.slides)) if self.custom_names else ([0,1] if len(prs.slides)>1 else [0])
        for i in indices:
            slide = prs.slides[i]
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame'):
                    for p in shape.text_frame.paragraphs:
                        original = p.text
                        names = self.custom_names or self.detect_human_names(original)
                        new = self.redact_text(original, names)
                        if new != original:
                            for run in p.runs:
                                run.text = new
                            count += 1
        prs.save(out_path)
        return count

    def process_pdf(self, file_bytes, out_path):
        pdf = fitz.open(stream=file_bytes, filetype='pdf')
        count = 0
        indices = range(pdf.page_count) if self.custom_names else ([0,1] if pdf.page_count>1 else [0])
        for i in indices:
            page = pdf[i]
            text = page.get_text()
            names = self.custom_names or self.detect_human_names(text)
            insts = []
            for name in names:
                insts += page.search_for(name, flags=fitz.TEXT_DEHYPHENATE)
            for inst in insts:
                page.add_redact_annot(inst, fill=(0,0,0))
                count += 1
        pdf.apply_redactions()
        pdf.save(out_path)
        return count

    def process(self, uploaded):
        ext = uploaded.name.rsplit('.',1)[-1].lower()
        base = uploaded.name.rsplit('.',1)[0]
        out = f"{base}-Redacted.{ext}"
        data = uploaded.read()
        if ext=='docx':
            count = self.process_word(data, out)
        elif ext in ('ppt','pptx'):
            count = self.process_ppt(data, out)
        elif ext=='pdf':
            count = self.process_pdf(data, out)
        else:
            st.error(f"Unsupported format: {ext}")
            return None,0
        return out, count


def main():
    st.set_page_config(page_title="Name Redactor", layout="wide")
    st.title("Document Name Redactor")
    redactor = DocumentRedactor()

    st.sidebar.header("Upload Names (CSV)")
    name_file = st.sidebar.file_uploader("CSV with 'Name' column", type=['csv'])
    if name_file:
        redactor.load_names_from_csv(name_file)
        if redactor.custom_names:
            st.sidebar.success(f"Loaded {len(redactor.custom_names)} names for redaction")

    st.header("Upload Document")
    uploaded = st.file_uploader("Choose a file (DOCX, PPTX, PDF)", type=['docx','ppt','pptx','pdf'])
    if uploaded:
        if not redactor.custom_names:
            st.warning("No name list — scanning first two pages/slides for multi-word human names.")
        if st.button("Redact Document"):
            out_name, count = redactor.process(uploaded)
            if out_name:
                with open(out_name,'rb') as f:
                    st.download_button(label="Download Redacted", data=f, file_name=out_name)
                st.success(f"Redacted {count} item(s)")

if __name__=='__main__':
    main()
