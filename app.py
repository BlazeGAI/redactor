import streamlit as st
import logging
import re
import unicodedata
import pandas as pd
from io import BytesIO
from docx import Document
from pptx import Presentation
import fitz  # PyMuPDF

# --- Redactor Class ---
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
        # generate simple variations
        variations = []
        for name in names:
            parts = name.split()
            variations += [p for p in parts if len(p) > 2]
        # preserve order and uniqueness
        seen = []
        for n in names + variations:
            if n not in seen:
                seen.append(n)
        self.custom_names = seen
        logging.info(f"Loaded {len(self.custom_names)} names for redaction")

    def normalize(self, text):
        text = unicodedata.normalize('NFKD', text)
        return ''.join(c if ord(c) >= 32 else ' ' for c in text)

    def apply_case(self, source, repl):
        if source.isupper():
            return repl.upper()
        if source.istitle():
            return repl.title()
        return repl

    def redact_names_text(self, text):
        flags = re.IGNORECASE
        redacted = text
        for name in sorted(self.custom_names, key=len, reverse=True):
            pattern = rf"\b{re.escape(name)}\b"
            redacted = re.sub(
                pattern,
                lambda m: self.apply_case(m.group(), self.redaction_text),
                redacted,
                flags=flags
            )
        return redacted

    def redact_roles_text(self, text):
        pattern = r"\b(Professor|Dr|Mr|Mrs|Ms|Miss)\.?\s+[A-Z][a-z]+\b"
        return re.sub(pattern, self.redaction_text, text, flags=re.IGNORECASE)

    def process_word(self, file_bytes, out_path):
        doc = Document(BytesIO(file_bytes))
        count = 0
        paras = doc.paragraphs
        targets = paras if self.custom_names else (paras[:3] + paras[-3:])
        for p in targets:
            original = p.text
            new = (self.redact_names_text(original) if self.custom_names else self.redact_roles_text(original))
            if new != original:
                for run in p.runs:
                    run.text = new
                count += 1
        # tables, headers, footers likewise
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text
                    new = (self.redact_names_text(text) if self.custom_names else self.redact_roles_text(text))
                    if new != text:
                        cell.text = new
                        count += 1
        for section in doc.sections:
            for part in (section.header, section.footer):
                for p in part.paragraphs:
                    text = p.text
                    new = (self.redact_names_text(text) if self.custom_names else self.redact_roles_text(text))
                    if new != text:
                        p.text = new
                        count += 1
        doc.save(out_path)
        return count

    def process_ppt(self, file_bytes, out_path):
        prs = Presentation(BytesIO(file_bytes))
        count = 0
        slide_indices = range(len(prs.slides)) if self.custom_names else [0, len(prs.slides) - 1]
        for i in slide_indices:
            slide = prs.slides[i]
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame'):
                    for p in shape.text_frame.paragraphs:
                        original = p.text
                        new = (self.redact_names_text(original) if self.custom_names else self.redact_roles_text(original))
                        if new != original:
                            for run in p.runs:
                                run.text = new
                            count += 1
        prs.save(out_path)
        return count

    def process_pdf(self, file_bytes, out_path):
        pdf = fitz.open(stream=file_bytes, filetype='pdf')
        count = 0
        total_pages = pdf.page_count
        for page in pdf:
            to_redact = []
            if self.custom_names:
                for name in self.custom_names:
                    to_redact += page.search_for(name)
            else:
                if page.number in (0, total_pages - 1):
                    matches = page.search_for(r"Professor", flags=fitz.TEXT_DEHYPHENATE)
                    to_redact += matches
                    # repeat for other honorifics as needed
            for inst in to_redact:
                page.add_redact_annot(inst, fill=(0, 0, 0))
                count += 1
        pdf.apply_redactions()
        pdf.save(out_path)
        return count

    def process(self, uploaded):
        ext = uploaded.name.rsplit('.', 1)[-1].lower()
        base = uploaded.name.rsplit('.', 1)[0]
        out_name = f"{base}-Redacted.{ext}"
        data = uploaded.read()
        if ext == 'docx':
            count = self.process_word(data, out_name)
        elif ext in ('ppt', 'pptx'):
            count = self.process_ppt(data, out_name)
        elif ext == 'pdf':
            count = self.process_pdf(data, out_name)
        else:
            st.error(f"Unsupported format: {ext}")
            return None, 0
        return out_name, count

# --- Streamlit UI ---

def main():
    st.set_page_config(page_title="Name Redactor", layout="wide")
    st.title("Document Name Redactor")

    redactor = DocumentRedactor()

    st.sidebar.header("Upload Names (CSV)")
    name_file = st.sidebar.file_uploader("CSV with 'Name' column", type=['csv'])
    if name_file:
        redactor.load_names_from_csv(name_file)
        if redactor.custom_names:
            st.sidebar.success(f"Loaded {len(redactor.custom_names)} names")

    st.header("Upload Document")
    uploaded = st.file_uploader("Choose a file (DOCX, PPTX, PDF)", type=['docx', 'ppt', 'pptx', 'pdf'])

    if uploaded:
        if not redactor.custom_names:
            st.warning("No name list provided — scanning title and references for honorifics.")
        if st.button("Redact Document"):
            out_name, count = redactor.process(uploaded)
            if out_name:
                with open(out_name, 'rb') as f:
                    st.download_button(
                        label="Download Redacted",
                        data=f,
                        file_name=out_name
                    )
                st.success(f"Redacted {count} item(s)")

if __name__ == '__main__':
    main()
