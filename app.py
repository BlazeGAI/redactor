import streamlit as st
import logging
import re
import unicodedata
import pandas as pd
from io import BytesIO
from docx import Document
from pptx import Presentation
import fitz  # PyMuPDF

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

    def redact_fallback_text(self, text):
        # Honorific-based titles
        patterns = [
            r"\b(Professor|Dr|Mr|Mrs|Ms|Miss)\.?\s+[A-Z][a-zA-Z]+\b",
        ]
        # Two-word TitleCase human names, excluding common organization words
        org_exclusions = ["University", "College", "Institute", "Department", "School"]
        excl = "|".join(org_exclusions)
        titlecase_pattern = rf"\b(?!{excl})[A-Z][a-zA-Z]+\s+(?!{excl})[A-Z][a-zA-Z]+\b"
        patterns.append(titlecase_pattern)

        for pat in patterns:
            text = re.sub(pat, self.redaction_text, text, flags=re.IGNORECASE)
        return text

    def process_word(self, file_bytes, out_path):
        doc = Document(BytesIO(file_bytes))
        count = 0
        paras = doc.paragraphs
        if self.custom_names:
            targets = paras
        else:
            # Fallback: approximate first two pages via first 10 paragraphs
            targets = paras[:10]
        for p in targets:
            original = p.text
            new = (self.redact_names_text(original) if self.custom_names else self.redact_fallback_text(original))
            if new != original:
                for run in p.runs:
                    run.text = new
                count += 1
        # Tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text
                    new = (self.redact_names_text(text) if self.custom_names else self.redact_fallback_text(text))
                    if new != text:
                        cell.text = new
                        count += 1
        # Headers & footers
        for section in doc.sections:
            for part in (section.header, section.footer):
                for p in part.paragraphs:
                    text = p.text
                    new = (self.redact_names_text(text) if self.custom_names else self.redact_fallback_text(text))
                    if new != text:
                        p.text = new
                        count += 1
        doc.save(out_path)
        return count

    def process_ppt(self, file_bytes, out_path):
        prs = Presentation(BytesIO(file_bytes))
        count = 0
        if self.custom_names:
            slide_indices = range(len(prs.slides))
        else:
            # Fallback: first two slides
            slide_indices = [0, 1] if len(prs.slides) > 1 else [0]
        for i in slide_indices:
            slide = prs.slides[i]
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame'):
                    for p in shape.text_frame.paragraphs:
                        original = p.text
                        new = (self.redact_names_text(original) if self.custom_names else self.redact_fallback_text(original))
                        if new != original:
                            for run in p.runs:
                                run.text = new
                            count += 1
        prs.save(out_path)
        return count

    def process_pdf(self, file_bytes, out_path):
        pdf = fitz.open(stream=file_bytes, filetype='pdf')
        count = 0
        if self.custom_names:
            page_indices = range(pdf.page_count)
        else:
            # Fallback: first two pages
            page_indices = [0, 1] if pdf.page_count > 1 else [0]
        for i in page_indices:
            page = pdf[i]
            to_redact = []
            if self.custom_names:
                for name in self.custom_names:
                    to_redact += page.search_for(name)
            else:
                # Honorifics and TitleCase fallback
                patterns = [r"Professor", r"Dr\.", r"Mr\.", r"Mrs\.", r"Ms\.", r"Miss"]
                for pat in patterns:
                    to_redact += page.search_for(pat, flags=fitz.TEXT_DEHYPHENATE)
                # Exclude organizations in TitleCase
                for inst in page.search_for(r"[A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+", flags=fitz.TEXT_DEHYPHENATE):
                    text = page.get_textbox(inst)
                    if not re.match(rf"^(University|College|Institute|Department|School)$", text.split()[-1], re.IGNORECASE):
                        to_redact.append(inst)
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
    uploaded = st.file_uploader("Choose a file (DOCX, PPTX, PDF)", type=['docx', 'ppt', 'pptx', 'pdf'])

    if uploaded:
        if not redactor.custom_names:
            st.warning("No name list — scanning first two pages/slides for honorifics and human names (excluding orgs).")
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
