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
    # Load the English model; install with:
    #   python -m spacy download en_core_web_sm
    nlp = spacy.load("en_core_web_sm")
except Exception:
    nlp = None

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
        # also include individual parts
        parts = [p for name in names for p in name.split() if len(p) > 2]
        # preserve order, remove duplicates
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

    def detect_human_names(self, text):
        names = set()
        # spaCy NER
        if nlp:
            for ent in nlp(text).ents:
                if ent.label_ == "PERSON":
                    names.add(ent.text)
        # honorifics
        honor_pattern = (
            r"\b(Professor|Dr|Mr|Mrs|Ms|Miss)\.?\s+[A-Z][a-zA-Z'’-]+"
            r"(?:\s+[A-Z][a-zA-Z'’-]+)*\b"
        )
        for m in re.finditer(honor_pattern, text):
            names.add(m.group())
        # TitleCase sliding-window 2–4 words
        tokens = re.findall(r"\b[^\s]+\b", text)
        orgs = {"University","College","Institute","Department","School"}
        for n in range(2, 5):
            for i in range(len(tokens)-n+1):
                seq = " ".join(tokens[i:i+n])
                words = seq.split()
                if all(re.match(r"^[A-Z][a-zA-Z'’-]+$", w) for w in words):
                    if not any(w in orgs for w in words):
                        names.add(seq)
        return list(names)

    def redact_text(self, text, names_list):
        redacted = text
        flags = re.IGNORECASE
        # sort by length to avoid substring conflicts
        for name in sorted(names_list, key=len, reverse=True):
            pattern = rf"\b{re.escape(name)}\b"
            redacted = re.sub(
                pattern,
                lambda m: self.apply_case(m.group(), self.redaction_text),
                redacted,
                flags=flags
            )
        return redacted

    def process_word(self, data, out_path):
        doc = Document(BytesIO(data))
        count = 0
        paras = doc.paragraphs[:10] if not self.custom_names else doc.paragraphs
        for p in paras:
            original = p.text
            detected = set(self.detect_human_names(original))
            combined = list(detected.union(self.custom_names))
            new = self.redact_text(original, combined)
            if new != original:
                for run in p.runs:
                    run.text = new
                count += 1
        # tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text
                    detected = set(self.detect_human_names(text))
                    combined = list(detected.union(self.custom_names))
                    new = self.redact_text(text, combined)
                    if new != text:
                        cell.text = new
                        count += 1
        # headers & footers
        for section in doc.sections:
            for part in (section.header, section.footer):
                for p in part.paragraphs:
                    text = p.text
                    detected = set(self.detect_human_names(text))
                    combined = list(detected.union(self.custom_names))
                    new = self.redact_text(text, combined)
                    if new != text:
                        p.text = new
                        count += 1
        doc.save(out_path)
        return count

    def process_ppt(self, data, out_path):
        prs = Presentation(BytesIO(data))
        count = 0
        slides = range(len(prs.slides)) if self.custom_names else ([0,1] if len(prs.slides)>1 else [0])
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
        pages = range(pdf.page_count) if self.custom_names else ([0,1] if pdf.page_count>1 else [0])
        for i in pages:
            page = pdf[i]
            text = page.get_text()
            detected = set(self.detect_human_names(text))
            combined = list(detected.union(self.custom_names))
            rects = []
            for name in combined:
                rects.extend(page.search_for(name, flags=fitz.TEXT_DEHYPHENATE))
            for r in rects:
                page.add_redact_annot(r, fill=(0,0,0))
                count += 1
        pdf.apply_redactions()
        pdf.save(out_path)
        return count

    def process(self, uploaded):
        ext = uploaded.name.rsplit(".",1)[-1].lower()
        base = uploaded.name.rsplit(".",1)[0]
        out = f"{base}-Redacted.{ext}"
        data = uploaded.read()
        if ext == "docx":
            cnt = self.process_word(data, out)
        elif ext in ("ppt","pptx"):
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
    uploaded = st.file_uploader(
        "Choose a file (DOCX, PPTX, PDF)",
        type=["docx", "ppt", "pptx", "pdf"]
    )
    if uploaded:
        if not redactor.custom_names:
            st.warning("No name list—scanning first two pages/slides for human names.")
        if st.button("Redact Document"):
            out_name, cnt = redactor.process(uploaded)
            if out_name:
                with open(out_name, "rb") as f:
                    st.download_button(
                        label="Download Redacted", data=f, file_name=out_name
                    )
                st.success(f"Redacted {cnt} item(s)")

if __name__ == "__main__":
    main()
