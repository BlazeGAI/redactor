import streamlit as st
import logging
import re
import pandas as pd
from io import BytesIO
from docx import Document
from pptx import Presentation
import fitz  # PyMuPDF

# on import, try to load the model; if missing, download it in code
try:
    import spacy
    nlp = spacy.load("en_core_web_sm")
except (ImportError, OSError):
    import spacy, spacy.cli
    spacy.cli.download("en_core_web_sm")
    nlp = spacy.load("en_core_web_sm")

class DocumentRedactor:
    def __init__(self):
        self.redaction_text = "[REDACTED]"
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )

    def apply_case(self, source, repl):
        if source.isupper():
            return repl.upper()
        if source.istitle():
            return repl.title()
        return repl

    def extract_titlecase_ngrams(self, text, min_w=2, max_w=4):
        exclude = {"University","College","Institute","Department","School"}
        tokens = re.findall(r"\b[^\s]+\b", text)
        results = set()
        for n in range(min_w, max_w+1):
            for i in range(len(tokens)-n+1):
                seq = " ".join(tokens[i:i+n])
                words = seq.split()
                if all(re.match(r"^[A-Z][a-zA-Z'’-]+$", w) for w in words):
                    if not any(w in exclude for w in words):
                        results.add(seq)
        return results

    def detect_human_names(self, text):
        names = set()
        for ent in nlp(text).ents:
            if ent.label_ == "PERSON":
                names.add(ent.text)
        honor = (
            r"\b(Professor|Dr|Mr|Mrs|Ms|Miss)\.?\s+[A-Z][a-zA-Z'’-]+"
            r"(?:\s+[A-Z][a-zA-Z'’-]+)*\b"
        )
        for m in re.finditer(honor, text):
            names.add(m.group())
        names |= self.extract_titlecase_ngrams(text)
        return list(names)

    def redact_text(self, text):
        names = self.detect_human_names(text)
        redacted = text
        for name in sorted(names, key=len, reverse=True):
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
        # scan first ~10 paragraphs for title‐page names
        for p in doc.paragraphs[:10]:
            new = self.redact_text(p.text)
            if new != p.text:
                for run in p.runs:
                    run.text = new
                count += 1
        # add tables, headers/footers if needed…
        doc.save(out_path)
        return count

    def process_ppt(self, data, out_path):
        prs = Presentation(BytesIO(data))
        count = 0
        slides = [0,1] if len(prs.slides) > 1 else [0]
        for i in slides:
            for shape in prs.slides[i].shapes:
                if hasattr(shape, "text_frame"):
                    for p in shape.text_frame.paragraphs:
                        new = self.redact_text(p.text)
                        if new != p.text:
                            for run in p.runs:
                                run.text = new
                            count += 1
        prs.save(out_path)
        return count

    def process_pdf(self, data, out_path):
        pdf = fitz.open(stream=data, filetype="pdf")
        count = 0
        pages = [0,1] if pdf.page_count > 1 else [0]
        for i in pages:
            page = pdf[i]
            text = page.get_text()
            for name in self.detect_human_names(text):
                for inst in page.search_for(name, flags=fitz.TEXT_DEHYPHENATE):
                    page.add_redact_annot(inst, fill=(0,0,0))
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

    st.header("Upload Document")
    uploaded = st.file_uploader("Choose DOCX, PPTX or PDF", type=["docx","ppt","pptx","pdf"])
    if uploaded and st.button("Redact Document"):
        out_name, cnt = redactor.process(uploaded)
        if out_name:
            with open(out_name, "rb") as f:
                st.download_button("Download Redacted", data=f, file_name=out_name)
            st.success(f"Redacted {cnt} item(s)")

if __name__ == "__main__":
    main()
