import streamlit as st
import logging
import re
from io import BytesIO
from docx import Document
from pptx import Presentation
import fitz  # PyMuPDF

# spaCy setup: auto-download if missing
try:
    import spacy
    nlp = spacy.load("en_core_web_sm")
except:
    import spacy, spacy.cli
    spacy.cli.download("en_core_web_sm")
    nlp = spacy.load("en_core_web_sm")

class DocumentRedactor:
    def __init__(self):
        self.redaction_text = "[REDACTED]"
        logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

    def apply_case(self, source, repl):
        if source.isupper():
            return repl.upper()
        if source.istitle():
            return repl.title()
        return repl

    def extract_titlecase_ngrams(self, text, min_words=2, max_words=4):
        orgs = {"University","College","Institute","Department","School"}
        tokens = re.findall(r"\b[^\s]+\b", text)
        candidates = set()
        for n in range(min_words, max_words+1):
            for i in range(len(tokens)-n+1):
                seq = " ".join(tokens[i:i+n])
                words = seq.split()
                if all(re.match(r"^[A-Z][a-zA-Z'’-]+$", w) for w in words) \
                   and not any(w in orgs for w in words):
                    candidates.add(seq)
        return candidates

    def detect_human_names(self, text):
        names = set()
        # spaCy PERSON
        for ent in nlp(text).ents:
            if ent.label_ == "PERSON":
                names.add(ent.text)
        # honorifics
        honor = r"\b(Professor|Dr|Mr|Mrs|Ms|Miss)\.?\s+[A-Z][a-zA-Z'’-]+(?:\s+[A-Z][a-zA-Z'’-]+)*\b"
        for m in re.finditer(honor, text):
            names.add(m.group())
        # TitleCase n-grams fallback
        names |= self.extract_titlecase_ngrams(text)
        return list(names)

    def redact_text(self, text):
        names = self.detect_human_names(text)
        redacted = text
        for name in sorted(names, key=len, reverse=True):
            pattern = rf"\b{re.escape(name)}\b"
            redacted = re.sub(
                pattern,
                lambda m: self.apply_case(m.group(), self.redaction_text),
                redacted,
                flags=re.IGNORECASE
            )
        return redacted

    def process_word(self, data, out_path):
        doc = Document(BytesIO(data))
        count = 0
        # scan first ~10 paras to approximate first 2 pages
        paras = doc.paragraphs[:10]  
        for p in paras:
            new = self.redact_text(p.text)
            if new != p.text:
                for run in p.runs:
                    run.text = new
                count += 1
        # you can expand to tables, headers, footers similarly...
        doc.save(out_path)
        return count

    def process_ppt(self, data, out_path):
        prs = Presentation(BytesIO(data))
        count = 0
        slides = [0,1] if len(prs.slides)>1 else [0]
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
        pages = [0,1] if pdf.page_count>1 else [0]
        for i in pages:
            page = pdf[i]
            text = page.get_text()
            names = self.detect_human_names(text)
            for name in names:
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
        if ext=="docx":
            cnt = self.process_word(data, out)
        elif ext in ("ppt","pptx"):
            cnt = self.process_ppt(data, out)
        elif ext=="pdf":
            cnt = self.process_pdf(data, out)
        else:
            st.error(f"Unsupported: {ext}")
            return None, 0
        return out, cnt

def main():
    st.set_page_config(page_title="Name Redactor", layout="wide")
    st.title("Document Name Redactor (Auto Only)")
    redactor = DocumentRedactor()

    st.header("Upload Document")
    uploaded = st.file_uploader("Choose a file (DOCX, PPTX, PDF)", type=["docx","ppt","pptx","pdf"])
    if uploaded and st.button("Redact Document"):
        out_name, cnt = redactor.process(uploaded)
        if out_name:
            with open(out_name, "rb") as f:
                st.download_button("Download Redacted", data=f, file_name=out_name)
            st.success(f"Redacted {cnt} item(s)")

if __name__=="__main__":
    main()
