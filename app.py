import streamlit as st
import os
import logging
import json
from datetime import datetime
from pathlib import Path
import pandas as pd
import re
from docx import Document
from pptx import Presentation
import unicodedata

class Settings:
    def __init__(self):
        self.config_file = "redactor_settings.json"
        self.default_settings = {
            "redaction_text": "[REDACTED]",
            "preserve_case": True,
            "backup_files": True,
            "case_insensitive": True,
            "fuzzy_match": True,
            "fuzzy_threshold": 75,
            "recent_files": [],
            "max_recent_files": 5
        }
        self.load_settings()

    def load_settings(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    self.settings = json.load(f)
            except Exception:
                self.settings = self.default_settings.copy()
        else:
            self.settings = self.default_settings.copy()

    def save_settings(self):
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.settings, f)
        except Exception as e:
            st.error(f"Failed to save settings: {e}")

    def get(self, key):
        return self.settings.get(key, self.default_settings.get(key))

    def set(self, key, value):
        self.settings[key] = value
        self.save_settings()

class DocumentRedactor:
    def __init__(self):
        self.setup_logging()
        self.settings = Settings()
        self.custom_names = []
        self.debug_mode = False

    def setup_logging(self):
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"redactor_{datetime.now():%Y%m%d_%H%M%S}.log")
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )

    def normalize_text(self, text):
        if not text:
            return ""
        text = unicodedata.normalize('NFKD', text)
        return ''.join(c if ord(c) >= 32 else ' ' for c in text)

    def load_names_from_csv(self, csv_file):
        try:
            import io
            df = pd.read_csv(io.StringIO(csv_file.getvalue().decode('utf-8')))
            if "Name" not in df.columns:
                st.error("CSV must have a 'Name' column.")
                return
            names = df["Name"].dropna().astype(str).str.strip()
            names = [n for n in names if len(n) > 2]
            variations = []
            for name in names:
                parts = name.split()
                variations += [p for p in parts if len(p) > 2]
            self.custom_names = list(dict.fromkeys(names + variations))
            logging.info(f"Loaded {len(self.custom_names)} names for redaction")
        except Exception as e:
            st.error(f"Failed to load names: {e}")
            logging.error(f"CSV load error: {e}")

    def apply_case(self, source, replacement):
        if source.isupper():
            return replacement.upper()
        if source.istitle():
            return replacement.title()
        return replacement

    def redact_names(self, text):
        if not text or not self.custom_names:
            return text
        redacted = text
        flags = re.IGNORECASE if self.settings.get("case_insensitive") else 0
        sorted_names = sorted(self.custom_names, key=len, reverse=True)
        repl = self.settings.get("redaction_text")
        for name in sorted_names:
            pattern = rf"\b{re.escape(name)}\b"
            redacted = re.sub(pattern, lambda m: self.apply_case(m.group(), repl), redacted, flags=flags)
        return redacted

    def redact_roles_only(self, text):
        repl = self.settings.get("redaction_text")
        patterns = [r"\b(Professor|Dr|Mr|Mrs|Ms|Miss)\.?\s+[A-Z][a-z]+\b"]
        for pat in patterns:
            text = re.sub(pat, repl, text, flags=re.IGNORECASE)
        return text

    def process_word_document(self, input_file, output_path):
        doc = Document(input_file)
        count = 0
        # If no custom names, apply roles-only to first and last 3 paras
        if not self.custom_names:
            targets = list(doc.paragraphs[:3]) + list(doc.paragraphs[-3:])
            for p in targets:
                new = self.redact_roles_only(p.text)
                if new != p.text:
                    p.text = new
                    count += 1
        else:
            for para in doc.paragraphs:
                new = self.redact_names(para.text)
                if new != para.text:
                    para.text = new
                    count += 1
        # Tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    txt = cell.text
                    new = self.custom_names and self.redact_names(txt) or self.redact_roles_only(txt)
                    if new != txt:
                        cell.text = new
                        count += 1
        # Headers & footers
        for section in doc.sections:
            for part in (section.header, section.footer):
                for p in part.paragraphs:
                    txt = p.text
                    new = self.custom_names and self.redact_names(txt) or self.redact_roles_only(txt)
                    if new != txt:
                        p.text = new
                        count += 1
        doc.save(output_path)
        return True, count

    def process_powerpoint(self, input_file, output_path):
        prs = Presentation(input_file)
        count = 0
        slides = prs.slides
        target_indices = ([0, -1] if not self.custom_names else range(len(slides)))
        for idx in target_indices:
            slide = slides[idx]
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame'):
                    for para in shape.text_frame.paragraphs:
                        txt = para.text
                        if not txt.strip():
                            continue
                        new = (self.custom_names and self.redact_names(txt)
                               or self.redact_roles_only(txt))
                        if new != txt:
                            for run in para.runs:
                                run.text = new  # preserve style
                            count += 1
        prs.save(output_path)
        return True, count

    def process_document(self, input_file, is_preview=False):
        ext = input_file.name.rsplit('.', 1)[-1].lower()
        base = input_file.name.rsplit('.', 1)[0]
        out = f"{base}-Redacted.{ext}"
        if ext == 'docx':
            return self.process_word_document(input_file, out)
        if ext in ('ppt', 'pptx'):
            return self.process_powerpoint(input_file, out)
        st.error(f"Unsupported format: {ext}")
        return False, 0


def main():
    st.set_page_config(page_title="Document Name Redactor", layout="wide")
    st.title("Document Name Redactor")
    redactor = DocumentRedactor()

    with st.sidebar:
        st.header("Settings")
        redactor.settings.set("redaction_text", st.text_input("Redaction text", redactor.settings.get("redaction_text")))
        c1, c2 = st.columns(2)
        with c1:
            redactor.settings.set("preserve_case", st.checkbox("Preserve case", redactor.settings.get("preserve_case")))
        with c2:
            redactor.settings.set("case_insensitive", st.checkbox("Case-insensitive", redactor.settings.get("case_insensitive")))
        redactor.settings.set("fuzzy_match", st.checkbox("Fuzzy matching", redactor.settings.get("fuzzy_match")))
        if redactor.settings.get("fuzzy_match"):
            redactor.settings.set("fuzzy_threshold", st.slider("Fuzzy threshold", 50, 100, redactor.settings.get("fuzzy_threshold")))
        redactor.settings.set("backup_files", st.checkbox("Backup files", redactor.settings.get("backup_files")))
        st.header("Upload Name List (CSV)")
        name_csv = st.file_uploader("CSV with 'Name' column", type=['csv'])
        if name_csv:
            redactor.load_names_from_csv(name_csv)

    st.header("Upload Document")
    uploaded = st.file_uploader("Choose Word or PowerPoint file", type=['docx', 'ppt', 'pptx'])
    if uploaded:
        if not redactor.custom_names:
            st.warning("No name list. The app will scan title and references for roles.")
            if st.button("Scan for author/instructor roles"):
                success, count = redactor.process_document(uploaded)
                if success:
                    with open(f"{uploaded.name.rsplit('.',1)[0]}-Redacted.{uploaded.name.rsplit('.',1)[1]}", 'rb') as f:
                        st.download_button("Download redacted", f, file_name=f.name)
                    st.success(f"Applied fallback redaction ({count} items)")
        else:
            if st.button("Preview redaction"):
                preview, redacted = redactor.process_document(uploaded, is_preview=True)
                st.subheader("Preview")
                st.text_area("Original", preview)
                st.text_area("Redacted", redacted)
            if st.button("Redact document"):
                success, count = redactor.process_document(uploaded)
                if success:
                    out_name = f"{uploaded.name.rsplit('.',1)[0]}-Redacted.{uploaded.name.rsplit('.',1)[1]}"
                    with open(out_name, 'rb') as f:
                        st.download_button("Download redacted", f, file_name=out_name)
                    st.success(f"Redacted {count} items")

if __name__ == '__main__':
    main()
