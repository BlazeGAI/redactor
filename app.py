import streamlit as st
import os
import logging
import json
from datetime import datetime
from pathlib import Path
import pandas as pd
import spacy
import re
from docx import Document
from pptx import Presentation
from rapidfuzz import fuzz

class Settings:
    def __init__(self):
        self.config_file = "redactor_settings.json"
        self.default_settings = {
            "redaction_text": "[REDACTED]",
            "preserve_case": True,
            "backup_files": True,
            "case_insensitive": True,
            "fuzzy_match": True,
            "fuzzy_threshold": 90,
            "recent_files": [],
            "max_recent_files": 5
        }
        self.load_settings()

    def load_settings(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    self.settings = json.load(f)
            else:
                self.settings = self.default_settings.copy()
        except Exception:
            self.settings = self.default_settings.copy()

    def save_settings(self):
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.settings, f)
        except Exception as e:
            st.error(f"Failed to save settings: {str(e)}")

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

    def setup_logging(self):
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"redactor_{datetime.now():%Y%m%d_%H%M%S}.log")

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )

    def load_names_from_csv(self, csv_file):
        try:
            import io
            df = pd.read_csv(io.StringIO(csv_file.getvalue().decode('utf-8')))
            self.custom_names = df["Name"].dropna().astype(str).tolist()
            logging.info(f"Loaded {len(self.custom_names)} custom names for redaction.")
        except Exception as e:
            st.error(f"Failed to load names from CSV: {e}")
            self.custom_names = []

    def redact_names(self, text):
        redacted_text = text
        redaction_text = self.settings.get("redaction_text")
        preserve_case = self.settings.get("preserve_case")
        case_insensitive = self.settings.get("case_insensitive")
        fuzzy_match = self.settings.get("fuzzy_match")
        threshold = self.settings.get("fuzzy_threshold")

        for name in self.custom_names:
            if case_insensitive:
                text_to_check = redacted_text.lower()
                name_to_check = name.lower()
            else:
                text_to_check = redacted_text
                name_to_check = name

            if fuzzy_match:
                words = redacted_text.split()
                for word in words:
                    if fuzz.ratio(word.lower(), name.lower()) >= threshold:
                        pattern = word if not case_insensitive else re.compile(re.escape(word), re.IGNORECASE)
                        replacement = self.apply_case(word, redaction_text) if preserve_case else redaction_text
                        redacted_text = re.sub(pattern, replacement, redacted_text)
            else:
                if name_to_check in text_to_check:
                    replacement = self.apply_case(name, redaction_text) if preserve_case else redaction_text
                    pattern = name if not case_insensitive else re.compile(re.escape(name), re.IGNORECASE)
                    redacted_text = re.sub(pattern, replacement, redacted_text)

        return redacted_text

    def apply_case(self, source, replacement):
        if source.isupper():
            return replacement.upper()
        elif source.istitle():
            return replacement.title()
        else:
            return replacement

    def process_word_document(self, input_file, output_path):
        try:
            if self.settings.get("backup_files"):
                backup_path = f"{input_file.name}.backup"
                with open(backup_path, "wb") as backup_file:
                    backup_file.write(input_file.getvalue())
                logging.info(f"Created backup: {backup_path}")

            doc = Document(input_file)

            for paragraph in doc.paragraphs:
                paragraph.text = self.redact_names(paragraph.text)

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell.text = self.redact_names(cell.text)

            doc.save(output_path)
            logging.info(f"Successfully processed: {input_file.name}")
            return True

        except Exception as e:
            logging.error(f"Error processing {input_file.name}: {str(e)}")
            st.error(f"Error processing document: {str(e)}")
            return False

    def process_powerpoint(self, input_file, output_path):
        try:
            if self.settings.get("backup_files"):
                backup_path = f"{input_file.name}.backup"
                with open(backup_path, "wb") as backup_file:
                    backup_file.write(input_file.getvalue())
                logging.info(f"Created backup: {backup_path}")

            prs = Presentation(input_file)

            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        shape.text = self.redact_names(shape.text)

            prs.save(output_path)
            logging.info(f"Successfully processed: {input_file.name}")
            return True

        except Exception as e:
            logging.error(f"Error processing {input_file.name}: {str(e)}")
            st.error(f"Error processing document: {str(e)}")
            return False

def main():
    st.title("📄 Document Name Redactor")

    redactor = DocumentRedactor()

    with st.sidebar:
        st.header("⚙️ Redaction Settings")
        redactor.settings.set("redaction_text", st.text_input("Redaction Text", value=redactor.settings.get("redaction_text")))
        redactor.settings.set("preserve_case", st.checkbox("Preserve Case", value=redactor.settings.get("preserve_case")))
        redactor.settings.set("case_insensitive", st.checkbox("Case-Insensitive Matching", value=redactor.settings.get("case_insensitive")))
        redactor.settings.set("fuzzy_match", st.checkbox("Enable Fuzzy Matching", value=redactor.settings.get("fuzzy_match")))
        redactor.settings.set("fuzzy_threshold", st.slider("Fuzzy Match Threshold", 0, 100, value=redactor.settings.get("fuzzy_threshold")))
        redactor.settings.set("backup_files", st.checkbox("Create Backup Before Processing", value=redactor.settings.get("backup_files")))

        st.header("🔍 Upload Name List (CSV)")
        name_csv = st.file_uploader("CSV file with names", type=["csv"])
        if name_csv:
            redactor.load_names_from_csv(name_csv)

    uploaded_file = st.file_uploader("Choose a document", type=['docx', 'pptx'])

    if uploaded_file is not None:
        file_extension = uploaded_file.name.split('.')[-1]
        base_name = uploaded_file.name.rsplit('.', 1)[0]
        output_filename = f"{base_name}-Redacted.{file_extension}"

        if st.button("Preview Redaction"):
            try:
                with st.expander("Redaction Preview"):
                    if file_extension == 'docx':
                        doc = Document(uploaded_file)
                        preview_text = "\n\n".join([p.text for p in doc.paragraphs])
                    else:
                        prs = Presentation(uploaded_file)
                        preview_text = "\n\n".join([
                            shape.text for slide in prs.slides
                            for shape in slide.shapes if hasattr(shape, "text")
                        ])

                    redactor.load_names_from_csv(name_csv) if name_csv else None
                    redacted_preview = redactor.redact_names(preview_text)

                    col1, col2 = st.columns(2)
                    with col1:
                        st.subheader("Original")
                        st.text_area("Original Text", preview_text, height=300)
                    with col2:
                        st.subheader("Redacted")
                        st.text_area("Redacted Text", redacted_preview, height=300)
            except Exception as e:
                st.error(f"Error creating preview: {e}")

        if st.button("Redact Document"):
            try:
                if file_extension == 'docx':
                    redactor.process_word_document(uploaded_file, output_filename)
                else:
                    redactor.process_powerpoint(uploaded_file, output_filename)

                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="Download Redacted Document",
                        data=file,
                        file_name=output_filename,
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                            if file_extension == 'docx'
                            else 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                    )
                os.remove(output_filename)
                st.success("Document successfully redacted!")
            except Exception as e:
                st.error(f"Error redacting document: {e}")

if __name__ == "__main__":
    main()
