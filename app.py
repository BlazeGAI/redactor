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
            "fuzzy_threshold": 80,  # Lowered threshold for better matching
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
            if "Name" in df.columns:
                self.custom_names = df["Name"].dropna().astype(str).tolist()
                # Remove any empty strings or whitespace-only strings
                self.custom_names = [name for name in self.custom_names if name and name.strip()]
                logging.info(f"Loaded {len(self.custom_names)} custom names for redaction.")
                if self.custom_names:
                    logging.info(f"First few names: {self.custom_names[:5]}")
                else:
                    logging.warning("No valid names found in the CSV file.")
            else:
                st.error("CSV file must contain a column named 'Name'")
                self.custom_names = []
        except Exception as e:
            st.error(f"Failed to load names from CSV: {e}")
            logging.error(f"Failed to load names from CSV: {e}")
            self.custom_names = []

    def redact_names(self, text):
        if not text or not self.custom_names:
            return text

        redacted_text = text
        redaction_text = self.settings.get("redaction_text")
        preserve_case = self.settings.get("preserve_case")
        case_insensitive = self.settings.get("case_insensitive")
        fuzzy_match = self.settings.get("fuzzy_match")
        threshold = self.settings.get("fuzzy_threshold")
        
        logging.info(f"Redacting text with {len(self.custom_names)} names. Fuzzy match: {fuzzy_match}, Threshold: {threshold}")
        
        # Sort names by length (descending) to handle longer names first
        sorted_names = sorted(self.custom_names, key=len, reverse=True)
        
        for name in sorted_names:
            if not name:
                continue
                
            # Prepare for case handling
            flags = re.IGNORECASE if case_insensitive else 0
            replacement = self.apply_case(name, redaction_text) if preserve_case else redaction_text
            
            if fuzzy_match:
                # Split text into chunks to check for fuzzy matches
                # This implementation finds words that might be similar to the name
                words = re.findall(r'\b\w+(?:\s+\w+){0,2}\b', redacted_text)
                for word in words:
                    name_lower = name.lower() if case_insensitive else name
                    word_lower = word.lower() if case_insensitive else word
                    
                    similarity = fuzz.ratio(word_lower, name_lower)
                    if similarity >= threshold:
                        logging.info(f"Fuzzy match: '{word}' matches '{name}' with similarity {similarity}")
                        # Use word boundaries for more precise replacement
                        pattern = rf'\b{re.escape(word)}\b'
                        redacted_text = re.sub(pattern, replacement, redacted_text, flags=flags)
            else:
                # Exact matching with word boundaries
                pattern = rf'\b{re.escape(name)}\b'
                if re.search(pattern, redacted_text, flags=flags):
                    logging.info(f"Exact match found for '{name}'")
                    redacted_text = re.sub(pattern, replacement, redacted_text, flags=flags)

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
            redaction_count = 0

            for paragraph in doc.paragraphs:
                original_text = paragraph.text
                redacted_text = self.redact_names(original_text)
                if original_text != redacted_text:
                    paragraph.text = redacted_text
                    redaction_count += 1

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        original_text = cell.text
                        redacted_text = self.redact_names(original_text)
                        if original_text != redacted_text:
                            cell.text = redacted_text
                            redaction_count += 1

            # Save the document
            doc.save(output_path)
            logging.info(f"Successfully processed: {input_file.name} with {redaction_count} redactions")
            return True, redaction_count

        except Exception as e:
            logging.error(f"Error processing {input_file.name}: {str(e)}")
            st.error(f"Error processing document: {str(e)}")
            return False, 0

    def process_powerpoint(self, input_file, output_path):
        try:
            if self.settings.get("backup_files"):
                backup_path = f"{input_file.name}.backup"
                with open(backup_path, "wb") as backup_file:
                    backup_file.write(input_file.getvalue())
                logging.info(f"Created backup: {backup_path}")

            prs = Presentation(input_file)
            redaction_count = 0

            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text_frame") and hasattr(shape.text_frame, "text"):
                        original_text = shape.text_frame.text
                        redacted_text = self.redact_names(original_text)
                        if original_text != redacted_text:
                            # For PowerPoint, we need to update each paragraph in the text frame
                            for i, paragraph in enumerate(shape.text_frame.paragraphs):
                                if i < len(shape.text_frame.paragraphs):
                                    original_para_text = paragraph.text
                                    redacted_para_text = self.redact_names(original_para_text)
                                    if original_para_text != redacted_para_text:
                                        paragraph.text = redacted_para_text
                                        redaction_count += 1

            prs.save(output_path)
            logging.info(f"Successfully processed: {input_file.name} with {redaction_count} redactions")
            return True, redaction_count

        except Exception as e:
            logging.error(f"Error processing {input_file.name}: {str(e)}")
            st.error(f"Error processing document: {str(e)}")
            return False, 0

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
        st.markdown("Upload a CSV file with a column named 'Name' containing the names to redact.")
        name_csv = st.file_uploader("CSV file with names", type=["csv"])
        if name_csv:
            redactor.load_names_from_csv(name_csv)
            st.info(f"Loaded {len(redactor.custom_names)} names for redaction")

    st.header("📂 Upload Document")
    uploaded_file = st.file_uploader("Choose a document", type=['docx', 'pptx'])

    if uploaded_file is not None:
        file_extension = uploaded_file.name.split('.')[-1]
        base_name = uploaded_file.name.rsplit('.', 1)[0]
        output_filename = f"{base_name}-Redacted.{file_extension}"

        # Ensure names are loaded before attempting redaction
        if not redactor.custom_names and name_csv:
            redactor.load_names_from_csv(name_csv)

        # Check if names are available
        if not redactor.custom_names:
            st.warning("Please upload a CSV file with names to redact first.")
        else:
            if st.button("Preview Redaction"):
                try:
                    with st.spinner("Generating preview..."):
                        if file_extension == 'docx':
                            doc = Document(uploaded_file)
                            preview_text = "\n\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                        else:
                            prs = Presentation(uploaded_file)
                            preview_text = "\n\n".join([
                                shape.text_frame.text for slide in prs.slides
                                for shape in slide.shapes if hasattr(shape, "text_frame") and hasattr(shape.text_frame, "text") 
                                and shape.text_frame.text.strip()
                            ])

                        redacted_preview = redactor.redact_names(preview_text)
                        
                        # Display preview only if there's content to show
                        if preview_text:
                            with st.expander("Redaction Preview", expanded=True):
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.subheader("Original")
                                    st.text_area("Original Text", preview_text, height=300)
                                with col2:
                                    st.subheader("Redacted")
                                    st.text_area("Redacted Text", redacted_preview, height=300)
                                
                                # Show if any changes were made
                                if preview_text == redacted_preview:
                                    st.warning("No redactions were made in the preview. Check your name list and settings.")
                                else:
                                    st.success("Preview shows redactions! You can now proceed to redact the full document.")
                        else:
                            st.warning("No text content found in the document for preview.")
                except Exception as e:
                    st.error(f"Error creating preview: {e}")
                    logging.error(f"Preview error: {e}")

            if st.button("Redact Document"):
                try:
                    with st.spinner("Redacting document..."):
                        success = False
                        redaction_count = 0
                        
                        if file_extension == 'docx':
                            success, redaction_count = redactor.process_word_document(uploaded_file, output_filename)
                        else:
                            success, redaction_count = redactor.process_powerpoint(uploaded_file, output_filename)
                        
                        if success:
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
                            
                            if redaction_count > 0:
                                st.success(f"Document successfully redacted with {redaction_count} redactions!")
                            else:
                                st.warning("Document processed but no redactions were made. Check your name list and settings.")
                except Exception as e:
                    st.error(f"Error redacting document: {e}")
                    logging.error(f"Redaction error: {e}")

if __name__ == "__main__":
    main()
