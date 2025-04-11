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
            "fuzzy_threshold": 75,  # Lowered for better matching
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
        self.debug_mode = False  # Set to True for verbose logging

    def setup_logging(self):
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"redactor_{datetime.now():%Y%m%d_%H%M%S}.log")

        logging.basicConfig(
            level=logging.DEBUG,  # Set to DEBUG for more detailed logs
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )

    def normalize_text(self, text):
        """Normalize text by removing control chars and normalizing Unicode"""
        if not text:
            return ""
        # Normalize Unicode characters
        text = unicodedata.normalize('NFKD', text)
        # Replace control characters with spaces
        text = ''.join(char if ord(char) >= 32 else ' ' for char in text)
        return text

    def load_names_from_csv(self, csv_file):
        try:
            import io
            df = pd.read_csv(io.StringIO(csv_file.getvalue().decode('utf-8')))
            
            if "Name" in df.columns:
                # Get names from the "Name" column and clean them
                self.custom_names = df["Name"].dropna().astype(str).tolist()
                # Remove empty strings and whitespace-only strings
                self.custom_names = [name.strip() for name in self.custom_names if name and name.strip()]
                
                # Create variations of multi-word names (first name, last name, etc.)
                name_variations = []
                for name in self.custom_names:
                    parts = name.split()
                    # Add individual parts if they're long enough (to avoid short words)
                    name_variations.extend([part for part in parts if len(part) > 2])
                
                # Add variations to the list but keep original names first
                self.custom_names = self.custom_names + [v for v in name_variations if v not in self.custom_names]
                
                logging.info(f"Loaded {len(self.custom_names)} names and variations for redaction")
                if self.custom_names and self.debug_mode:
                    logging.debug(f"Names list: {self.custom_names}")
            else:
                st.error("CSV file must contain a column named 'Name'")
                self.custom_names = []
        except Exception as e:
            st.error(f"Failed to load names from CSV: {e}")
            logging.error(f"Failed to load names from CSV: {e}")
            self.custom_names = []

    def apply_case(self, source, replacement):
        """Preserve case pattern in the replacement text"""
        if source.isupper():
            return replacement.upper()
        elif source.istitle():
            return replacement.title()
        else:
            return replacement

    def redact_names(self, text):
        """Redact names from the given text"""
        if not text or not self.custom_names:
            return text
            
        # Normalize the input text
        orig_text = text
        text = self.normalize_text(text)
        redacted_text = text
        
        redaction_text = self.settings.get("redaction_text")
        preserve_case = self.settings.get("preserve_case")
        case_insensitive = self.settings.get("case_insensitive")
        fuzzy_match = self.settings.get("fuzzy_match")
        threshold = self.settings.get("fuzzy_threshold")
        
        if self.debug_mode:
            logging.debug(f"Original text: '{text[:50]}...'")  # Log start of text
            
        # Sort names by length (descending) for better matching
        sorted_names = sorted(self.custom_names, key=len, reverse=True)
        
        # First pass: Exact matching
        flags = re.IGNORECASE if case_insensitive else 0
        for name in sorted_names:
            if not name or len(name) < 3:  # Skip very short names
                continue
                
            replacement = self.apply_case(name, redaction_text) if preserve_case else redaction_text
            
            # Process exact matches with word boundaries
            pattern = rf'\b{re.escape(name)}\b'
            if re.search(pattern, redacted_text, flags=flags):
                logging.info(f"Exact match found for '{name}'")
                redacted_text = re.sub(pattern, replacement, redacted_text, flags=flags)
        
        # Second pass: Fuzzy matching if enabled
        if fuzzy_match:
            # Split text into words and small phrases for fuzzy matching
            words = re.findall(r'\b\w+(?:\s+\w+){0,2}\b', redacted_text)
            for word in words:
                if len(word) < 3:  # Skip very short words
                    continue
                    
                for name in sorted_names:
                    if len(name) < 3:  # Skip very short names
                        continue
                        
                    name_lower = name.lower() if case_insensitive else name
                    word_lower = word.lower() if case_insensitive else word
                    
                    # Calculate similarity
                    similarity = fuzz.ratio(word_lower, name_lower)
                    
                    # Check if similarity exceeds threshold
                    if similarity >= threshold:
                        logging.info(f"Fuzzy match: '{word}' matches '{name}' with similarity {similarity}")
                        replacement = self.apply_case(word, redaction_text) if preserve_case else redaction_text
                        # Use word boundaries for more precise replacement
                        pattern = rf'\b{re.escape(word)}\b'
                        redacted_text = re.sub(pattern, replacement, redacted_text, flags=flags)
        
        # If we made changes but they're not reflected in the output, try direct replacement
        if redacted_text != text and redacted_text == orig_text:
            logging.warning("Normalized text changes not reflected in original. Using direct replacement.")
            for name in sorted_names:
                if name in orig_text:
                    replacement = self.apply_case(name, redaction_text) if preserve_case else redaction_text
                    orig_text = orig_text.replace(name, replacement)
            return orig_text
            
        return redacted_text

    def process_word_document(self, input_file, output_path):
        """Process a Word document for redaction"""
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
                if original_text.strip():  # Skip empty paragraphs
                    redacted_text = self.redact_names(original_text)
                    if original_text != redacted_text:
                        paragraph.text = redacted_text
                        redaction_count += 1

            # Process tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        original_text = cell.text
                        if original_text.strip():  # Skip empty cells
                            redacted_text = self.redact_names(original_text)
                            if original_text != redacted_text:
                                cell.text = redacted_text
                                redaction_count += 1

            # Process headers and footers
            for section in doc.sections:
                for header in section.header.paragraphs:
                    original_text = header.text
                    if original_text.strip():
                        redacted_text = self.redact_names(original_text)
                        if original_text != redacted_text:
                            header.text = redacted_text
                            redaction_count += 1
                
                for footer in section.footer.paragraphs:
                    original_text = footer.text
                    if original_text.strip():
                        redacted_text = self.redact_names(original_text)
                        if original_text != redacted_text:
                            footer.text = redacted_text
                            redaction_count += 1

            # Save the document
            doc.save(output_path)
            logging.info(f"Successfully processed Word document: {input_file.name} with {redaction_count} redactions")
            return True, redaction_count

        except Exception as e:
            logging.error(f"Error processing Word document {input_file.name}: {str(e)}")
            st.error(f"Error processing document: {str(e)}")
            return False, 0

    def process_powerpoint(self, input_file, output_path):
        """Process a PowerPoint presentation for redaction"""
        try:
            if self.settings.get("backup_files"):
                backup_path = f"{input_file.name}.backup"
                with open(backup_path, "wb") as backup_file:
                    backup_file.write(input_file.getvalue())
                logging.info(f"Created backup: {backup_path}")

            prs = Presentation(input_file)
            redaction_count = 0

            for slide_num, slide in enumerate(prs.slides):
                logging.debug(f"Processing slide {slide_num+1}")
                
                # Process all shapes that might contain text
                for shape in slide.shapes:
                    # Different ways text can appear in PowerPoint shapes
                    if hasattr(shape, "text_frame") and hasattr(shape.text_frame, "text"):
                        # Debug log the text
                        if self.debug_mode:
                            logging.debug(f"Shape text: '{shape.text_frame.text[:50]}...'")
                        
                        # Process each paragraph in the text frame
                        for paragraph in shape.text_frame.paragraphs:
                            original_text = paragraph.text
                            if original_text.strip():  # Skip empty paragraphs
                                redacted_text = self.redact_names(original_text)
                                
                                if original_text != redacted_text:
                                    # PowerPoint requires special handling for text replacement
                                    # We need to clear all runs and create a new one
                                    for run in paragraph.runs:
                                        run.text = ""  # Clear existing text
                                    
                                    # If we have no runs, create one
                                    if not paragraph.runs:
                                        paragraph.add_run().text = redacted_text
                                    else:
                                        # Use the first run for the new text
                                        paragraph.runs[0].text = redacted_text
                                    
                                    redaction_count += 1
                    
                    # Handle table cells in PowerPoint
                    if hasattr(shape, "table"):
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if hasattr(cell, "text_frame"):
                                    for paragraph in cell.text_frame.paragraphs:
                                        original_text = paragraph.text
                                        if original_text.strip():
                                            redacted_text = self.redact_names(original_text)
                                            if original_text != redacted_text:
                                                # Clear all runs and create new one
                                                for run in paragraph.runs:
                                                    run.text = ""
                                                
                                                # Use first run or create new one
                                                if not paragraph.runs:
                                                    paragraph.add_run().text = redacted_text
                                                else:
                                                    paragraph.runs[0].text = redacted_text
                                                    
                                                redaction_count += 1

            # Save the presentation
            prs.save(output_path)
            logging.info(f"Successfully processed PowerPoint: {input_file.name} with {redaction_count} redactions")
            return True, redaction_count

        except Exception as e:
            logging.error(f"Error processing PowerPoint {input_file.name}: {str(e)}")
            st.error(f"Error processing presentation: {str(e)}")
            return False, 0

    def process_document(self, input_file, is_preview=False):
        """Process document or generate preview based on file type"""
        file_extension = input_file.name.split('.')[-1].lower()
        
        try:
            if file_extension == 'docx':
                if is_preview:
                    # For Word document preview
                    doc = Document(input_file)
                    preview_text = "\n\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                    return preview_text, self.redact_names(preview_text)
                else:
                    # For full Word document processing
                    base_name = input_file.name.rsplit('.', 1)[0]
                    output_filename = f"{base_name}-Redacted.{file_extension}"
                    return self.process_word_document(input_file, output_filename)
            
            elif file_extension in ['ppt', 'pptx']:
                if is_preview:
                    # For PowerPoint preview
                    prs = Presentation(input_file)
                    
                    # Collect text from all slides
                    preview_texts = []
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if hasattr(shape, "text_frame") and hasattr(shape.text_frame, "text"):
                                text = shape.text_frame.text.strip()
                                if text:
                                    preview_texts.append(text)
                    
                    preview_text = "\n\n".join(preview_texts)
                    return preview_text, self.redact_names(preview_text)
                else:
                    # For full PowerPoint processing
                    base_name = input_file.name.rsplit('.', 1)[0]
                    output_filename = f"{base_name}-Redacted.{file_extension}"
                    return self.process_powerpoint(input_file, output_filename)
            else:
                raise ValueError(f"Unsupported file format: {file_extension}")
                
        except Exception as e:
            logging.error(f"Error processing {input_file.name}: {str(e)}")
            if is_preview:
                return f"Error generating preview: {str(e)}", ""
            else:
                return False, 0

def main():
    st.set_page_config(
        page_title="Document Name Redactor",
        page_icon="📄",
        layout="wide",
    )
    
    st.title("📄 Document Name Redactor")
    st.markdown("Redact names and sensitive information from your documents")

    redactor = DocumentRedactor()

    # Create a sidebar for settings
    with st.sidebar:
        st.header("⚙️ Redaction Settings")
        
        # Redaction text
        redactor.settings.set("redaction_text", st.text_input(
            "Redaction Text", 
            value=redactor.settings.get("redaction_text"),
            help="Text that will replace redacted names"
        ))
        
        # Case settings
        col1, col2 = st.columns(2)
        with col1:
            redactor.settings.set("preserve_case", st.checkbox(
                "Preserve Case", 
                value=redactor.settings.get("preserve_case"),
                help="Keep the same case pattern (UPPER, Title, lower) when redacting"
            ))
        
        with col2:
            redactor.settings.set("case_insensitive", st.checkbox(
                "Case-Insensitive", 
                value=redactor.settings.get("case_insensitive"),
                help="Match names regardless of capitalization"
            ))
        
        # Fuzzy matching settings
        redactor.settings.set("fuzzy_match", st.checkbox(
            "Enable Fuzzy Matching", 
            value=redactor.settings.get("fuzzy_match"),
            help="Match names even with slight spelling variations"
        ))
        
        if redactor.settings.get("fuzzy_match"):
            redactor.settings.set("fuzzy_threshold", st.slider(
                "Fuzzy Match Threshold", 
                50, 100, 
                value=redactor.settings.get("fuzzy_threshold"),
                help="Higher values require closer matches (recommended: 75-85)"
            ))
        
        # Backup setting
        redactor.settings.set("backup_files", st.checkbox(
            "Create Backup Files", 
            value=redactor.settings.get("backup_files"),
            help="Save original files before redacting"
        ))

        # Debug mode toggle (hidden in UI but available for development)
        # Uncomment this to enable debug mode in the UI
        # redactor.debug_mode = st.checkbox("Debug Mode", value=redactor.debug_mode, help="Enable detailed logging")

        st.header("🔍 Upload Name List (CSV)")
        st.markdown("Upload a CSV file with a column named 'Name' containing the names to redact.")
        name_csv = st.file_uploader("CSV file with names", type=["csv"])
        if name_csv:
            redactor.load_names_from_csv(name_csv)
            if redactor.custom_names:
                st.success(f"✅ Loaded {len(redactor.custom_names)} names for redaction")
                with st.expander("View loaded names"):
                    st.write(redactor.custom_names[:20])  # Show first 20 names
                    if len(redactor.custom_names) > 20:
                        st.write(f"...and {len(redactor.custom_names) - 20} more")
            else:
                st.error("❌ No valid names found in the CSV file")

    # Main content area
    st.header("📂 Upload Document")
    
    uploaded_file = st.file_uploader(
        "Choose a document to redact", 
        type=['docx', 'pptx'], 
        help="Supported formats: Word documents (.docx) and PowerPoint presentations (.pptx)"
    )

    if uploaded_file is not None:
        st.write(f"Selected file: **{uploaded_file.name}**")
        
        # Ensure names are loaded before attempting redaction
        if not redactor.custom_names and name_csv:
            redactor.load_names_from_csv(name_csv)

        # Check if names are available
        if not redactor.custom_names:
            st.warning("⚠️ Please upload a CSV file with names to redact first.")
        else:
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🔍 Preview Redaction", use_container_width=True):
                    try:
                        with st.spinner("Generating preview..."):
                            original_preview, redacted_preview = redactor.process_document(uploaded_file, is_preview=True)
                            
                            if original_preview and original_preview != "Error":
                                with st.expander("Redaction Preview", expanded=True):
                                    preview_col1, preview_col2 = st.columns(2)
                                    
                                    with preview_col1:
                                        st.subheader("Original")
                                        st.text_area("Original Text", original_preview, height=300)
                                    
                                    with preview_col2:
                                        st.subheader("Redacted")
                                        st.text_area("Redacted Text", redacted_preview, height=300)
                                    
                                    # Show if any changes were made
                                    if original_preview == redacted_preview:
                                        st.warning("⚠️ No redactions were made in the preview. Check your name list and settings.")
                                    else:
                                        st.success("✅ Preview shows redactions! You can now proceed to redact the full document.")
                            else:
                                st.error("❌ Error generating preview or no text content found.")
                    except Exception as e:
                        st.error(f"❌ Error creating preview: {e}")
                        logging.error(f"Preview error: {e}")
            
            with col2:
                if st.button("🔒 Redact Document", use_container_width=True):
                    try:
                        with st.spinner("Redacting document..."):
                            success, redaction_count = redactor.process_document(uploaded_file, is_preview=False)
                            
                            if success:
                                file_extension = uploaded_file.name.split('.')[-1].lower()
                                base_name = uploaded_file.name.rsplit('.', 1)[0]
                                output_filename = f"{base_name}-Redacted.{file_extension}"
                                
                                with open(output_filename, "rb") as file:
                                    mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' \
                                        if file_extension == 'docx' \
                                        else 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                                    
                                    st.download_button(
                                        label="⬇️ Download Redacted Document",
                                        data=file,
                                        file_name=output_filename,
                                        mime=mime_type,
                                        use_container_width=True
                                    )
                                
                                # Remove the temporary file
                                os.remove(output_filename)
                                
                                if redaction_count > 0:
                                    st.success(f"✅ Document successfully redacted with {redaction_count} redactions!")
                                else:
                                    st.warning("⚠️ Document processed but no redactions were made. Check your name list and settings.")
                    except Exception as e:
                        st.error(f"❌ Error redacting document: {e}")
                        logging.error(f"Redaction error: {e}")

    # Add help section at the bottom
    with st.expander("ℹ️ Help & Instructions"):
        st.markdown("""
        ### How to use this tool:
        
        1. **Upload a CSV file** with names to redact in the sidebar. The CSV must have a column named "Name".
        2. **Configure redaction settings** in the sidebar (redaction text, case sensitivity, etc.).
        3. **Upload your document** (Word or PowerPoint).
        4. **Preview the redaction** to see what will be changed.
        5. **Redact the document** when you're satisfied with the preview.
        6. **Download the redacted document**.
        
        ### Troubleshooting:
        
        - If names aren't being redacted, try enabling fuzzy matching and lowering the threshold.
        - Make sure your CSV file is properly formatted with a "Name" column.
        - For PowerPoint files, complex formatting might affect redaction results.
        - Check the logs folder for detailed information if you encounter issues.
        """)

if __name__ == "__main__":
    main()
