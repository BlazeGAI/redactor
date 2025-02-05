import streamlit as st
import os
import logging
import json
from datetime import datetime
from pathlib import Path

import spacy
from docx import Document
from pptx import Presentation

class Settings:
    def __init__(self):
        self.config_file = "redactor_settings.json"
        self.default_settings = {
            "redaction_text": "[REDACTED]",
            "entity_types": ["PERSON"],
            "preserve_case": True,
            "backup_files": True,
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
        # Setup logging
        self.setup_logging()
        
        # Initialize settings
        self.settings = Settings()
        
        # Load NLP model
        self.load_nlp_model()
    
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
    
    def load_nlp_model(self):
        try:
            self.nlp = spacy.load("en_core_web_sm")
        except Exception as e:
            st.error(f"Error loading NLP model: {e}")
            st.stop()
    
    def redact_names(self, text):
        """Enhanced redaction with multiple entity types and case preservation"""
        doc = self.nlp(text)
        redacted_text = text
        entities = []
        
        # Collect all entities to redact
        selected_types = self.settings.get("entity_types")
        for ent in doc.ents:
            if ent.label_ in selected_types:
                entities.append(ent.text)
        
        # Sort by length to avoid partial replacements
        entities.sort(key=len, reverse=True)
        
        # Replace entities with redaction text
        redaction_text = self.settings.get("redaction_text")
        for entity in entities:
            if self.settings.get("preserve_case"):
                if entity.isupper():
                    replacement = redaction_text.upper()
                elif entity.istitle():
                    replacement = redaction_text.title()
                else:
                    replacement = redaction_text
            else:
                replacement = redaction_text
            
            redacted_text = redacted_text.replace(entity, replacement)
        
        return redacted_text
    
    def process_word_document(self, input_file, output_path):
        """Process Word document with redaction"""
        try:
            # Create backup if enabled
            if self.settings.get("backup_files"):
                backup_path = f"{input_file.name}.backup"
                with open(backup_path, "wb") as backup_file:
                    backup_file.write(input_file.getvalue())
                logging.info(f"Created backup: {backup_path}")
            
            # Read the document
            doc = Document(input_file)
            
            # Process paragraphs
            for paragraph in doc.paragraphs:
                original_text = paragraph.text
                redacted_text = self.redact_names(original_text)
                if original_text != redacted_text:
                    logging.info(f"Redacted paragraph: {original_text} -> {redacted_text}")
                paragraph.text = redacted_text
            
            # Process tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        original_text = cell.text
                        redacted_text = self.redact_names(original_text)
                        if original_text != redacted_text:
                            logging.info(f"Redacted table cell: {original_text} -> {redacted_text}")
                        cell.text = redacted_text
            
            # Save the redacted document
            doc.save(output_path)
            logging.info(f"Successfully processed: {input_file.name}")
            return True
        
        except Exception as e:
            logging.error(f"Error processing {input_file.name}: {str(e)}")
            st.error(f"Error processing document: {str(e)}")
            return False
    
    def process_powerpoint(self, input_file, output_path):
        """Process PowerPoint document with redaction"""
        try:
            # Create backup if enabled
            if self.settings.get("backup_files"):
                backup_path = f"{input_file.name}.backup"
                with open(backup_path, "wb") as backup_file:
                    backup_file.write(input_file.getvalue())
                logging.info(f"Created backup: {backup_path}")
            
            # Read the presentation
            prs = Presentation(input_file)
            
            # Process slides
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        original_text = shape.text
                        redacted_text = self.redact_names(original_text)
                        if original_text != redacted_text:
                            logging.info(f"Redacted slide shape: {original_text} -> {redacted_text}")
                        shape.text = redacted_text
            
            # Save the redacted presentation
            prs.save(output_path)
            logging.info(f"Successfully processed: {input_file.name}")
            return True
        
        except Exception as e:
            logging.error(f"Error processing {input_file.name}: {str(e)}")
            st.error(f"Error processing document: {str(e)}")
            return False

def main():
    st.title("📄 Document Name Redactor")
    
    # Initialize redactor
    redactor = DocumentRedactor()
    
    # Sidebar for settings
    with st.sidebar:
        st.header("⚙️ Redaction Settings")
        
        # Redaction text
        redaction_text = st.text_input(
            "Redaction Text", 
            value=redactor.settings.get("redaction_text")
        )
        redactor.settings.set("redaction_text", redaction_text)
        
        # Entity types
        entity_types = st.multiselect(
            "Entity Types to Redact", 
            options=["PERSON", "ORG", "GPE", "DATE"],
            default=redactor.settings.get("entity_types")
        )
        redactor.settings.set("entity_types", entity_types)
        
        # Case preservation
        preserve_case = st.checkbox(
            "Preserve Case", 
            value=redactor.settings.get("preserve_case")
        )
        redactor.settings.set("preserve_case", preserve_case)
        
        # Backup files
        backup_files = st.checkbox(
            "Create Backup Before Processing", 
            value=redactor.settings.get("backup_files")
        )
        redactor.settings.set("backup_files", backup_files)
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a document", 
        type=['docx', 'pptx']
    )
    
    if uploaded_file is not None:
        # Prepare output
        file_extension = uploaded_file.name.split('.')[-1]
        output_filename = uploaded_file.name.replace(f'.{file_extension}', '_redacted.{file_extension}')
        
        # Preview option
        if st.button("Preview Redaction"):
            try:
                with st.expander("Redaction Preview"):
                    if file_extension == 'docx':
                        doc = Document(uploaded_file)
                        preview_text = "\n\n".join([p.text for p in doc.paragraphs])
                    else:  # pptx
                        prs = Presentation(uploaded_file)
                        preview_text = "\n\n".join([
                            shape.text for slide in prs.slides 
                            for shape in slide.shapes if hasattr(shape, "text")
                        ])
                    
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
        
        # Process and download
        if st.button("Redact Document"):
            try:
                # Process based on file type
                if file_extension == 'docx':
                    redactor.process_word_document(uploaded_file, output_filename)
                else:  # pptx
                    redactor.process_powerpoint(uploaded_file, output_filename)
                
                # Provide download
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="Download Redacted Document",
                        data=file,
                        file_name=output_filename,
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
                             if file_extension == 'docx' 
                             else 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                    )
                
                # Clean up
                os.remove(output_filename)
                
                st.success("Document successfully redacted!")
            
            except Exception as e:
                st.error(f"Error redacting document: {e}")

if __name__ == "__main__":
    main()
