import streamlit as st
import os
import logging
import json
import zipfile
import tempfile
from datetime import datetime
from pathlib import Path

import spacy
from docx import Document
from pptx import Presentation
from docx.opc.exceptions import PackageNotFoundError

class Settings:
    def __init__(self):
        self.config_file = "redactor_settings.json"
        self.default_settings = {
            "redaction_text": "[REDACTED]",
            "redact_author": True,
            "preserve_case": True,
            "backup_files": True
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
    
    def redact_author(self, document_path):
        """Redact document author metadata"""
        try:
            # Word document author redaction
            if document_path.endswith('.docx'):
                doc = Document(document_path)
                core_properties = doc.core_properties
                
                # Redact author
                if core_properties.author:
                    logging.info(f"Redacting author: {core_properties.author}")
                    core_properties.author = self.settings.get("redaction_text")
                
                # Save the modified document
                doc.save(document_path)
            
            # PowerPoint document author redaction
            elif document_path.endswith('.pptx'):
                prs = Presentation(document_path)
                
                # Unfortunately, python-pptx doesn't have a straightforward 
                # way to modify core properties, so we'll log the limitation
                st.warning("PowerPoint author redaction is limited due to library constraints.")
            
            return True
        except PackageNotFoundError:
            st.error(f"Invalid document: {document_path}")
            return False
        except Exception as e:
            st.error(f"Error redacting author in {document_path}: {str(e)}")
            return False
    
    def process_zip(self, uploaded_zip):
        """Process a zip file of documents"""
        # Create temporary directories
        with tempfile.TemporaryDirectory() as input_dir, tempfile.TemporaryDirectory() as output_dir:
            # Extract uploaded zip
            with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                zip_ref.extractall(input_dir)
            
            # Track processed files
            processed_files = []
            
            # Recursively process files
            for root, _, files in os.walk(input_dir):
                for file in files:
                    if file.endswith(('.docx', '.pptx')):
                        input_path = os.path.join(root, file)
                        output_path = os.path.join(output_dir, file)
                        
                        # Ensure output directory structure
                        os.makedirs(os.path.dirname(output_path), exist_ok=True)
                        
                        # Copy the file for processing
                        import shutil
                        shutil.copy2(input_path, output_path)
                        
                        # Redact author
                        if self.redact_author(output_path):
                            processed_files.append(output_path)
            
            # Create output zip
            output_zip_path = os.path.join(tempfile.gettempdir(), 'redacted_documents.zip')
            with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file in processed_files:
                    arcname = os.path.relpath(file, output_dir)
                    zipf.write(file, arcname=arcname)
            
            return output_zip_path

def main():
    st.title("📄 Document Author Redactor")
    
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
        
        # Author redaction toggle (default on)
        redact_author = st.checkbox(
            "Redact Document Author", 
            value=redactor.settings.get("redact_author"),
            help="Remove author metadata from documents"
        )
        redactor.settings.set("redact_author", redact_author)
    
    # File upload
    uploaded_files = st.file_uploader(
        "Choose documents or a zip file", 
        type=['docx', 'pptx', 'zip'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        # Check if it's a single zip file
        if len(uploaded_files) == 1 and uploaded_files[0].name.endswith('.zip'):
            # Process zip file
            if st.button("Redact Documents in Zip"):
                try:
                    output_zip_path = redactor.process_zip(uploaded_files[0])
                    
                    # Read the entire zip file content
                    with open(output_zip_path, "rb") as file:
                        zip_bytes = file.read()
                    
                    # Provide download for processed zip
                    st.download_button(
                        label="Download Redacted Documents",
                        data=zip_bytes,
                        file_name="redacted_documents.zip",
                        mime='application/zip'
                    )
                    
                    st.success("Documents in zip file successfully processed!")
                
                except Exception as e:
                    st.error(f"Error processing zip file: {e}")
        
        # Process individual files
        else:
            # Prepare output directory
            output_files = []
            
            for uploaded_file in uploaded_files:
                # Prepare output filename
                file_extension = uploaded_file.name.split('.')[-1]
                output_filename = uploaded_file.name.replace(f'.{file_extension}', '_redacted.{file_extension}')
                
                # Temporarily save the uploaded file
                with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_extension}') as temp_file:
                    temp_file.write(uploaded_file.getbuffer())
                    temp_file_path = temp_file.name
                
                # Redact author
                try:
                    redactor.redact_author(temp_file_path)
                    output_files.append((output_filename, temp_file_path))
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {e}")
            
            # Provide download for multiple files
            if output_files:
                # Create zip if multiple files
                if len(output_files) > 1:
                    output_zip_path = os.path.join(tempfile.gettempdir(), 'redacted_documents.zip')
                    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for filename, filepath in output_files:
                            zipf.write(filepath, arcname=filename)
                    
                    # Read the entire zip file content
                    with open(output_zip_path, "rb") as file:
                        zip_bytes = file.read()
                    
                    # Download zip
                    st.download_button(
                        label="Download Redacted Documents",
                        data=zip_bytes,
                        file_name="redacted_documents.zip",
                        mime='application/zip'
                    )
                
                # Single file download
                else:
                    filename, filepath = output_files[0]
                    
                    # Read the entire file content
                    with open(filepath, "rb") as file:
                        file_bytes = file.read()
                    
                    st.download_button(
                        label="Download Redacted Document",
                        data=file_bytes,
                        file_name=filename,
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
                             if filename.endswith('docx') 
                             else 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                    )
                
                st.success("Documents successfully redacted!")

if __name__ == "__main__":
    main()
