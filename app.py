import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import fitz  # PyMuPDF
import io
import os
import zipfile

# --- Helper Functions ---

def parse_names(names_string):
    """Parses comma-separated names, strips whitespace, and filters empty strings."""
    if not names_string:
        return []
    names = [name.strip() for name in names_string.split(',')]
    return [name for name in names if name]

def redact_text_in_runs(runs, names_to_redact, redaction_string="[REDACTED]"):
    """
    Iterates through runs and replaces occurrences of names.
    Sorts names by length (descending) to handle overlapping names correctly.
    """
    sorted_names = sorted(names_to_redact, key=len, reverse=True)
    modified = False
    for run in runs:
        original_text = run.text
        new_text = original_text
        for name in sorted_names:
            start_index = 0
            while True:
                # Case-insensitive find
                idx = new_text.lower().find(name.lower(), start_index)
                if idx == -1:
                    break
                # Replace preserving the segment of new_text
                new_text = new_text[:idx] + redaction_string + new_text[idx+len(name):]
                start_index = idx + len(redaction_string)
                modified = True
        if new_text != original_text:
            run.text = new_text
    return modified

def redact_docx(docx_file_stream, names_to_redact, redaction_string):
    """Redacts names in a .docx file stream and returns a BytesIO object."""
    doc = Document(docx_file_stream)
    modified_doc = False

    for para in doc.paragraphs:
        if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
            modified_doc = True

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
                        modified_doc = True
    
    # Basic Header/Footer Redaction (can be expanded)
    for section in doc.sections:
        for header_para in section.header.paragraphs:
            if redact_text_in_runs(header_para.runs, names_to_redact, redaction_string):
                modified_doc = True
        for footer_para in section.footer.paragraphs:
            if redact_text_in_runs(footer_para.runs, names_to_redact, redaction_string):
                modified_doc = True

    if modified_doc:
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    return None

def redact_pptx(pptx_file_stream, names_to_redact, redaction_string):
    """Redacts names in a .pptx file stream and returns a BytesIO object."""
    prs = Presentation(pptx_file_stream)
    modified_prs = False

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
                        modified_prs = True
            
            if shape.has_table:
                table = shape.table
                for r_idx in range(len(table.rows)):
                    for c_idx in range(len(table.columns)):
                        cell = table.cell(r_idx, c_idx)
                        if cell.text_frame:
                            for para in cell.text_frame.paragraphs:
                                if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
                                    modified_prs = True
            
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide
            if notes_slide.notes_text_frame:
                for para in notes_slide.notes_text_frame.paragraphs:
                     if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
                        modified_prs = True

    if modified_prs:
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio
    return None

def redact_pdf(pdf_file_stream, names_to_redact, redaction_string):
    """Redacts names in a .pdf file stream and returns a BytesIO object."""
    pdf_bytes = pdf_file_stream.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    modified_pdf = False
    
    # Sort names by length (descending) to handle overlapping names (e.g., "John Smith" before "Smith")
    sorted_names = sorted(names_to_redact, key=len, reverse=True)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        page_modified_this_iteration = True

        # It's often better to re-search after each redaction if names overlap significantly
        # or redaction changes layout. For simplicity, we iterate names then apply.
        # A more robust approach might involve multiple passes or more complex logic.
        
        redactions_for_page = []

        for name in sorted_names:
            # TEXT_SEARCH_CASE_INSENSITIVE is default in newer PyMuPDF, explicit for clarity
            text_instances = page.search_for(name, flags=fitz.TEXT_SEARCH_CASE_INSENSITIVE | fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
            for inst in text_instances:
                # Add redaction annotation
                # Using fill=(0,0,0) for black box.
                # The `text` param in add_redact_annot is for the annotation itself, not usually visible replacement.
                # For visual replacement with `redaction_string`, one would typically draw a rect
                # then insert text, which is more complex. Blacking out is standard.
                annot = page.add_redact_annot(inst, text="", fill=(0, 0, 0)) # Black fill
                if annot: # Check if annotation was successfully added
                    redactions_for_page.append(annot) # Keep track if needed, though apply_redactions works on all
                    modified_pdf = True
        
        if redactions_for_page: # Only apply if annotations were added
            # Apply all redactions on the page.
            # images=fitz.PDF_REDACT_IMAGE_NONE: Don't remove images touched by redaction rect.
            # images=fitz.PDF_REDACT_IMAGE_PIXELS: Pixelate images under redaction.
            # images=fitz.PDF_REDACT_IMAGE_REMOVE: Remove images under redaction.
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE) 

    if modified_pdf:
        bio = io.BytesIO()
        # Save with garbage collection and deflation for smaller size
        doc.save(bio, garbage=3, deflate=True, clean=True)
        bio.seek(0)
        doc.close()
        return bio
    
    doc.close()
    return None

# --- Streamlit App ---
st.set_page_config(layout="wide", page_title="File Redactor")
st.title("📄 Word, PowerPoint & PDF Name Redactor 📛")

st.markdown("""
Welcome to the File Redactor! This tool helps you remove sensitive names from your documents.
Upload your files, specify the names, and download the redacted versions.
""")
st.markdown("---")

# Using columns for a slightly better layout
col1, col2 = st.columns([2,1])

with col1:
    uploaded_files = st.file_uploader(
        "Upload .docx, .pptx, .pdf, or .zip files containing them",
        type=["docx", "pptx", "pdf", "zip"],
        accept_multiple_files=True,
        help="You can upload multiple files or a ZIP archive."
    )

with col2:
    names_input_str = st.text_area(
        "Names to redact (comma-separated):",
        placeholder="e.g., John Doe, Jane Smith, Dr. Evil, Contoso Ltd",
        height=150,
        help="List all names, case-insensitive. Longer names are processed first (e.g., 'John Smith' before 'Smith')."
    )
    redaction_text_docx_pptx = st.text_input("Redaction text for DOCX/PPTX:", value="[REDACTED]",
                                        help="This text will replace names in Word and PowerPoint files.")
    st.caption("For PDF files, names will be blacked out.")


if st.button("Redact Files", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ Please upload at least one file.")
    elif not names_input_str.strip():
        st.warning("⚠️ Please enter at least one name to redact.")
    else:
        names_list = parse_names(names_input_str)
        if not names_list:
            st.warning("⚠️ No valid names provided after parsing.")
        else:
            st.info(f"Attempting to redact: **{', '.join(names_list)}**")
            st.info(f"DOCX/PPTX redaction string: '{redaction_text_docx_pptx}'. PDFs will be blacked out.")
            
            processed_files_count = 0
            redacted_files_info = [] # List to store tuples of (original_name, display_name, BytesIO_object)

            with st.spinner("🔧 Processing files... This might take a moment..."):
                for uploaded_file_obj in uploaded_files:
                    original_file_name = uploaded_file_obj.name
                    file_extension = os.path.splitext(original_file_name)[1].lower()
                    
                    files_to_process = [] # List of (BytesIO_stream, name_for_processing, original_ext)

                    # Handle ZIP files
                    if file_extension == ".zip":
                        st.write(f"--- Extracting from ZIP: **{original_file_name}** ---")
                        try:
                            with zipfile.ZipFile(uploaded_file_obj, 'r') as zip_ref:
                                for member_name in zip_ref.namelist():
                                    member_ext = os.path.splitext(member_name)[1].lower()
                                    if member_ext in [".docx", ".pptx", ".pdf"]:
                                        if not member_name.startswith('__MACOSX') and not member_name.endswith('/'): # Skip macOS resource forks and directories
                                            st.caption(f"Found {member_ext} file in ZIP: {member_name}")
                                            member_bytes = zip_ref.read(member_name)
                                            files_to_process.append(
                                                (io.BytesIO(member_bytes), 
                                                 f"{os.path.splitext(original_file_name)[0]}_{member_name}", # More descriptive name
                                                 member_ext)
                                            )
                                    elif member_ext: # Log other file types if needed
                                        st.caption(f"Skipping non-supported file in ZIP: {member_name}")
                        except zipfile.BadZipFile:
                            st.error(f"❌ Error: Could not read ZIP file '{original_file_name}'. It might be corrupted.")
                            continue # Skip to next uploaded file
                        except Exception as e:
                            st.error(f"❌ Error processing ZIP file '{original_file_name}': {e}")
                            continue
                    else:
                        # For non-ZIP files, reset stream position
                        uploaded_file_obj.seek(0)
                        files_to_process.append((uploaded_file_obj, original_file_name, file_extension))

                    # Process each identified file (either single upload or extracted from ZIP)
                    for file_stream, current_file_name_for_processing, current_file_ext in files_to_process:
                        st.write(f"--- Processing: **{current_file_name_for_processing}** ---")
                        redacted_content = None
                        try:
                            if current_file_ext == ".docx":
                                redacted_content = redact_docx(file_stream, names_list, redaction_text_docx_pptx)
                            elif current_file_ext == ".pptx":
                                redacted_content = redact_pptx(file_stream, names_list, redaction_text_docx_pptx)
                            elif current_file_ext == ".pdf":
                                redacted_content = redact_pdf(file_stream, names_list, redaction_text_docx_pptx) # redaction_text not used for PDF visual
                            
                            if redacted_content:
                                display_name = f"redacted_{current_file_name_for_processing}"
                                redacted_files_info.append((original_file_name, display_name, redacted_content, current_file_ext))
                                st.success(f"✅ Successfully redacted: **{current_file_name_for_processing}**")
                                processed_files_count += 1
                            else:
                                st.info(f"ℹ️ No redactions made in **{current_file_name_for_processing}** (no names found or file unchanged).")
                                # Offer original for download if no changes
                                file_stream.seek(0) 
                                display_name = f"original_{current_file_name_for_processing}"
                                redacted_files_info.append((original_file_name, display_name, file_stream, current_file_ext))


                        except Exception as e:
                            st.error(f"❌ Error processing **{current_file_name_for_processing}**: {e}")
                            # import traceback
                            # st.error(traceback.format_exc()) # Uncomment for detailed debugging

            if redacted_files_info:
                st.markdown("---")
                st.subheader("⬇️ Download Files")
                
                # Prepare for download buttons in columns
                max_cols = 3 
                cols = st.columns(max_cols)
                col_idx = 0

                for orig_name, display_name, data_stream, ext in redacted_files_info:
                    mime_types = {
                        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        ".pdf": "application/pdf",
                    }
                    mime_type = mime_types.get(ext, "application/octet-stream")
                    
                    # Ensure data_stream is ready to be read
                    data_stream.seek(0)
                    
                    with cols[col_idx % max_cols]:
                        st.download_button(
                            label=f"Download {display_name}",
                            data=data_stream,
                            file_name=display_name,
                            mime=mime_type,
                            key=f"download_{display_name}_{orig_name}" # Unique key
                        )
                    col_idx += 1
            
            if processed_files_count == 0 and uploaded_files:
                st.info("ℹ️ No files were modified as no matching names were found, or uploaded files were not of supported types within ZIPs.")
            elif processed_files_count > 0:
                st.balloons()

st.markdown("---")
with st.expander("📜 Instructions & Notes", expanded=False):
    st.markdown(
        """
        #### How to Use:
        1.  **Upload Files:**
            *   Click "Browse files" to select Word (.docx), PowerPoint (.pptx), PDF (.pdf) files, or ZIP archives (.zip) containing these file types.
            *   You can upload multiple files at once.
        2.  **Enter Names:**
            *   In the "Names to redact" box, type the names you want to remove, separated by commas (e.g., `John Doe, Jane Smith, Confidential Project`).
            *   The redaction is case-insensitive.
            *   The tool prioritizes longer names (e.g., if "Jane Smith" and "Smith" are both listed, "Jane Smith" will be redacted first).
        3.  **Redaction Text (for DOCX/PPTX):**
            *   Optionally, change the text that will replace the names in Word and PowerPoint files (default is `[REDACTED]`).
            *   **For PDF files, names will be blacked out.** The redaction text input does not apply to PDF visual output.
        4.  **Redact:**
            *   Click the "Redact Files" button.
        5.  **Download:**
            *   Download buttons for the processed files will appear.
            *   If a file had no names to redact, it will be offered as "original\_filename".
            *   Files extracted from ZIP archives will be offered individually.

        #### Important Notes:
        *   **PDF Redaction:** Names in PDF files are "blacked out" by covering them. The underlying text is removed where the redaction is applied.
        *   **ZIP Files:** The app extracts supported files from ZIPs and processes them. It does **not** re-ZIP the redacted files.
        *   **Complex Documents:** For very complex layouts, embedded objects, or scanned (image-based) PDFs without OCR text, redaction might be incomplete. This tool works best with text-based documents.
        *   **Backup:** Always keep a backup of your original files before redacting.
        """
    )
