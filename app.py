import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import fitz  # PyMuPDF
import io
import os
import zipfile

# --- Helper Functions (Unchanged from your previous correct versions) ---
def parse_names(names_string):
    if not names_string: return []
    names = [name.strip() for name in names_string.split(',')]
    return [name for name in names if name]

def redact_text_in_runs(runs, names_to_redact, redaction_string="[REDACTED]"):
    sorted_names = sorted(names_to_redact, key=len, reverse=True)
    modified = False
    for run in runs:
        original_text = run.text
        new_text = original_text
        for name in sorted_names:
            start_index = 0
            while True:
                idx = new_text.lower().find(name.lower(), start_index)
                if idx == -1: break
                new_text = new_text[:idx] + redaction_string + new_text[idx+len(name):]
                start_index = idx + len(redaction_string)
                modified = True
        if new_text != original_text:
            run.text = new_text
    return modified

def redact_docx(docx_file_stream, names_to_redact, redaction_string):
    doc = Document(docx_file_stream)
    modified_doc = False
    for para in doc.paragraphs:
        if redact_text_in_runs(para.runs, names_to_redact, redaction_string): modified_doc = True
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if redact_text_in_runs(para.runs, names_to_redact, redaction_string): modified_doc = True
    for section in doc.sections: # Process headers/footers
        for header_para in section.header.paragraphs:
            if redact_text_in_runs(header_para.runs, names_to_redact, redaction_string): modified_doc = True
        for footer_para in section.footer.paragraphs:
            if redact_text_in_runs(footer_para.runs, names_to_redact, redaction_string): modified_doc = True
    if modified_doc:
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    return None

def redact_pptx(pptx_file_stream, names_to_redact, redaction_string):
    prs = Presentation(pptx_file_stream)
    modified_prs = False
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    if redact_text_in_runs(para.runs, names_to_redact, redaction_string): modified_prs = True
            if shape.has_table:
                table = shape.table
                for r_idx in range(len(table.rows)):
                    for c_idx in range(len(table.columns)):
                        cell = table.cell(r_idx, c_idx)
                        if cell.text_frame:
                            for para in cell.text_frame.paragraphs:
                                if redact_text_in_runs(para.runs, names_to_redact, redaction_string): modified_prs = True
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            for para in slide.notes_slide.notes_text_frame.paragraphs:
                 if redact_text_in_runs(para.runs, names_to_redact, redaction_string): modified_prs = True
    if modified_prs:
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio
    return None

def redact_pdf(pdf_file_stream, names_to_redact, redaction_string): # redaction_string not used for PDF visual
    pdf_bytes = pdf_file_stream.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    modified_pdf = False
    sorted_names = sorted(names_to_redact, key=len, reverse=True)
    # PyMuPDF flags: TEXT_PRESERVE_LIGATURES=1, TEXT_PRESERVE_WHITESPACE=2
    # search_for is case-insensitive by default in recent versions.
    search_flags = 1 | 2 

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        redactions_for_page = []
        for name in sorted_names:
            try:
                text_instances = page.search_for(name, flags=search_flags)
            except AttributeError: 
                text_instances = page.search_for(name) # Fallback for older PyMuPDF
            for inst in text_instances:
                annot = page.add_redact_annot(inst, text="", fill=(0, 0, 0))
                if annot:
                    redactions_for_page.append(annot)
                    modified_pdf = True
        if redactions_for_page:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    if modified_pdf:
        bio = io.BytesIO()
        doc.save(bio, garbage=3, deflate=True, clean=True)
        bio.seek(0)
        doc.close()
        return bio
    doc.close()
    return None

# --- Streamlit App ---
st.set_page_config(layout="wide", page_title="File Redactor")

st.markdown("""
<style>
    .stButton>button[kind="primary"] { 
        background-color: #4CAF50 !important; color: white !important; border: none !important;
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #45a049 !important; color: white !important;
    }
    .stButton>button[kind="primary"]:active {
        background-color: #3e8e41 !important; color: white !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("🐉 TU File Name Redactor")
st.markdown("Welcome to the TU File Name Redactor! This tool helps you remove sensitive names from your documents. Upload your files, specify the names, and download the redacted versions.")
st.markdown("---")

# --- Sidebar for Settings ---
with st.sidebar:
    st.header("⚙️ Redaction Settings")
    names_input_str = st.text_area(
        "Names to redact (comma-separated):", placeholder="e.g., John Doe, Jane Smith, Contoso Ltd",
        height=100, help="Case-insensitive. Longer names processed first.")
    
    redaction_text_docx_pptx = st.text_input(
        "Redaction text for DOCX/PPTX:", value="[REDACTED]",
        help="Text replacement for Word/PowerPoint. PDFs are blacked out.")
    
    st.markdown("---")
    st.subheader("📦 ZIP Output Options")
    st.caption("(These options apply *only* when processing an uploaded ZIP file)")
    
    output_zip_name_user_input = st.text_input(
        "Custom output ZIP name base:", 
        placeholder="e.g., MyProject_Redacted",
        help="If you upload a ZIP, the output ZIP will use this as its base name. If blank, defaults to 'redacted_[original_zip_name]'. '.zip' is added automatically."
    )
    
    user_initials_for_zip_files = st.text_input(
        "Your initials (for files in ZIP):", 
        max_chars=5, 
        placeholder="e.g., JD",
        help="If provided, files *inside* the output ZIP will be renamed to '[ZIPNameBase]_[Initials]_[Number].ext'."
    )

# --- Main Area for Upload and Downloads ---
uploaded_files = st.file_uploader(
    "Upload .docx, .pptx, .pdf, or .zip files", type=["docx", "pptx", "pdf", "zip"],
    accept_multiple_files=True, help="Upload individual files or ZIP archives.")

if st.button("Redact Files", type="primary", use_container_width=True, key="redact_button"):
    if not uploaded_files: st.warning("⚠️ Please upload at least one file.")
    elif not names_input_str.strip(): st.warning("⚠️ Please enter at least one name to redact.")
    else:
        names_list = parse_names(names_input_str)
        if not names_list: st.warning("⚠️ No valid names provided after parsing.")
        else:
            st.info(f"Attempting to redact: **{', '.join(names_list)}**")
            st.info(f"DOCX/PPTX redaction: '{redaction_text_docx_pptx}'. PDFs blacked out.")
            
            overall_docs_modified_count = 0
            files_for_download = [] # Stores (original_input_name, display_name_for_download, BytesIO_object, final_extension)

            with st.spinner("🔧 Processing files... This might take a moment..."):
                for uploaded_file_obj in uploaded_files:
                    original_input_name = uploaded_file_obj.name
                    file_extension = os.path.splitext(original_input_name)[1].lower()
                    
                    if file_extension == ".zip":
                        st.write(f"--- Processing ZIP: **{original_input_name}** ---")
                        processed_zip_members_data = [] # List of (original_member_name, BytesIO_of_processed_member_content)
                        members_found_for_processing_in_zip = False
                        
                        # Determine output ZIP name base and initials from sidebar inputs
                        custom_zip_name_base_from_input = output_zip_name_user_input.strip()
                        # Use custom name if provided, else default to "redacted_originalzipname"
                        actual_output_zip_base_name = custom_zip_name_base_from_input if custom_zip_name_base_from_input \
                                                      else f"redacted_{os.path.splitext(original_input_name)[0]}"
                        
                        current_user_initials = user_initials_for_zip_files.strip().upper() # Standardize initials to uppercase

                        try:
                            with zipfile.ZipFile(uploaded_file_obj, 'r') as zip_ref:
                                for member_name in zip_ref.namelist():
                                    member_ext_zip = os.path.splitext(member_name)[1].lower()
                                    
                                    if member_name.endswith('/') or member_name.startswith('__MACOSX'): # Skip directories and macOS resource forks
                                        continue 

                                    if member_ext_zip in [".docx", ".pptx", ".pdf"]:
                                        members_found_for_processing_in_zip = True
                                        st.caption(f"  Processing member: {member_name}")
                                        member_bytes = zip_ref.read(member_name)
                                        member_stream = io.BytesIO(member_bytes)
                                        
                                        redacted_member_content = None
                                        if member_ext_zip == ".docx": redacted_member_content = redact_docx(member_stream, names_list, redaction_text_docx_pptx)
                                        elif member_ext_zip == ".pptx": redacted_member_content = redact_pptx(member_stream, names_list, redaction_text_docx_pptx)
                                        elif member_ext_zip == ".pdf": redacted_member_content = redact_pdf(member_stream, names_list, redaction_text_docx_pptx)
                                        
                                        if redacted_member_content:
                                            processed_zip_members_data.append((member_name, redacted_member_content))
                                            st.success(f"    ✅ Redacted: {member_name}")
                                            overall_docs_modified_count += 1
                                        else: # No redactions, include original member content
                                            member_stream.seek(0) # Reset stream to beginning
                                            processed_zip_members_data.append((member_name, member_stream))
                                            st.info(f"    ℹ️ No redactions in: {member_name} (original included)")
                            
                            if members_found_for_processing_in_zip and processed_zip_members_data:
                                output_zip_stream = io.BytesIO()
                                file_counter_in_zip = 1
                                with zipfile.ZipFile(output_zip_stream, 'w', zipfile.ZIP_DEFLATED) as new_zip_archive:
                                    for m_orig_name, m_stream_content in processed_zip_members_data:
                                        m_stream_content.seek(0) # Ensure stream is at the beginning
                                        member_original_ext = os.path.splitext(m_orig_name)[1]
                                        
                                        internal_filename_in_zip = m_orig_name # Default to original name if no initials
                                        if current_user_initials: # Only rename if initials are provided
                                            internal_filename_in_zip = f"{actual_output_zip_base_name}_{current_user_initials}_{file_counter_in_zip:04d}{member_original_ext}"
                                        
                                        new_zip_archive.writestr(internal_filename_in_zip, m_stream_content.read())
                                        file_counter_in_zip +=1
                                
                                output_zip_stream.seek(0)
                                display_zip_name = f"{actual_output_zip_base_name}.zip"
                                files_for_download.append(
                                    (original_input_name, display_zip_name, output_zip_stream, ".zip")
                                )
                                st.success(f"📦 Created new ZIP: **{display_zip_name}** containing processed files.")
                                if current_user_initials:
                                    st.caption(f"   Files inside '{display_zip_name}' are renamed using base '{actual_output_zip_base_name}' and initials '{current_user_initials}'.")
                                else:
                                    st.caption(f"   Files inside '{display_zip_name}' retain their original names (as initials were not provided for renaming).")

                            elif not members_found_for_processing_in_zip:
                                st.info(f"ℹ️ No supported files (.docx, .pptx, .pdf) found in ZIP: **{original_input_name}** to process.")
                        except zipfile.BadZipFile: st.error(f"❌ Error: ZIP '{original_input_name}' appears to be corrupted.")
                        except Exception as e: st.error(f"❌ Error processing ZIP '{original_input_name}': {e}")
                    
                    else: # Process individual (non-ZIP) files
                        uploaded_file_obj.seek(0) # Reset stream for individual file
                        st.write(f"--- Processing: **{original_input_name}** ---")
                        redacted_content = None
                        try:
                            if file_extension == ".docx": redacted_content = redact_docx(uploaded_file_obj, names_list, redaction_text_docx_pptx)
                            elif file_extension == ".pptx": redacted_content = redact_pptx(uploaded_file_obj, names_list, redaction_text_docx_pptx)
                            elif file_extension == ".pdf": redacted_content = redact_pdf(uploaded_file_obj, names_list, redaction_text_docx_pptx)
                            
                            if redacted_content:
                                display_name = f"redacted_{original_input_name}"
                                files_for_download.append((original_input_name, display_name, redacted_content, file_extension))
                                st.success(f"✅ Successfully redacted: **{original_input_name}**")
                                overall_docs_modified_count += 1
                            else:
                                st.info(f"ℹ️ No redactions made in **{original_input_name}** (file unchanged).")
                                uploaded_file_obj.seek(0) # Reset for download
                                display_name = f"original_{original_input_name}"
                                files_for_download.append((original_input_name, display_name, uploaded_file_obj, file_extension))
                        except Exception as e: st.error(f"❌ Error processing **{original_input_name}**: {e}")

            if files_for_download:
                st.markdown("---"); st.subheader("⬇️ Download Files")
                # Dynamically adjust columns based on number of files, up to a max
                num_files = len(files_for_download)
                max_cols = 3 # Max columns for download buttons
                num_cols_to_use = min(num_files, max_cols) if num_files > 0 else 1
                
                cols = st.columns(num_cols_to_use)
                for i, (orig_name, display_name, data_stream, ext) in enumerate(files_for_download):
                    mime_types = { ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                   ".pdf": "application/pdf", ".zip": "application/zip"}
                    mime_type = mime_types.get(ext, "application/octet-stream")
                    data_stream.seek(0) # Ensure stream is at the beginning for download
                    
                    col_index = i % num_cols_to_use # Distribute download buttons across columns
                    with cols[col_index]:
                        st.download_button(
                            label=f"Download {display_name}", data=data_stream,
                            file_name=display_name, mime=mime_type,
                            key=f"dl_btn_{i}_{display_name.replace(' ','_').replace('.','_').replace('/','_')}" # Make key more robust
                        )
            
            if overall_docs_modified_count == 0 and uploaded_files:
                st.info("ℹ️ No documents were modified based on the provided names, or no supported files were found in ZIPs.")
            elif overall_docs_modified_count > 0:
                st.balloons()

st.markdown("---")
with st.expander("📜 Instructions & Notes", expanded=False):
    st.markdown("""
        #### How to Use:
        1.  **Configure Settings (Sidebar):**
            *   Enter **Names to redact** (comma-separated).
            *   Set custom **Redaction text** for DOCX/PPTX files if desired.
            *   Optionally, for **ZIP Output Options** (these apply *only* when you upload a .zip file):
                *   **Custom output ZIP name base:** If you input `MyProject`, an uploaded `data.zip` will result in `MyProject.zip`. If left blank, it defaults to `redacted_data.zip`.
                *   **Your initials:** If you input `JD` and the ZIP name base is `MyProject`, files inside the output ZIP will be named like `MyProject_JD_0001.docx`, `MyProject_JD_0002.pdf`, etc. If initials are not provided, internal files keep their original names.
        2.  **Upload Files (Main Area):** Select one or more .docx, .pptx, .pdf, or .zip files.
        3.  **Redact:** Click the "Redact Files" button.
        4.  **Download:**
            *   Processed individual files (or original if no changes) will be available for download.
            *   If you uploaded a ZIP file, a **new ZIP archive** will be created with the (potentially custom) name and (potentially renamed) internal files, ready for download.

        #### Important Notes:
        *   **PDF Redaction:** Names in PDF files are "blacked out." The custom redaction text does not apply to PDF visual output.
        *   **ZIP File Processing:** The "ZIP Output Options" in the sidebar control the naming of the output ZIP and its contents. Files within the original ZIP that are not .docx, .pptx, or .pdf are currently **not** included in the output ZIP.
        *   **Complex Documents:** For very complex layouts, embedded objects, or scanned (image-based) PDFs without OCR text, redaction might be incomplete. This tool works best with text-based documents.
        *   **Backup:** Always keep a backup of your original files before redacting!
        """)
