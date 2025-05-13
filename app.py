import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import fitz  # PyMuPDF
import io
import os
import zipfile

# --- Helper Functions (parse_names, redact_text_in_runs, redact_docx, redact_pptx, redact_pdf - UNCHANGED) ---
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
    for section in doc.sections:
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
    search_flags = fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        redactions_for_page = []
        for name in sorted_names:
            try:
                text_instances = page.search_for(name, flags=search_flags)
            except AttributeError: # Fallback for very old PyMuPDF versions or flag issues
                text_instances = page.search_for(name) 
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

st.title("📄 Word, PowerPoint & PDF Name Redactor 📛")
st.markdown("Welcome to the File Redactor! This tool helps you remove sensitive names from your documents. Upload your files, specify the names, and download the redacted versions.")
st.markdown("---")

col1, col2 = st.columns([2,1])
with col1:
    uploaded_files = st.file_uploader(
        "Upload .docx, .pptx, .pdf, or .zip files", type=["docx", "pptx", "pdf", "zip"],
        accept_multiple_files=True, help="Upload individual files or ZIP archives.")
with col2:
    names_input_str = st.text_area(
        "Names to redact (comma-separated):", placeholder="e.g., John Doe, Jane Smith, Contoso Ltd",
        height=150, help="Case-insensitive. Longer names processed first.")
    redaction_text_docx_pptx = st.text_input("Redaction text for DOCX/PPTX:", value="[REDACTED]",
                                        help="Text replacement for Word/PowerPoint. PDFs are blacked out.")
    st.caption("PDF names are blacked out.")

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
            # Stores (original_uploaded_filename, display_name_for_download, BytesIO_object, final_extension)
            files_for_download = [] 

            with st.spinner("🔧 Processing files... This might take a moment..."):
                for uploaded_file_obj in uploaded_files:
                    original_input_name = uploaded_file_obj.name
                    file_extension = os.path.splitext(original_input_name)[1].lower()
                    
                    if file_extension == ".zip":
                        st.write(f"--- Processing ZIP: **{original_input_name}** ---")
                        # Stores (member_name_in_zip, BytesIO_of_processed_member) for re-zipping
                        processed_zip_members_streams = [] 
                        members_found_for_processing_in_zip = False

                        try:
                            with zipfile.ZipFile(uploaded_file_obj, 'r') as zip_ref:
                                for member_name in zip_ref.namelist():
                                    member_ext = os.path.splitext(member_name)[1].lower()
                                    
                                    if member_name.endswith('/') or member_name.startswith('__MACOSX'):
                                        continue # Skip directories and macOS resource forks

                                    if member_ext in [".docx", ".pptx", ".pdf"]:
                                        members_found_for_processing_in_zip = True
                                        st.caption(f"  Processing member: {member_name}")
                                        member_bytes = zip_ref.read(member_name)
                                        member_stream = io.BytesIO(member_bytes)
                                        
                                        redacted_member_content = None
                                        if member_ext == ".docx":
                                            redacted_member_content = redact_docx(member_stream, names_list, redaction_text_docx_pptx)
                                        elif member_ext == ".pptx":
                                            redacted_member_content = redact_pptx(member_stream, names_list, redaction_text_docx_pptx)
                                        elif member_ext == ".pdf":
                                            redacted_member_content = redact_pdf(member_stream, names_list, redaction_text_docx_pptx)
                                        
                                        if redacted_member_content:
                                            processed_zip_members_streams.append((member_name, redacted_member_content))
                                            st.success(f"    ✅ Redacted: {member_name}")
                                            overall_docs_modified_count += 1
                                        else:
                                            member_stream.seek(0) # Reset original stream if not modified
                                            processed_zip_members_streams.append((member_name, member_stream))
                                            st.info(f"    ℹ️ No redactions in: {member_name} (will be included as original)")
                            
                            if members_found_for_processing_in_zip and processed_zip_members_streams:
                                output_zip_stream = io.BytesIO()
                                with zipfile.ZipFile(output_zip_stream, 'w', zipfile.ZIP_DEFLATED) as new_zip_archive:
                                    for m_name, m_stream in processed_zip_members_streams:
                                        m_stream.seek(0)
                                        new_zip_archive.writestr(m_name, m_stream.read())
                                
                                output_zip_stream.seek(0)
                                display_zip_name = f"redacted_{os.path.splitext(original_input_name)[0]}.zip"
                                files_for_download.append(
                                    (original_input_name, display_zip_name, output_zip_stream, ".zip")
                                )
                                st.success(f"📦 Created new ZIP: **{display_zip_name}** containing processed files.")
                            elif not members_found_for_processing_in_zip : # No supported files found at all
                                st.info(f"ℹ️ No supported files (.docx, .pptx, .pdf) found in ZIP: **{original_input_name}** to create a new ZIP.")
                            # If members_found_for_processing_in_zip is true but processed_zip_members_streams is empty, it's an anomaly (should not happen with current logic)

                        except zipfile.BadZipFile:
                            st.error(f"❌ Error: ZIP '{original_input_name}' is corrupted.")
                        except Exception as e:
                            st.error(f"❌ Error processing ZIP '{original_input_name}': {e}")
                    
                    else: # Process individual (non-ZIP) files
                        uploaded_file_obj.seek(0) # Reset stream for individual file
                        st.write(f"--- Processing: **{original_input_name}** ---")
                        redacted_content = None
                        try:
                            if file_extension == ".docx":
                                redacted_content = redact_docx(uploaded_file_obj, names_list, redaction_text_docx_pptx)
                            elif file_extension == ".pptx":
                                redacted_content = redact_pptx(uploaded_file_obj, names_list, redaction_text_docx_pptx)
                            elif file_extension == ".pdf":
                                redacted_content = redact_pdf(uploaded_file_obj, names_list, redaction_text_docx_pptx)
                            
                            if redacted_content:
                                display_name = f"redacted_{original_input_name}"
                                files_for_download.append((original_input_name, display_name, redacted_content, file_extension))
                                st.success(f"✅ Successfully redacted: **{original_input_name}**")
                                overall_docs_modified_count += 1
                            else:
                                st.info(f"ℹ️ No redactions in **{original_input_name}** (file unchanged).")
                                uploaded_file_obj.seek(0) 
                                display_name = f"original_{original_input_name}"
                                files_for_download.append((original_input_name, display_name, uploaded_file_obj, file_extension))
                        except Exception as e:
                            st.error(f"❌ Error processing **{original_input_name}**: {e}")

            if files_for_download:
                st.markdown("---"); st.subheader("⬇️ Download Files")
                max_cols = 3; cols = st.columns(max_cols)
                for i, (orig_name, display_name, data_stream, ext) in enumerate(files_for_download):
                    mime_types = {
                        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        ".pdf": "application/pdf", ".zip": "application/zip",
                    }
                    mime_type = mime_types.get(ext, "application/octet-stream")
                    data_stream.seek(0)
                    with cols[i % max_cols]:
                        st.download_button(
                            label=f"Download {display_name}", data=data_stream,
                            file_name=display_name, mime=mime_type,
                            key=f"dl_btn_{i}_{display_name.replace(' ','_').replace('.','_')}" 
                        )
            
            if overall_docs_modified_count == 0 and uploaded_files:
                st.info("ℹ️ No documents were modified based on the provided names.")
            elif overall_docs_modified_count > 0:
                st.balloons()

st.markdown("---")
with st.expander("📜 Instructions & Notes", expanded=False):
    st.markdown("""
        #### How to Use:
        1.  **Upload Files:** .docx, .pptx, .pdf, or .zip archives.
        2.  **Enter Names:** Comma-separated, case-insensitive.
        3.  **Redaction Text (DOCX/PPTX):** Customize replacement string. PDFs are blacked out.
        4.  **Redact:** Click "Redact Files".
        5.  **Download:**
            *   Processed individual files are offered for download.
            *   If you upload a ZIP, a **new ZIP archive** containing the processed versions of its supported files will be created and offered for download. Files within the ZIP that are not .docx, .pptx, or .pdf are currently **not** included in the output ZIP.
            *   If a file (or a member within a ZIP) had no names to redact, its original version will be included.

        #### Important Notes:
        *   **PDFs:** Names are blacked out.
        *   **ZIPs:** Output is a new ZIP with processed .docx, .pptx, .pdf files from the original.
        *   **Complex Docs/Scanned PDFs:** Redaction might be incomplete. Best for text-based files.
        *   **Backup:** Always keep originals!
        """)
