import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import fitz  # PyMuPDF
import io
import os
import zipfile
import re # For splitting filename into parts

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
    search_flags = 1 | 2 

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        redactions_for_page = []
        for name in sorted_names:
            try:
                text_instances = page.search_for(name, flags=search_flags)
            except AttributeError: 
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

# --- New Helper Function to Extract Initials ---
def extract_initials_from_filename(filename_base):
    """
    Attempts to extract initials from a filename base.
    e.g., "John-Doe_Document" -> "JD"
    "Activity 6 Kaitlynn Deardorff" -> "KD" (takes last two typically)
    "singleword" -> "S"
    """
    # Remove common extra characters and split into potential name parts
    # Split by space, hyphen, underscore. Filter out empty strings.
    parts = re.split(r'[\s_-]+', filename_base)
    parts = [part for part in parts if part] # Remove empty strings from multiple delimiters

    if not parts:
        return ""

    initials = []
    # Try to get initials from the first few words, or last few if it seems like a name
    # This is heuristic. A more robust solution would need specific filename patterns.
    
    # Simple approach: take the first letter of each part, up to 2-3 initials
    # and convert to uppercase.
    for part in parts:
        if part and part[0].isalpha(): # Ensure it starts with a letter
            initials.append(part[0].upper())
    
    if len(initials) > 2: # If many parts, e.g., "Very Long Document Title From User"
                          # try to be smarter, maybe take first and last, or first two.
                          # For now, let's cap at 2-3 or based on common name patterns.
                          # A common pattern is Firstname Lastname.
        if len(parts) >= 2:
            # If "FirstName LastName Other stuff", try taking first letters of first two.
            # If "Other Stuff FirstName LastName", try taking first letters of last two.
            # Let's try a simple heuristic: take the first letter of the first two alphabetic parts.
            alpha_initials = [p[0].upper() for p in parts if p and p[0].isalpha()]
            if len(alpha_initials) >= 2:
                return "".join(alpha_initials[:2])
            elif alpha_initials:
                return alpha_initials[0]
            else: # no alphabetic parts
                return "" 
        else: # Only one part
             return initials[0] if initials else ""
    
    return "".join(initials)


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
st.markdown("Welcome! This tool redacts names from documents. Configure settings, upload files, and download.")
st.markdown("---")

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
    st.caption("Initials for files within the ZIP will be automatically extracted from their original filenames if possible.")


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
            files_for_download = [] 

            with st.spinner("🔧 Processing files... This might take a moment..."):
                for uploaded_file_obj in uploaded_files:
                    original_input_name = uploaded_file_obj.name
                    file_extension = os.path.splitext(original_input_name)[1].lower()
                    
                    if file_extension == ".zip":
                        st.write(f"--- Processing ZIP: **{original_input_name}** ---")
                        processed_zip_members_data = [] 
                        members_found_for_processing_in_zip = False
                        
                        custom_zip_name_base_from_input = output_zip_name_user_input.strip()
                        actual_output_zip_base_name = custom_zip_name_base_from_input if custom_zip_name_base_from_input \
                                                      else f"redacted_{os.path.splitext(original_input_name)[0]}"
                        
                        try:
                            with zipfile.ZipFile(uploaded_file_obj, 'r') as zip_ref:
                                for member_name in zip_ref.namelist(): # member_name is full path in zip
                                    member_base_name, member_ext_zip = os.path.splitext(os.path.basename(member_name)) # Get just the filename part
                                    
                                    if member_name.endswith('/') or member_name.startswith('__MACOSX'):
                                        continue 

                                    if member_ext_zip.lower() in [".docx", ".pptx", ".pdf"]:
                                        members_found_for_processing_in_zip = True
                                        st.caption(f"  Processing member: {member_name}")
                                        member_bytes = zip_ref.read(member_name)
                                        member_stream = io.BytesIO(member_bytes)
                                        
                                        redacted_member_content = None
                                        if member_ext_zip.lower() == ".docx": redacted_member_content = redact_docx(member_stream, names_list, redaction_text_docx_pptx)
                                        elif member_ext_zip.lower() == ".pptx": redacted_member_content = redact_pptx(member_stream, names_list, redaction_text_docx_pptx)
                                        elif member_ext_zip.lower() == ".pdf": redacted_member_content = redact_pdf(member_stream, names_list, redaction_text_docx_pptx)
                                        
                                        # Store original full member path, and the processed content
                                        if redacted_member_content:
                                            processed_zip_members_data.append((member_name, redacted_member_content))
                                            st.success(f"    ✅ Redacted: {member_name}")
                                            overall_docs_modified_count += 1
                                        else: 
                                            member_stream.seek(0)
                                            processed_zip_members_data.append((member_name, member_stream))
                                            st.info(f"    ℹ️ No redactions in: {member_name} (original included)")
                            
                            if members_found_for_processing_in_zip and processed_zip_members_data:
                                output_zip_stream = io.BytesIO()
                                file_counter_in_zip = 1
                                with zipfile.ZipFile(output_zip_stream, 'w', zipfile.ZIP_DEFLATED) as new_zip_archive:
                                    for m_full_path_in_zip, m_stream_content in processed_zip_members_data:
                                        m_stream_content.seek(0)
                                        
                                        # Extract initials from the original member's base filename
                                        original_member_base_name = os.path.splitext(os.path.basename(m_full_path_in_zip))[0]
                                        extracted_initials = extract_initials_from_filename(original_member_base_name)
                                        original_member_ext = os.path.splitext(m_full_path_in_zip)[1] # Get original extension

                                        if extracted_initials:
                                            internal_filename_in_zip = f"{actual_output_zip_base_name}_{extracted_initials}_{file_counter_in_zip:04d}{original_member_ext}"
                                        else: # Fallback if no initials could be extracted
                                            internal_filename_in_zip = f"{actual_output_zip_base_name}_{file_counter_in_zip:04d}{original_member_ext}"
                                        
                                        # Preserve directory structure if m_full_path_in_zip contains it
                                        if os.path.dirname(m_full_path_in_zip):
                                            internal_filename_in_zip = os.path.join(os.path.dirname(m_full_path_in_zip), os.path.basename(internal_filename_in_zip))
                                        
                                        new_zip_archive.writestr(internal_filename_in_zip, m_stream_content.read())
                                        file_counter_in_zip +=1
                                
                                output_zip_stream.seek(0)
                                display_zip_name = f"{actual_output_zip_base_name}.zip"
                                files_for_download.append(
                                    (original_input_name, display_zip_name, output_zip_stream, ".zip")
                                )
                                st.success(f"📦 Created new ZIP: **{display_zip_name}**.")
                                st.caption(f"   Files inside '{display_zip_name}' are renamed using base '{actual_output_zip_base_name}', auto-extracted initials (if any), and a number.")

                            elif not members_found_for_processing_in_zip:
                                st.info(f"ℹ️ No supported files (.docx, .pptx, .pdf) found in ZIP: **{original_input_name}** to process.")
                        except zipfile.BadZipFile: st.error(f"❌ Error: ZIP '{original_input_name}' appears to be corrupted.")
                        except Exception as e: st.error(f"❌ Error processing ZIP '{original_input_name}': {e}")
                    
                    else: # Process individual (non-ZIP) files
                        uploaded_file_obj.seek(0)
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
                                uploaded_file_obj.seek(0) 
                                display_name = f"original_{original_input_name}"
                                files_for_download.append((original_input_name, display_name, uploaded_file_obj, file_extension))
                        except Exception as e: st.error(f"❌ Error processing **{original_input_name}**: {e}")

            if files_for_download:
                st.markdown("---"); st.subheader("⬇️ Download Files")
                num_files = len(files_for_download)
                max_cols = 3 
                num_cols_to_use = min(num_files, max_cols) if num_files > 0 else 1
                cols = st.columns(num_cols_to_use)
                for i, (orig_name, display_name, data_stream, ext) in enumerate(files_for_download):
                    mime_types = { ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                   ".pdf": "application/pdf", ".zip": "application/zip"}
                    mime_type = mime_types.get(ext, "application/octet-stream")
                    data_stream.seek(0) 
                    col_index = i % num_cols_to_use
                    with cols[col_index]:
                        st.download_button(
                            label=f"Download {display_name}", data=data_stream,
                            file_name=display_name, mime=mime_type,
                            key=f"dl_btn_{i}_{display_name.replace(' ','_').replace('.','_').replace('/','_')}"
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
            *   Enter **Names to redact**.
            *   Set custom **Redaction text** for DOCX/PPTX.
            *   For **ZIP Output Options** (only for .zip uploads):
                *   **Custom output ZIP name base:** Define the base name for the output ZIP file. (e.g., `MyProject` results in `MyProject.zip`). Defaults to `redacted_[original_zip_name]`.
                *   **Initials for renaming files *inside* the ZIP** are now automatically extracted from the original filenames of the member files (e.g., a file named `John-Doe-Report.docx` might contribute `JD` as initials).
        2.  **Upload Files (Main Area):** Select .docx, .pptx, .pdf, or .zip files.
        3.  **Redact:** Click "Redact Files".
        4.  **Download:**
            *   Processed individual files are available.
            *   For uploaded ZIPs, a new ZIP archive is created. Files inside this new ZIP will be named: `[OutputZipNameBase]_[ExtractedInitials]_[Number].ext`. If initials cannot be extracted, that part is omitted.

        #### Important Notes:
        *   **PDFs:** Names are blacked out.
        *   **ZIPs:** Output is a new ZIP. Auto-extracted initials depend on the filename structure of the original files within the ZIP. Directory structure within the original ZIP is preserved in the output ZIP.
        *   **Complex Docs/Scanned PDFs:** Redaction might be incomplete.
        *   **Backup:** Always keep originals!
        """)
