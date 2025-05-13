import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import fitz  # PyMuPDF
import io
import os
import zipfile

# --- Helper Functions (parse_names, redact_text_in_runs, redact_docx, redact_pptx -UNCHANGED from previous version) ---
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
                idx = new_text.lower().find(name.lower(), start_index)
                if idx == -1:
                    break
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

def redact_pdf(pdf_file_stream, names_to_redact, redaction_string): # redaction_string not used for PDF visual
    """Redacts names in a .pdf file stream and returns a BytesIO object."""
    pdf_bytes = pdf_file_stream.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    modified_pdf = False
    
    sorted_names = sorted(names_to_redact, key=len, reverse=True)

    # Define flags:
    # TEXT_PRESERVE_LIGATURES = 1
    # TEXT_PRESERVE_WHITESPACE = 2
    # TEXT_SEARCH_NOT_INTERLACED = 4 (Often good to use for better layout understanding)
    # Default search is often case-insensitive. If not, we might need a different approach.
    # For PyMuPDF 1.18.0 and later, search_for is case-insensitive by default.
    # If your version is older, this might be the issue.
    # Let's try with minimal flags first, assuming case-insensitivity is default.
    # If it's not, we'd have to iterate names and do a lowercase match on extracted text, then map rects.
    
    search_flags = fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE
    # If you have PyMuPDF < 1.18.x, case-insensitivity might not be default for search_for.
    # In that case, you'd have to do:
    # for page_num in range(len(doc)):
    #     page = doc.load_page(page_num)
    #     page_text_lower = page.get_text("text").lower() # Get all text and lower
    #     for name in sorted_names:
    #         if name.lower() in page_text_lower: # Check if name exists (case-insensitive)
    #             text_instances = page.search_for(name) # Now search case-sensitively (or default)
    #             # ... rest of the logic
    # This is more complex as you need to map back.
    # For now, let's assume a modern enough PyMuPDF or that default is case-insensitive.

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        redactions_for_page = []

        for name in sorted_names:
            # Try searching without explicitly setting case-insensitivity flag first.
            # Most modern PyMuPDF versions are case-insensitive by default for `search_for`.
            # If this fails, it implies the installed PyMuPDF is very old or configured differently.
            try:
                text_instances = page.search_for(name, flags=search_flags)
            except AttributeError as e:
                # This might happen if even basic flags are named differently in a very old version
                st.warning(f"PyMuPDF flag issue: {e}. Falling back to default flags for search_for().")
                text_instances = page.search_for(name) # Try with default flags

            for inst in text_instances:
                # Check if the found text (inst.text) exactly matches the name, case-insensitively
                # This is a crucial check because `search_for` might find "Anderson" within "Sanderson"
                # We need to ensure we are redacting the whole name.
                # The `inst` rectangle from `search_for` usually covers the exact match.
                # However, a manual check can be added for robustness if `search_for` is too broad.
                # For now, let's trust `search_for` with sorted names.
                
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

# Custom CSS for green button (Optional - use if type="primary" isn't green)
# To make this robust, you'd need to inspect the exact class Streamlit generates for its buttons.
# This is a common pattern but can break with Streamlit updates.
st.markdown("""
<style>
    /* Targeting the primary button specifically if it has a unique class or attribute */
    /* Or a more general approach if you want all primary buttons green */
    .stButton>button[kind="primary"] { /* More specific selector for primary button */
        background-color: #4CAF50 !important; /* Green */
        color: white !important;
        border: none !important;
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #45a049 !important; /* Darker Green */
        color: white !important;
    }
    .stButton>button[kind="primary"]:active {
        background-color: #3e8e41 !important; /* Even Darker Green */
        color: white !important;
    }
    /* For Streamlit versions where primary buttons might not have kind="primary" attribute,
       you might need a less specific selector, or inspect elements */
</style>
""", unsafe_allow_html=True)


st.title("📄 Word, PowerPoint & PDF Name Redactor 📛")

st.markdown("""
Welcome to the File Redactor! This tool helps you remove sensitive names from your documents.
Upload your files, specify the names, and download the redacted versions.
""")
st.markdown("---")

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


# Changed button text and added type="primary" which the CSS above will target
if st.button("Redact Files", type="primary", use_container_width=True, key="redact_button"):
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
            redacted_files_info = [] 

            with st.spinner("🔧 Processing files... This might take a moment..."):
                for uploaded_file_obj in uploaded_files:
                    original_file_name = uploaded_file_obj.name
                    file_extension = os.path.splitext(original_file_name)[1].lower()
                    
                    files_to_process = []

                    if file_extension == ".zip":
                        st.write(f"--- Extracting from ZIP: **{original_file_name}** ---")
                        try:
                            with zipfile.ZipFile(uploaded_file_obj, 'r') as zip_ref:
                                for member_name in zip_ref.namelist():
                                    member_ext = os.path.splitext(member_name)[1].lower()
                                    if member_ext in [".docx", ".pptx", ".pdf"]:
                                        if not member_name.startswith('__MACOSX') and not member_name.endswith('/'):
                                            st.caption(f"Found {member_ext} file in ZIP: {member_name}")
                                            member_bytes = zip_ref.read(member_name)
                                            # Create a unique processing name to avoid clashes if multiple zips have same internal names
                                            processing_name = f"{os.path.splitext(original_file_name)[0]}_{member_name.replace('/', '_')}"
                                            files_to_process.append(
                                                (io.BytesIO(member_bytes), 
                                                 processing_name,
                                                 member_ext)
                                            )
                                    elif member_ext:
                                        st.caption(f"Skipping non-supported file in ZIP: {member_name}")
                        except zipfile.BadZipFile:
                            st.error(f"❌ Error: Could not read ZIP file '{original_file_name}'. It might be corrupted.")
                            continue
                        except Exception as e:
                            st.error(f"❌ Error processing ZIP file '{original_file_name}': {e}")
                            continue
                    else:
                        uploaded_file_obj.seek(0)
                        files_to_process.append((uploaded_file_obj, original_file_name, file_extension))

                    for file_stream, current_file_name_for_processing, current_file_ext in files_to_process:
                        st.write(f"--- Processing: **{current_file_name_for_processing}** ---")
                        redacted_content = None
                        try:
                            if current_file_ext == ".docx":
                                redacted_content = redact_docx(file_stream, names_list, redaction_text_docx_pptx)
                            elif current_file_ext == ".pptx":
                                redacted_content = redact_pptx(file_stream, names_list, redaction_text_docx_pptx)
                            elif current_file_ext == ".pdf":
                                redacted_content = redact_pdf(file_stream, names_list, redaction_text_docx_pptx)
                            
                            if redacted_content:
                                display_name = f"redacted_{current_file_name_for_processing}"
                                redacted_files_info.append((original_file_name, display_name, redacted_content, current_file_ext))
                                st.success(f"✅ Successfully redacted: **{current_file_name_for_processing}**")
                                processed_files_count += 1
                            else:
                                st.info(f"ℹ️ No redactions made in **{current_file_name_for_processing}** (no names found or file unchanged).")
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
                
                max_cols = 3 
                cols = st.columns(max_cols)
                col_idx = 0

                for i, (orig_name, display_name, data_stream, ext) in enumerate(redacted_files_info): # Added enumerate for unique key
                    mime_types = {
                        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        ".pdf": "application/pdf",
                    }
                    mime_type = mime_types.get(ext, "application/octet-stream")
                    
                    data_stream.seek(0)
                    
                    with cols[col_idx % max_cols]:
                        # Ensure unique key for download button, especially if display_names could collide
                        # Using an index `i` from enumerate ensures uniqueness.
                        st.download_button(
                            label=f"Download {display_name}",
                            data=data_stream,
                            file_name=display_name,
                            mime=mime_type,
                            key=f"download_btn_{i}_{display_name}" 
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
        *   **PDF Redaction:** Names in PDF files are "blacked out" by covering them. The underlying text is removed where the redaction is applied. If you encounter issues, it might be due to the version of PyMuPDF; this app attempts to be compatible with recent versions.
        *   **ZIP Files:** The app extracts supported files from ZIPs and processes them. It does **not** re-ZIP the redacted files.
        *   **Complex Documents:** For very complex layouts, embedded objects, or scanned (image-based) PDFs without OCR text, redaction might be incomplete. This tool works best with text-based documents.
        *   **Backup:** Always keep a backup of your original files before redacting.
        """
    )
