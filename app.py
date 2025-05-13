import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io
import os

# --- Helper Functions ---

def parse_names(names_string):
    """Parses comma-separated names, strips whitespace, and filters empty strings."""
    if not names_string:
        return []
    names = [name.strip() for name in names_string.split(',')]
    return [name for name in names if name] # Filter out empty names

def redact_text_in_runs(runs, names_to_redact, redaction_string="[REDACTED]"):
    """
    Iterates through runs and replaces occurrences of names.
    Sorts names by length (descending) to handle overlapping names correctly
    (e.g., "John Smith" before "Smith").
    """
    sorted_names = sorted(names_to_redact, key=len, reverse=True)
    modified = False
    for run in runs:
        original_text = run.text
        new_text = original_text
        for name in sorted_names:
            # Case-insensitive replacement
            # This simple replace might catch parts of words if not careful.
            # For more robust matching, regex with word boundaries \b would be better,
            # but python-docx/pptx operate on runs, making regex across runs complex.
            # This approach is a good starting point for many use cases.
            if name.lower() in new_text.lower():
                # To maintain original casing as much as possible before redaction:
                # This is a bit tricky. A simpler approach is to just replace.
                # For a more precise case-preserving replace, regex with re.sub and a function
                # would be needed, which is much harder with run-based text.
                # Let's stick to a simpler replace for now.
                start_index = 0
                while True:
                    idx = new_text.lower().find(name.lower(), start_index)
                    if idx == -1:
                        break
                    new_text = new_text[:idx] + redaction_string + new_text[idx+len(name):]
                    start_index = idx + len(redaction_string) # Move past the inserted redaction
                    modified = True
        if new_text != original_text:
            run.text = new_text
    return modified

def redact_docx(docx_file, names_to_redact, redaction_string):
    """Redacts names in a .docx file stream and returns a BytesIO object."""
    doc = Document(docx_file)
    modified_doc = False

    # Redact in paragraphs
    for para in doc.paragraphs:
        if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
            modified_doc = True

    # Redact in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
                        modified_doc = True
    
    # Redact in headers and footers (more complex, might need specific handling if required)
    # For simplicity, this example focuses on main body and tables.

    if modified_doc:
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    return None # No changes made

def redact_pptx(pptx_file, names_to_redact, redaction_string):
    """Redacts names in a .pptx file stream and returns a BytesIO object."""
    prs = Presentation(pptx_file)
    modified_prs = False

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
                        modified_prs = True
            
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        # Check if cell.text_frame exists before accessing paragraphs
                        if cell.text_frame:
                            for para in cell.text_frame.paragraphs:
                                if redact_text_in_runs(para.runs, names_to_redact, redaction_string):
                                    modified_prs = True
            
            # Notes slides (if needed)
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
    return None # No changes made

# --- Streamlit App ---
st.set_page_config(layout="wide")
st.title("📄 Word & PowerPoint Name Redactor 📛")

st.sidebar.header("Instructions")
st.sidebar.info(
    """
    1.  **Upload** your Word (.docx) and/or PowerPoint (.pptx) files.
    2.  Enter a **comma-separated list of names** to redact (e.g., John Doe, Jane Smith, Alpha Corp).
        The redaction is case-insensitive.
    3.  Optionally, change the **redaction text** (default is "[REDACTED]").
    4.  Click "**Redact Files**".
    5.  Download your redacted files.
    """
)

uploaded_files = st.file_uploader(
    "Upload .docx or .pptx files",
    type=["docx", "pptx"],
    accept_multiple_files=True
)

names_input_str = st.text_area(
    "Names to redact (comma-separated):",
    placeholder="e.g., John Doe, Jane Smith, Dr. Evil, Contoso Ltd",
    height=100
)

redaction_text = st.text_input("Redaction text:", value="[REDACTED]")

if st.button("Redact Files  Procesados"):
    if not uploaded_files:
        st.warning("Please upload at least one file.")
    elif not names_input_str.strip():
        st.warning("Please enter at least one name to redact.")
    else:
        names_list = parse_names(names_input_str)
        if not names_list:
            st.warning("No valid names provided after parsing.")
        else:
            st.info(f"Attempting to redact the following names: {', '.join(names_list)}")
            st.info(f"Using redaction string: '{redaction_text}'")
            
            processed_files_count = 0
            redacted_files_data = [] # To store (filename, BytesIO_object)

            with st.spinner("Processing files... This might take a moment."):
                for uploaded_file in uploaded_files:
                    file_name = uploaded_file.name
                    file_extension = os.path.splitext(file_name)[1].lower()
                    redacted_content = None
                    
                    st.write(f"--- Processing: **{file_name}** ---")

                    try:
                        if file_extension == ".docx":
                            redacted_content = redact_docx(uploaded_file, names_list, redaction_text)
                        elif file_extension == ".pptx":
                            redacted_content = redact_pptx(uploaded_file, names_list, redaction_text)
                        
                        if redacted_content:
                            redacted_files_data.append((f"redacted_{file_name}", redacted_content))
                            st.success(f"Successfully redacted content in **{file_name}**.")
                            processed_files_count += 1
                        else:
                            st.info(f"No names found to redact in **{file_name}**, or file was not modified.")
                            # Offer original for download if no changes
                            uploaded_file.seek(0) # Reset stream position
                            redacted_files_data.append((f"original_{file_name}", uploaded_file))


                    except Exception as e:
                        st.error(f"Error processing {file_name}: {e}")
                        # Optionally, log the full traceback for debugging
                        # import traceback
                        # st.error(traceback.format_exc())

            if redacted_files_data:
                st.markdown("---")
                st.subheader("Download Redacted Files")
                cols = st.columns(3) # Adjust number of columns as needed
                col_idx = 0
                for i, (filename, data) in enumerate(redacted_files_data):
                    with cols[col_idx % len(cols)]:
                        st.download_button(
                            label=f"Download {filename}",
                            data=data,
                            file_name=filename,
                            mime= "application/octet-stream" # Generic mime type
                        )
                    col_idx += 1
            
            if processed_files_count == 0 and uploaded_files:
                st.info("No files were modified or no names were found matching the criteria.")
            elif processed_files_count > 0:
                st.balloons()
