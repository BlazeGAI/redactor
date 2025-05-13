import streamlit as st
import docx
from pptx import Presentation
import fitz  # PyMuPDF
import io
import re
import os

# --- Helper Functions for Redaction ---

def redact_names_in_text(text, names_to_redact, placeholder="[REDACTED NAME]"):
    """
    Redacts a list of names from a given text string.
    Uses regex for whole-word, case-insensitive matching.
    """
    redacted_text = text
    for name in names_to_redact:
        if not name.strip():  # Skip empty strings
            continue
        # Regex to match whole word, case-insensitive
        # \b ensures word boundaries (e.g., "Tom" doesn't match "Tomorrow")
        pattern = r'\b' + re.escape(name) + r'\b'
        redacted_text = re.sub(pattern, placeholder, redacted_text, flags=re.IGNORECASE)
    return redacted_text

def redact_docx(file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
    """
    Redacts names from the first two pages (conceptually, as Word doesn't have hard pages)
    of a DOCX file. It processes all paragraphs and tables.
    """
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        
        # Word doesn't have a strict concept of "pages" like PDF.
        # We will process all paragraphs and tables, assuming names are early.
        # For a more precise "first two pages" approach, one might need
        # to estimate based on content length or use commercial libraries.
        # For this assignment, processing all content is likely acceptable if names are expected early.

        # Redact in paragraphs
        for para in doc.paragraphs:
            if para.text.strip(): # Process only if paragraph has text
                # Store original runs and their text
                original_runs_text = [(run.text, run.style.name if run.style else None, run.font.bold, run.font.italic, run.font.underline) for run in para.runs]
                full_para_text = "".join([r[0] for r in original_runs_text])
                
                redacted_para_text = redact_names_in_text(full_para_text, names_to_redact, placeholder)

                if full_para_text != redacted_para_text:
                    # Clear existing runs
                    for i in range(len(para.runs)):
                        p = para._p
                        p.remove(para.runs[0]._r)
                    
                    # Add new run with redacted text, trying to preserve basic formatting
                    # This is a simplified approach. Complex formatting might be lost.
                    # A more robust way would be to iterate runs, redact, and reconstruct.
                    # However, if a name spans multiple runs, it gets very complex.
                    # For now, we'll replace the whole paragraph's text in a new run.
                    # A slightly better approach for preserving some formatting:
                    new_run = para.add_run(redacted_para_text)
                    if original_runs_text: # Try to apply style of first original run
                        first_run_props = original_runs_text[0]
                        # new_run.style = first_run_props[1] # Style application can be tricky
                        new_run.bold = first_run_props[2]
                        new_run.italic = first_run_props[3]
                        new_run.underline = first_run_props[4]

        # Redact in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.text.strip():
                            original_runs_text = [(run.text, run.style.name if run.style else None, run.font.bold, run.font.italic, run.font.underline) for run in para.runs]
                            full_para_text = "".join([r[0] for r in original_runs_text])
                            redacted_para_text = redact_names_in_text(full_para_text, names_to_redact, placeholder)
                            if full_para_text != redacted_para_text:
                                for i in range(len(para.runs)):
                                    p = para._p
                                    p.remove(para.runs[0]._r)
                                new_run = para.add_run(redacted_para_text)
                                if original_runs_text:
                                    first_run_props = original_runs_text[0]
                                    new_run.bold = first_run_props[2]
                                    new_run.italic = first_run_props[3]
                                    new_run.underline = first_run_props[4]


        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing DOCX: {e}")
        return None

def redact_pdf(file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
    """
    Redacts names from the first two pages of a PDF file.
    """
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        num_pages_to_process = min(2, doc.page_count)

        for page_num in range(num_pages_to_process):
            page = doc.load_page(page_num)
            for name in names_to_redact:
                if not name.strip():
                    continue
                # PyMuPDF's search is case-sensitive by default.
                # To make it case-insensitive, we'd need to search for variations or extract text first.
                # For simplicity here, we'll use its default. For true case-insensitivity,
                # you might extract text, find positions, then redact, or use regex on extracted text.
                # A common approach is to search for lowercase and uppercase versions if needed.
                # Here, we rely on the regex pattern for the placeholder content.
                
                # Using text instances for redaction
                text_instances = page.search_for(name, quads=True) # quads=True gives precise location
                for inst in text_instances:
                    # Add redaction annotation
                    annot = page.add_redact_annot(inst, text=placeholder, fill=(0,0,0)) # fill=(0,0,0) for black
            
            # Apply all redactions on the page. This is crucial.
            # It actually removes the text and replaces it.
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE) # Don't redact images

        output_buffer = io.BytesIO()
        doc.save(output_buffer, garbage=4, deflate=True, clean=True)
        doc.close()
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
        return None

def redact_pptx(file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
    """
    Redacts names from the first two slides of a PPTX file.
    """
    try:
        prs = Presentation(io.BytesIO(file_bytes))
        num_slides_to_process = min(2, len(prs.slides))

        for i in range(num_slides_to_process):
            slide = prs.slides[i]
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    # Store original runs and their text/formatting
                    original_runs_data = []
                    for run in paragraph.runs:
                        original_runs_data.append({
                            "text": run.text,
                            "bold": run.font.bold,
                            "italic": run.font.italic,
                            "underline": run.font.underline,
                            "name": run.font.name,
                            "size": run.font.size,
                            "color": run.font.color.rgb if run.font.color.rgb else None
                        })
                    
                    full_para_text = "".join([r['text'] for r in original_runs_data])
                    redacted_para_text = redact_names_in_text(full_para_text, names_to_redact, placeholder)

                    if full_para_text != redacted_para_text:
                        # Clear existing runs in the paragraph
                        p = paragraph._p
                        for idx in range(len(p.r_lst)):
                            p.remove(p.r_lst[0])
                        
                        # Add new run with redacted text.
                        # This simplified approach loses original run-level formatting if a name
                        # is redacted. A more complex approach would involve reconstructing runs.
                        new_run = paragraph.add_run()
                        new_run.text = redacted_para_text
                        # Try to apply formatting from the first original run if available
                        if original_runs_data:
                            first_run_format = original_runs_data[0]
                            new_run.font.bold = first_run_format["bold"]
                            new_run.font.italic = first_run_format["italic"]
                            new_run.font.underline = first_run_format["underline"]
                            if first_run_format["name"]: new_run.font.name = first_run_format["name"]
                            if first_run_format["size"]: new_run.font.size = first_run_format["size"]
                            # Color needs careful handling (RGBColor object)
                            # if first_run_format["color"]: new_run.font.color.rgb = RGBColor(*first_run_format["color"])

        output_buffer = io.BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing PPTX: {e}")
        return None

# --- Streamlit App ---
st.set_page_config(layout="wide")
st.title("📄 Document Name Redactor 🤫")
st.markdown("""
Upload a Word (.docx), PDF (.pdf), or PowerPoint (.pptx) file.
Enter the names to redact (comma-separated).
The tool will attempt to redact these names, primarily focusing on the **first two pages/slides**.
""")

# Input: Names to redact
st.sidebar.subheader("Names to Redact")
names_input = st.sidebar.text_area(
    "Enter names, separated by commas:",
    "Prof. Dumbledore, Albus Dumbledore, Minerva McGonagall, Severus Snape, Tom Riddle",
    height=150
)
names_to_redact = [name.strip() for name in names_input.split(',') if name.strip()]

# Input: Placeholder text
st.sidebar.subheader("Redaction Placeholder")
placeholder_text = st.sidebar.text_input("Text to replace names with:", "[REDACTED NAME]")


# Input: File uploader
uploaded_file = st.file_uploader(
    "Choose a file (DOCX, PDF, PPTX)",
    type=["docx", "pdf", "pptx"]
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_name = uploaded_file.name
    file_extension = os.path.splitext(file_name)[1].lower()

    st.markdown(f"**Uploaded file:** `{file_name}`")

    if not names_to_redact:
        st.warning("Please enter at least one name to redact in the sidebar.")
    else:
        if st.button("Redact Names 🕵️‍♀️"):
            redacted_file_bytes = None
            output_file_name = f"redacted_{file_name}"

            with st.spinner(f"Redacting names from {file_name}..."):
                if file_extension == ".docx":
                    redacted_file_bytes = redact_docx(file_bytes, names_to_redact, placeholder_text)
                elif file_extension == ".pdf":
                    redacted_file_bytes = redact_pdf(file_bytes, names_to_redact, placeholder_text)
                elif file_extension == ".pptx":
                    redacted_file_bytes = redact_pptx(file_bytes, names_to_redact, placeholder_text)
                else:
                    st.error("Unsupported file type.")

            if redacted_file_bytes:
                st.success("Redaction complete! 🎉")
                st.download_button(
                    label=f"Download Redacted File ({output_file_name})",
                    data=redacted_file_bytes,
                    file_name=output_file_name,
                    mime=uploaded_file.type  # Use the original mime type
                )
                st.info("ℹ️ Always review the redacted document carefully to ensure all desired information has been properly redacted and no unintended information was removed or formatting excessively altered.")

else:
    st.info("Upload a file to begin the redaction process.")

st.markdown("---")
st.markdown("Created with ❤️ by a Streamlit enthusiast.")
