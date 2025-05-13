import streamlit as st
import docx
from pptx import Presentation
from pptx.dml.color import RGBColor # MODIFIED: For setting RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE # MODIFIED: For checking color type
# from pptx.enum.dml import MSO_THEME_COLOR_INDEX # For type hinting if needed, but direct use is fine
import fitz  # PyMuPDF
import io
import re
import os
import traceback # For detailed error messages

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
        
        elements_to_process = list(doc.paragraphs) # Process top-level paragraphs
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    elements_to_process.extend(cell.paragraphs) # Add paragraphs from cells

        for para in elements_to_process: # Iterate through combined list
            if para.text.strip(): 
                original_runs_text_and_format = []
                for run in para.runs:
                    original_runs_text_and_format.append({
                        "text": run.text,
                        "bold": run.font.bold,
                        "italic": run.font.italic,
                        "underline": run.font.underline,
                        # Add other formatting as needed
                    })
                
                full_para_text = "".join([r['text'] for r in original_runs_text_and_format])
                redacted_para_text = redact_names_in_text(full_para_text, names_to_redact, placeholder)

                if full_para_text != redacted_para_text:
                    # Clear existing runs
                    for i in range(len(para.runs)):
                        p = para._p
                        p.remove(para.runs[0]._r)
                    
                    new_run = para.add_run(redacted_para_text)
                    # Try to apply formatting from the first original run if available
                    if original_runs_text_and_format:
                        first_run_props = original_runs_text_and_format[0]
                        if first_run_props["bold"] is not None: new_run.font.bold = first_run_props["bold"]
                        if first_run_props["italic"] is not None: new_run.font.italic = first_run_props["italic"]
                        if first_run_props["underline"] is not None: new_run.font.underline = first_run_props["underline"]
        
        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing DOCX: {e}")
        st.error(traceback.format_exc())
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
                
                text_instances = page.search_for(name, quads=True) 
                for inst in text_instances:
                    annot = page.add_redact_annot(inst, text=placeholder, fill=(0,0,0)) 
            
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE) 

        output_buffer = io.BytesIO()
        doc.save(output_buffer, garbage=4, deflate=True, clean=True)
        doc.close()
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
        st.error(traceback.format_exc())
        return None

# MODIFIED redact_pptx function
def redact_pptx(file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
    """
    Redacts names from the first two slides of a PPTX file.
    Handles different color types (RGB vs. Scheme).
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
                    original_runs_data = []
                    for run in paragraph.runs:
                        font_color_type = None
                        font_color_val = None
                        
                        # Check font.color.type before accessing specific color properties
                        if hasattr(run.font.color, 'type'): # Defensive check
                            if run.font.color.type == MSO_COLOR_TYPE.RGB:
                                font_color_type = MSO_COLOR_TYPE.RGB
                                font_color_val = run.font.color.rgb
                            elif run.font.color.type == MSO_COLOR_TYPE.SCHEME:
                                font_color_type = MSO_COLOR_TYPE.SCHEME
                                font_color_val = run.font.color.theme_color # Store the theme_color enum
                            # Add other MSO_COLOR_TYPE checks if needed (e.g., PRESET, CMYK)
                        
                        original_runs_data.append({
                            "text": run.text,
                            "bold": run.font.bold,
                            "italic": run.font.italic,
                            "underline": run.font.underline, # Can be True, False, None, or MSO_TEXT_UNDERLINE_TYPE
                            "name": run.font.name,
                            "size": run.font.size,
                            "color_type": font_color_type,
                            "color_val": font_color_val
                        })
                    
                    full_para_text = "".join([r['text'] for r in original_runs_data])
                    redacted_para_text = redact_names_in_text(full_para_text, names_to_redact, placeholder)

                    if full_para_text != redacted_para_text:
                        # Clear existing runs in the paragraph
                        p = paragraph._p
                        for idx in range(len(p.r_lst)):
                            p.remove(p.r_lst[0])
                        
                        new_run = paragraph.add_run()
                        new_run.text = redacted_para_text
                        
                        if original_runs_data:
                            first_run_format = original_runs_data[0]
                            if first_run_format["bold"] is not None: new_run.font.bold = first_run_format["bold"]
                            if first_run_format["italic"] is not None: new_run.font.italic = first_run_format["italic"]
                            if first_run_format["underline"] is not None: new_run.font.underline = first_run_format["underline"]
                            
                            if first_run_format["name"]: new_run.font.name = first_run_format["name"]
                            if first_run_format["size"]: new_run.font.size = first_run_format["size"]
                            
                            if first_run_format["color_type"] == MSO_COLOR_TYPE.RGB and first_run_format["color_val"]:
                                new_run.font.color.rgb = RGBColor(first_run_format["color_val"][0], first_run_format["color_val"][1], first_run_format["color_val"][2])
                            elif first_run_format["color_type"] == MSO_COLOR_TYPE.SCHEME and first_run_format["color_val"] is not None:
                                try:
                                    # first_run_format["color_val"] should be an MSO_THEME_COLOR_INDEX enum member
                                    new_run.font.color.theme_color = first_run_format["color_val"]
                                except Exception as e_theme:
                                    # Log or silently pass if theme color application fails for some reason
                                    # print(f"Note: Could not re-apply theme color: {e_theme}")
                                    pass
        
        output_buffer = io.BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing PPTX: {e}")
        st.error(traceback.format_exc()) # Provides full traceback for debugging
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
                    mime=uploaded_file.type
                )
                st.info("ℹ️ Always review the redacted document carefully to ensure all desired information has been properly redacted and no unintended information was removed or formatting excessively altered.")
else:
    st.info("Upload a file to begin the redaction process.")

st.markdown("---")
st.markdown("Created with ❤️ by a Streamlit enthusiast.")
