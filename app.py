import streamlit as st
import docx
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
# from pptx.enum.shapes import MSO_SHAPE_TYPE # Keep for potential future recursion for group shapes
import fitz  # PyMuPDF
import io
import re
import os
import traceback

# --- Helper Functions for Redaction ---

def redact_names_in_text(text, names_to_redact, placeholder="[REDACTED NAME]"):
    redacted_text = text
    for name in names_to_redact:
        if not name.strip():
            continue
        # Ensure name from list is also stripped for the pattern
        pattern = r'\b' + re.escape(name.strip()) + r'\b'
        redacted_text = re.sub(pattern, placeholder, redacted_text, flags=re.IGNORECASE)
    return redacted_text

# --- PPTX Text Frame Processing Helper ---
def _process_text_frame_pptx(text_frame, names_to_redact, placeholder, debug_prefix=""):
    """
    Helper to process paragraphs within a given PowerPoint text_frame.
    Returns True if any redaction occurred in this text_frame.
    """
    redaction_occurred_in_frame = False
    for para_idx, paragraph in enumerate(text_frame.paragraphs):
        original_runs_data = []
        for run in paragraph.runs:
            font_color_type = None
            font_color_val = None
            if hasattr(run.font.color, 'type'):
                if run.font.color.type == MSO_COLOR_TYPE.RGB:
                    font_color_type = MSO_COLOR_TYPE.RGB
                    font_color_val = run.font.color.rgb
                elif run.font.color.type == MSO_COLOR_TYPE.SCHEME:
                    font_color_type = MSO_COLOR_TYPE.SCHEME
                    font_color_val = run.font.color.theme_color
            original_runs_data.append({
                "text": run.text, "bold": run.font.bold, "italic": run.font.italic,
                "underline": run.font.underline, "name": run.font.name, "size": run.font.size,
                "color_type": font_color_type, "color_val": font_color_val
            })

        full_para_text = "".join([r['text'] for r in original_runs_data])
        if not full_para_text.strip():
            continue

        # --- DEBUG ---
        # st.write(f"{debug_prefix} Para {para_idx} Text: '{full_para_text}'")
        # -------------

        redacted_para_text = redact_names_in_text(full_para_text, names_to_redact, placeholder)

        if full_para_text != redacted_para_text:
            st.info(f"{debug_prefix} Para {para_idx}: REDACTING '{full_para_text[:30]}...' TO '{redacted_para_text[:30]}...'") # DEBUG
            redaction_occurred_in_frame = True
            
            p = paragraph._p
            for _ in range(len(p.r_lst)):
                p.remove(p.r_lst[0])
            
            new_run = paragraph.add_run()
            new_run.text = redacted_para_text
            
            if original_runs_data:
                first_run_format = original_runs_data[0]
                if first_run_format["bold"] is not None: new_run.font.bold = first_run_format["bold"]
                if first_run_format["italic"] is not None: new_run.font.italic = first_run_format["italic"]
                if first_run_format["underline"] is not None: # True, False, None, or Enum
                    try: new_run.font.underline = first_run_format["underline"]
                    except: pass # If direct assignment fails for some complex underline type
                if first_run_format["name"]: new_run.font.name = first_run_format["name"]
                if first_run_format["size"]: new_run.font.size = first_run_format["size"]
                
                if first_run_format["color_type"] == MSO_COLOR_TYPE.RGB and first_run_format["color_val"]:
                    new_run.font.color.rgb = RGBColor(*first_run_format["color_val"])
                elif first_run_format["color_type"] == MSO_COLOR_TYPE.SCHEME and first_run_format["color_val"] is not None:
                    try: new_run.font.color.theme_color = first_run_format["color_val"]
                    except: pass
        # else: # DEBUG
        #     st.write(f"{debug_prefix} Para {para_idx}: NO REDACTION for: '{full_para_text[:50]}...'") # DEBUG
            
    return redaction_occurred_in_frame


def redact_docx(file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        elements_to_process = []
        for para in doc.paragraphs:
            elements_to_process.append(para)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para_in_cell in cell.paragraphs: # Corrected variable name
                        elements_to_process.append(para_in_cell)

        redaction_occurred_overall = False
        st.write(f"DOCX: Names to redact: {names_to_redact}") # DEBUG

        for para_idx, para in enumerate(elements_to_process):
            if not para.text.strip():
                continue
            
            original_runs_data = []
            for run in para.runs:
                original_runs_data.append({
                    "text": run.text, "bold": run.font.bold,
                    "italic": run.font.italic, "underline": run.font.underline,
                })
            
            full_para_text = "".join([r['text'] for r in original_runs_data])
            if not full_para_text.strip():
                continue

            # st.write(f"DOCX Para {para_idx} Text: '{full_para_text}'") # DEBUG

            redacted_para_text = redact_names_in_text(full_para_text, names_to_redact, placeholder)

            if full_para_text != redacted_para_text:
                st.info(f"DOCX Para {para_idx}: REDACTING '{full_para_text[:30]}...' TO '{redacted_para_text[:30]}...'") # DEBUG
                redaction_occurred_overall = True
                
                for _ in range(len(para.runs)):
                    p_element = para._p
                    p_element.remove(para.runs[0]._r)
                
                new_run = para.add_run(redacted_para_text)
                
                if original_runs_data:
                    first_run_props = original_runs_data[0]
                    if first_run_props["bold"] is not None: new_run.font.bold = first_run_props["bold"]
                    if first_run_props["italic"] is not None: new_run.font.italic = first_run_props["italic"]
                    if first_run_props["underline"] is not None:
                        try: new_run.font.underline = first_run_props["underline"]
                        except: pass
        
        if not redaction_occurred_overall:
            st.warning("DOCX: No text was identified for redaction. Check input names and file content. Review debug output above if any.")

        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing DOCX: {e}")
        st.error(traceback.format_exc())
        return None

def redact_pdf(file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        num_pages_to_process = min(2, doc.page_count)
        redaction_occurred_overall = False
        st.write(f"PDF: Names to redact: {names_to_redact}") # DEBUG

        for page_num in range(num_pages_to_process):
            page = doc.load_page(page_num)
            page_redacted = False
            for name in names_to_redact:
                if not name.strip():
                    continue
                
                text_instances = page.search_for(name, quads=True)
                if text_instances:
                    st.info(f"PDF Page {page_num+1}: Found '{name}', adding redaction annotations.") # DEBUG
                    redaction_occurred_overall = True
                    page_redacted = True
                for inst in text_instances:
                    annot = page.add_redact_annot(inst, text=placeholder, fill=(0,0,0)) 
            
            if page_redacted: # Apply redactions only if something was added to this page
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        if not redaction_occurred_overall:
            st.warning("PDF: No text was identified for redaction. Check input names and file content. Note: PDF search is case-sensitive by default in PyMuPDF for the `search_for` method.")

        output_buffer = io.BytesIO()
        doc.save(output_buffer, garbage=4, deflate=True, clean=True)
        doc.close()
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
        st.error(traceback.format_exc())
        return None

def redact_pptx(file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
    try:
        prs = Presentation(io.BytesIO(file_bytes))
        num_slides_to_process = min(2, len(prs.slides))
        st.write(f"PPTX: Processing {num_slides_to_process} slides.")
        st.write(f"PPTX: Names to redact: {names_to_redact}") # DEBUG

        redaction_occurred_somewhere = False

        for i in range(num_slides_to_process):
            slide = prs.slides[i]
            # st.write(f"PPTX: --- Processing Slide {i+1} ---") # DEBUG

            # Process shapes directly on the slide
            for shape_idx, shape in enumerate(slide.shapes):
                debug_shape_prefix = f"PPTX Slide {i+1} Shape {shape_idx}"
                if shape.has_text_frame:
                    # st.write(f"{debug_shape_prefix}: Has direct text frame.") # DEBUG
                    if _process_text_frame_pptx(shape.text_frame, names_to_redact, placeholder, f"{debug_shape_prefix} DirectTF"):
                        redaction_occurred_somewhere = True
                
                if shape.has_table:
                    # st.write(f"{debug_shape_prefix}: Has table.") # DEBUG
                    table = shape.table
                    for row_idx, row in enumerate(table.rows):
                        for col_idx, cell in enumerate(row.cells):
                            if cell.text_frame:
                                # st.write(f"{debug_shape_prefix} Table R{row_idx}C{col_idx}: Has text frame.") # DEBUG
                                if _process_text_frame_pptx(cell.text_frame, names_to_redact, placeholder, f"{debug_shape_prefix} TableR{row_idx}C{col_idx}"):
                                    redaction_occurred_somewhere = True
                
                # To handle grouped shapes, you'd check shape.shape_type == MSO_SHAPE_TYPE.GROUP
                # and then recursively call a function on shape.shapes.
                # from pptx.enum.shapes import MSO_SHAPE_TYPE (at top)
                # if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                #   st.write(f"{debug_shape_prefix}: Is a GROUP. Recursing...") # DEBUG
                #   # Need a recursive helper function here for shape.shapes collection
                #   # For now, this is omitted for simplicity but is a known extension point.

        if not redaction_occurred_somewhere:
            st.warning("PPTX: No text was identified for redaction. Check input names and file content. Review debug output above if any.")
        
        output_buffer = io.BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error processing PPTX: {e}")
        st.error(traceback.format_exc()) 
        return None

# --- Streamlit App ---
st.set_page_config(layout="wide")
st.title("📄 Document Name Redactor 🤫")
st.markdown("""
Upload a Word (.docx), PDF (.pdf), or PowerPoint (.pptx) file.
Enter the names to redact (comma-separated).
The tool will attempt to redact these names, primarily focusing on the **first two pages/slides**.
**Note:** Review debug messages below the button after processing, especially if redaction doesn't seem to work as expected.
""")

# Input: Names to redact
st.sidebar.subheader("Names to Redact")
names_input = st.sidebar.text_area(
    "Enter names, separated by commas (case-insensitive match):",
    "Prof. Dumbledore, Albus Dumbledore, Minerva McGonagall, Severus Snape, Tom Riddle",
    height=150
)
raw_names = [name.strip() for name in names_input.split(',') if name.strip()]
# Filter out very short names to avoid accidental redactions if desired, e.g. names less than 3 chars
names_to_redact = [name for name in raw_names if len(name) > 2] 
if len(raw_names) != len(names_to_redact):
    st.sidebar.caption(f"Note: {len(raw_names) - len(names_to_redact)} short name(s) (<=2 chars) were ignored.")


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
        st.warning("Please enter at least one name (3+ characters) to redact in the sidebar.")
    else:
        if st.button("Redact Names 🕵️‍♀️"):
            st.markdown("---") # Separator for debug output
            st.subheader("Redaction Process Log:")
            redacted_file_bytes = None
            output_file_name = f"redacted_{file_name}"

            with st.spinner(f"Redacting names from {file_name}... This may take a moment."):
                if file_extension == ".docx":
                    redacted_file_bytes = redact_docx(file_bytes, names_to_redact, placeholder_text)
                elif file_extension == ".pdf":
                    redacted_file_bytes = redact_pdf(file_bytes, names_to_redact, placeholder_text)
                elif file_extension == ".pptx":
                    redacted_file_bytes = redact_pptx(file_bytes, names_to_redact, placeholder_text)
                else:
                    st.error("Unsupported file type.")

            if redacted_file_bytes:
                st.success("Redaction attempt complete! 🎉")
                st.download_button(
                    label=f"Download Redacted File ({output_file_name})",
                    data=redacted_file_bytes,
                    file_name=output_file_name,
                    mime=uploaded_file.type
                )
                st.info("ℹ️ Always review the redacted document carefully to ensure all desired information has been properly redacted and no unintended information was removed or formatting excessively altered. Check the log above for details.")
            else:
                st.error("Redaction process did not produce an output file. See logs above for details.")
else:
    st.info("Upload a file to begin the redaction process.")

st.markdown("---")
st.markdown("Created with ❤️ by a Streamlit enthusiast.")
