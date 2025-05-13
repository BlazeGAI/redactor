import streamlit as st
import fitz  # PyMuPDF
import io
import re
import os
import subprocess # For calling LibreOffice
import tempfile # For managing temporary files
import traceback # For detailed error messages
import shutil # For finding executable

# --- Configuration for LibreOffice ---
# Try to find soffice automatically, but allow override if needed
LIBREOFFICE_PATH = shutil.which("soffice") or shutil.which("libreoffice")
# If on Windows and not in PATH, you might need to set this manually:
# LIBREOFFICE_PATH = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
# If on macOS and not in PATH, you might need something like:
# LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

# --- Helper Functions ---

def redact_names_in_text(text, names_to_redact, placeholder="[REDACTED NAME]"):
    """
    Redacts a list of names from a given text string.
    Uses regex for whole-word, case-insensitive matching.
    """
    redacted_text = text
    for name in names_to_redact:
        if not name.strip():
            continue
        pattern = r'\b' + re.escape(name.strip()) + r'\b'
        redacted_text = re.sub(pattern, placeholder, redacted_text, flags=re.IGNORECASE)
    return redacted_text

def convert_to_pdf_libreoffice(input_file_path, output_directory):
    """
    Converts a document to PDF using LibreOffice.
    Returns the path to the converted PDF file or None on error.
    """
    if not LIBREOFFICE_PATH:
        st.error("LibreOffice 'soffice' command not found. Please ensure LibreOffice is installed and in your system's PATH.")
        return None

    base_name_without_ext = os.path.splitext(os.path.basename(input_file_path))[0]
    converted_pdf_name = f"{base_name_without_ext}.pdf"
    converted_pdf_path = os.path.join(output_directory, converted_pdf_name)

    try:
        st.write(f"Attempting conversion: {LIBREOFFICE_PATH} --headless --convert-to pdf \"{input_file_path}\" --outdir \"{output_directory}\"")
        process = subprocess.run(
            [LIBREOFFICE_PATH, "--headless", "--convert-to", "pdf", input_file_path, "--outdir", output_directory],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=60  # Timeout after 60 seconds
        )
        if process.returncode == 0:
            # Check if the expected PDF file was created
            if os.path.exists(converted_pdf_path):
                st.write(f"Successfully converted to: {converted_pdf_path}")
                return converted_pdf_path
            else:
                st.error(f"LibreOffice conversion seemed successful (exit code 0) but output PDF not found at: {converted_pdf_path}")
                st.error(f"LibreOffice stdout: {process.stdout.decode('utf-8', errors='ignore')}")
                st.error(f"LibreOffice stderr: {process.stderr.decode('utf-8', errors='ignore')}")
                return None
        else:
            st.error(f"LibreOffice conversion failed with exit code {process.returncode}.")
            st.error(f"Input file: {input_file_path}")
            st.error(f"Output directory: {output_directory}")
            st.error(f"LibreOffice stdout: {process.stdout.decode('utf-8', errors='ignore')}")
            st.error(f"LibreOffice stderr: {process.stderr.decode('utf-8', errors='ignore')}")
            return None
    except subprocess.TimeoutExpired:
        st.error("LibreOffice conversion timed out.")
        return None
    except Exception as e:
        st.error(f"An error occurred during LibreOffice conversion: {e}")
        st.error(traceback.format_exc())
        return None

def redact_pdf_bytes(pdf_file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
    """
    Redacts names from the first two pages of a PDF file (provided as bytes).
    """
    try:
        doc = fitz.open(stream=pdf_file_bytes, filetype="pdf")
        num_pages_to_process = min(2, doc.page_count)
        redaction_occurred_overall = False
        st.write(f"PDF Redaction: Names to redact: {names_to_redact}")

        for page_num in range(num_pages_to_process):
            page = doc.load_page(page_num)
            page_redacted_this_pass = False
            for name in names_to_redact:
                if not name.strip():
                    continue
                
                text_instances = page.search_for(name, quads=True)
                if text_instances:
                    st.info(f"PDF Page {page_num+1}: Found '{name}', adding redaction annotations.")
                    redaction_occurred_overall = True
                    page_redacted_this_pass = True
                for inst in text_instances:
                    # Add redaction annotation with placeholder text
                    # PyMuPDF doesn't directly overlay text on the redaction box by default in the same way
                    # you might fill a rectangle with text. The `text` parameter in `add_redact_annot`
                    # is for the 'Contents' entry of the annotation, not visible text.
                    # To put visible text, we'd fill the area, then add text.
                    # For simplicity, we'll use the standard black box redaction.
                    # If placeholder text *on* the redaction is critical, it's more complex.
                    page.add_redact_annot(inst, text="", fill=(0,0,0)) # Black fill, no overlay text

            if page_redacted_this_pass:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        if not redaction_occurred_overall:
            st.warning("PDF Redaction: No text was identified for redaction within the first two pages. Check input names and file content. Note: PDF search is case-sensitive by default in PyMuPDF.")

        output_buffer = io.BytesIO()
        doc.save(output_buffer, garbage=4, deflate=True, clean=True)
        doc.close()
        output_buffer.seek(0)
        return output_buffer
    except Exception as e:
        st.error(f"Error during PDF redaction: {e}")
        st.error(traceback.format_exc())
        return None

# --- Streamlit App ---
st.set_page_config(layout="wide")
st.title("📄 Universal Document Name Redactor (via PDF) 🤫")

st.markdown("""
Upload a Word (.docx), PowerPoint (.pptx), or PDF (.pdf) file.
The application will **convert DOCX and PPTX files to PDF** using LibreOffice first,
then attempt to redact the specified names from the **first two pages** of the (converted) PDF.
""")

if not LIBREOFFICE_PATH:
    st.error(
        "**Critical Setup Error:** LibreOffice (soffice) was not found. "
        "DOCX and PPTX conversion will fail. Please install LibreOffice and ensure "
        "'soffice' or 'libreoffice' is in your system's PATH, or configure `LIBREOFFICE_PATH` in the script."
    )
else:
    st.success(f"Using LibreOffice found at: {LIBREOFFICE_PATH}")


# Input: Names to redact
st.sidebar.subheader("Names to Redact")
names_input = st.sidebar.text_area(
    "Enter names, separated by commas (case-insensitive match for redaction):",
    "Prof. Dumbledore, Albus Dumbledore, Minerva McGonagall, Severus Snape, Tom Riddle",
    height=150
)
raw_names = [name.strip() for name in names_input.split(',') if name.strip()]
names_to_redact = [name for name in raw_names if len(name) > 2]
if len(raw_names) != len(names_to_redact):
    st.sidebar.caption(f"Note: {len(raw_names) - len(names_to_redact)} short name(s) (<=2 chars) were ignored.")

# Input: Placeholder text (Note: PyMuPDF redaction primarily blacks out, visible placeholder is harder)
st.sidebar.subheader("Redaction Style")
# placeholder_text = st.sidebar.text_input("Placeholder text (visual effect varies):", "[REDACTED]")
st.sidebar.info("Names will be redacted with a black box. Direct text overlay on redaction is complex with PyMuPDF.")
placeholder_text = "[REDACTED]" # Kept for consistency, but less visually prominent in PDF


# Input: File uploader
uploaded_file = st.file_uploader(
    "Choose a file (DOCX, PDF, PPTX)",
    type=["docx", "pdf", "pptx"]
)

if uploaded_file is not None:
    original_file_bytes = uploaded_file.getvalue()
    original_file_name = uploaded_file.name
    original_file_extension = os.path.splitext(original_file_name)[1].lower()

    st.markdown(f"**Uploaded file:** `{original_file_name}`")

    if not names_to_redact:
        st.warning("Please enter at least one name (3+ characters) to redact in the sidebar.")
    else:
        if st.button("Convert and Redact Names 🕵️‍♀️"):
            st.markdown("---")
            st.subheader("Processing Log:")
            redacted_output_bytes = None
            output_file_name_base = os.path.splitext(original_file_name)[0]
            final_output_filename = f"redacted_{output_file_name_base}.pdf" # Output is always PDF

            with st.spinner(f"Processing {original_file_name}... This may take a moment."):
                pdf_to_redact_bytes = None

                if original_file_extension in [".docx", ".pptx"]:
                    st.write(f"Detected {original_file_extension.upper()} file. Attempting conversion to PDF...")
                    with tempfile.TemporaryDirectory() as temp_dir:
                        temp_original_file_path = os.path.join(temp_dir, original_file_name)
                        with open(temp_original_file_path, "wb") as f_temp:
                            f_temp.write(original_file_bytes)

                        converted_pdf_path = convert_to_pdf_libreoffice(temp_original_file_path, temp_dir)

                        if converted_pdf_path and os.path.exists(converted_pdf_path):
                            st.write(f"Conversion successful. Reading converted PDF: {converted_pdf_path}")
                            with open(converted_pdf_path, "rb") as f_pdf:
                                pdf_to_redact_bytes = f_pdf.read()
                        else:
                            st.error("Failed to convert the document to PDF. Cannot proceed with redaction.")
                            pdf_to_redact_bytes = None # Ensure it's None
                elif original_file_extension == ".pdf":
                    st.write("Detected PDF file. Proceeding directly to redaction.")
                    pdf_to_redact_bytes = original_file_bytes
                else:
                    st.error("Unsupported file type.")
                    pdf_to_redact_bytes = None

                if pdf_to_redact_bytes:
                    st.write("Starting PDF redaction process...")
                    redacted_output_bytes = redact_pdf_bytes(pdf_to_redact_bytes, names_to_redact, placeholder_text)

            if redacted_output_bytes:
                st.success("Redaction attempt complete! 🎉 The output is a PDF.")
                st.download_button(
                    label=f"Download Redacted PDF ({final_output_filename})",
                    data=redacted_output_bytes,
                    file_name=final_output_filename,
                    mime="application/pdf"
                )
                st.info("ℹ️ Always review the redacted PDF carefully. Conversion and redaction might alter layout or miss some occurrences.")
            else:
                st.error("Redaction process did not produce an output file. See logs above for details.")
else:
    st.info("Upload a file to begin the conversion and redaction process.")

st.markdown("---")
st.markdown("Created with ❤️ by a Streamlit enthusiast.")
