import streamlit as st
import fitz  # PyMuPDF
import io
import re
import os
import subprocess # For attempting LibreOffice/unoconv if available
import tempfile
import traceback
import shutil

# --- Configuration ---
# On Streamlit Cloud, we probably CANNOT rely on a custom LIBREOFFICE_PATH.
# We'll try common command names.
CONVERSION_COMMAND_OPTIONS = [shutil.which("soffice"), shutil.which("libreoffice"), shutil.which("unoconv")]
CONVERTER_COMMAND = next((cmd for cmd in CONVERSION_COMMAND_OPTIONS if cmd is not None), None)

# --- Helper Functions ---

def redact_names_in_text(text, names_to_redact, placeholder="[REDACTED NAME]"):
    redacted_text = text
    for name in names_to_redact:
        if not name.strip():
            continue
        pattern = r'\b' + re.escape(name.strip()) + r'\b'
        redacted_text = re.sub(pattern, placeholder, redacted_text, flags=re.IGNORECASE)
    return redacted_text

def attempt_conversion_to_pdf(input_file_path, output_directory, original_file_extension):
    """
    Attempts to convert a document to PDF using available command-line tools.
    Returns the path to the converted PDF file or None on error.
    """
    if not CONVERTER_COMMAND:
        st.warning("No suitable office conversion command (soffice, libreoffice, unoconv) found in the environment. DOCX/PPTX conversion might fail or be unavailable.")
        # Try to see if specific python libraries might be an option (more complex to integrate here)
        # For now, we'll just indicate failure if no command is found.
        if original_file_extension == ".docx":
            st.info("For .docx, consider libraries like 'docx2pdf' if you can ensure its dependencies are met.")
        elif original_file_extension == ".pptx":
            st.info("For .pptx, pure Python PDF conversion is very challenging.")
        return None

    base_name_without_ext = os.path.splitext(os.path.basename(input_file_path))[0]
    # Ensure the output name is just .pdf, not .pptx.pdf etc.
    converted_pdf_name = f"{base_name_without_ext}.pdf"
    converted_pdf_path = os.path.join(output_directory, converted_pdf_name)

    cmd_list = []
    if "unoconv" in CONVERTER_COMMAND:
        # unoconv syntax: unoconv -f pdf -o output.pdf input.docx
        cmd_list = [CONVERTER_COMMAND, "-f", "pdf", "-o", converted_pdf_path, input_file_path]
    elif "soffice" in CONVERTER_COMMAND or "libreoffice" in CONVERTER_COMMAND:
        # soffice syntax: soffice --headless --convert-to pdf input.docx --outdir /output/directory/
        cmd_list = [CONVERTER_COMMAND, "--headless", "--convert-to", "pdf", input_file_path, "--outdir", output_directory]
    else: # Should not happen if CONVERTER_COMMAND is set
        st.error("Internal error: Converter command identified but syntax unknown.")
        return None

    try:
        st.write(f"Attempting conversion with: {' '.join(cmd_list)}")
        process = subprocess.run(
            cmd_list,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=90  # Increased timeout for potentially slower cloud environments
        )

        # LibreOffice/soffice creates output file with original name + .pdf in outdir
        # unoconv creates output file specified by -o
        # So we always check for `converted_pdf_path`
        
        if os.path.exists(converted_pdf_path): # Primary check
             st.write(f"Successfully converted to: {converted_pdf_path}")
             return converted_pdf_path
        elif process.returncode == 0 and not os.path.exists(converted_pdf_path) and ("soffice" in CONVERTER_COMMAND or "libreoffice" in CONVERTER_COMMAND):
            # Soffice/LibreOffice might have created input_file_name.pdf directly in outdir
            # This logic is slightly redundant due to how converted_pdf_name is constructed but good as a fallback check
            potential_soffice_output = os.path.join(output_directory, f"{os.path.splitext(os.path.basename(input_file_path))[0]}.pdf")
            if os.path.exists(potential_soffice_output):
                st.write(f"Successfully converted (found as soffice output): {potential_soffice_output}")
                return potential_soffice_output # Return the actual path found
            else:
                st.error(f"Conversion command (exit code 0) but expected PDF not found at '{converted_pdf_path}' or '{potential_soffice_output}'.")
                st.error(f"Stdout: {process.stdout.decode('utf-8', errors='ignore')}")
                st.error(f"Stderr: {process.stderr.decode('utf-8', errors='ignore')}")
                return None
        elif process.returncode != 0 :
            st.error(f"Conversion command failed with exit code {process.returncode}.")
            st.error(f"Stdout: {process.stdout.decode('utf-8', errors='ignore')}")
            st.error(f"Stderr: {process.stderr.decode('utf-8', errors='ignore')}")
            return None
        else: # process.returncode == 0 but file still not found (e.g. unoconv didn't error but didn't create)
            st.error(f"Conversion command (exit code 0) but output PDF not found at: {converted_pdf_path}")
            st.error(f"Stdout: {process.stdout.decode('utf-8', errors='ignore')}")
            st.error(f"Stderr: {process.stderr.decode('utf-8', errors='ignore')}")
            return None

    except subprocess.TimeoutExpired:
        st.error("Conversion command timed out.")
        return None
    except Exception as e:
        st.error(f"An error occurred during conversion: {e}")
        st.error(traceback.format_exc())
        return None


def redact_pdf_bytes(pdf_file_bytes, names_to_redact, placeholder="[REDACTED NAME]"):
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
                    st.info(f"PDF Page {page_num+1}: Found '{name}', adding redaction.")
                    redaction_occurred_overall = True
                    page_redacted_this_pass = True
                for inst in text_instances:
                    page.add_redact_annot(inst, text="", fill=(0,0,0))

            if page_redacted_this_pass:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        if not redaction_occurred_overall:
            st.warning("PDF Redaction: No text was identified for redaction within the first two pages. Check input names and file content. Note: PDF search is case-sensitive.")

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
st.title("📄 Document Name Redactor 🤫")

st.markdown("""
Upload a Word (.docx), PowerPoint (.pptx), or PDF (.pdf) file.
The application will **attempt to convert DOCX and PPTX files to PDF** using available system tools,
then redact the specified names from the **first two pages** of the (converted) PDF.
**Note:** DOCX/PPTX conversion success depends on the capabilities of the deployment environment.
If conversion fails, please convert your file to PDF manually and re-upload.
""")

if not CONVERTER_COMMAND:
    st.warning(
        "**Conversion Tool Note:** No standard office conversion command (like 'soffice' or 'unoconv') "
        "was automatically found in this environment. DOCX/PPTX to PDF conversion may not work. "
        "Uploading pre-converted PDF files is recommended."
    )
else:
    st.info(f"Attempting to use converter: {CONVERTER_COMMAND}")


st.sidebar.subheader("Names to Redact")
names_input = st.sidebar.text_area(
    "Enter names, separated by commas:",
    "Prof. Dumbledore, Albus Dumbledore, Minerva McGonagall",
    height=150
)
raw_names = [name.strip() for name in names_input.split(',') if name.strip()]
names_to_redact = [name for name in raw_names if len(name) > 2]
if len(raw_names) != len(names_to_redact):
    st.sidebar.caption(f"Note: {len(raw_names) - len(names_to_redact)} short name(s) (<=2 chars) were ignored.")

placeholder_text = "[REDACTED]" # Not visually used on PDF redaction box

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
        st.warning("Please enter at least one name (3+ characters) to redact.")
    else:
        if st.button("Process and Redact Names 🕵️‍♀️"):
            st.markdown("---")
            st.subheader("Processing Log:")
            redacted_output_bytes = None
            output_file_name_base = os.path.splitext(original_file_name)[0]
            final_output_filename = f"redacted_{output_file_name_base}.pdf"

            with st.spinner(f"Processing {original_file_name}..."):
                pdf_to_redact_bytes = None

                if original_file_extension in [".docx", ".pptx"]:
                    st.write(f"Detected {original_file_extension.upper()} file. Attempting conversion to PDF...")
                    with tempfile.TemporaryDirectory() as temp_dir:
                        temp_original_file_path = os.path.join(temp_dir, original_file_name) # Use original name for clarity
                        with open(temp_original_file_path, "wb") as f_temp:
                            f_temp.write(original_file_bytes)

                        converted_pdf_path = attempt_conversion_to_pdf(temp_original_file_path, temp_dir, original_file_extension)

                        if converted_pdf_path and os.path.exists(converted_pdf_path):
                            st.write(f"Conversion successful. Reading converted PDF: {converted_pdf_path}")
                            with open(converted_pdf_path, "rb") as f_pdf:
                                pdf_to_redact_bytes = f_pdf.read()
                        else:
                            st.error("Failed to convert the document to PDF with available tools. Please convert manually to PDF and re-upload.")
                elif original_file_extension == ".pdf":
                    st.write("Detected PDF file. Proceeding directly to redaction.")
                    pdf_to_redact_bytes = original_file_bytes
                else:
                    st.error("Unsupported file type.") # Should not happen due to file_uploader types

                if pdf_to_redact_bytes:
                    st.write("Starting PDF redaction process...")
                    redacted_output_bytes = redact_pdf_bytes(pdf_to_redact_bytes, names_to_redact, placeholder_text)

            if redacted_output_bytes:
                st.success("Processing complete! 🎉 The output is a PDF.")
                st.download_button(
                    label=f"Download Redacted PDF ({final_output_filename})",
                    data=redacted_output_bytes,
                    file_name=final_output_filename,
                    mime="application/pdf"
                )
                st.info("ℹ️ Always review the redacted PDF carefully.")
            else:
                st.error("Processing did not produce an output file. See logs above.")
else:
    st.info("Upload a file to begin.")

st.markdown("---")
st.markdown("If DOCX/PPTX conversion fails, please convert your file to PDF manually and upload the PDF.")
