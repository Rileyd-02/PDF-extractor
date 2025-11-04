import streamlit as st
import tabula
import pandas as pd
import io
import zipfile
import subprocess

# ----------------------- #
# Streamlit Page Settings #
# ----------------------- #
st.set_page_config(
    page_title="📄 PDF to Excel Converter",
    page_icon="📑",
    layout="centered"
)

st.title("📄 PDF to Excel Converter")
st.markdown(
    """
    Upload one or more **PDF files** and convert them into **Excel spreadsheets** automatically.  
    The app detects and extracts all tables — even from unstructured PDFs — and saves each one into a separate Excel sheet.  
    """
)

# ----------------------- #
# Utility Functions       #
# ----------------------- #

def check_java():
    """Verify that Java is installed (required for tabula)."""
    try:
        subprocess.run(["java", "-version"], check=True, capture_output=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


def convert_pdf_to_excel(pdf_file):
    """
    Reads tables from a PDF and returns a BytesIO buffer containing the Excel file.
    Automatically switches between lattice and stream modes for best results.
    """
    buffer = io.BytesIO()

    if not check_java():
        st.error("🚫 Java is not installed or not configured correctly.")
        st.info("If you're deploying on Streamlit Cloud, add a `packages.txt` file containing:\n`openjdk-11-jdk`")
        return None

    try:
        # Try lattice mode first (for bordered tables)
        try:
            tables = tabula.read_pdf(pdf_file, pages="all", multiple_tables=True, lattice=True)
        except Exception:
            tables = []

        # If no tables found, try stream mode
        if not tables:
            tables = tabula.read_pdf(pdf_file, pages="all", multiple_tables=True, stream=True)

        # Handle empty case
        if not tables:
            st.warning(f"No tables detected in `{pdf_file.name}`. Creating a blank Excel file.")
            tables = [pd.DataFrame({"Message": ["No tables found in PDF."]})]

        # Write all tables to Excel
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            for i, df in enumerate(tables):
                df = df.dropna(how="all").dropna(how="all", axis=1)
                if not df.empty:
                    df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"⚠️ Error processing `{pdf_file.name}`: {e}")
        df_err = pd.DataFrame({"Error": [str(e)]})
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df_err.to_excel(writer, sheet_name="Error", index=False)
        buffer.seek(0)
        return buffer


# ----------------------- #
# File Upload Section     #
# ----------------------- #
uploaded_files = st.file_uploader(
    "📂 Choose one or more PDF files",
    type="pdf",
    accept_multiple_files=True,
    help="You can upload multiple PDFs; the app will combine results into a ZIP file."
)

if st.button("🚀 Convert to Excel"):
    if not uploaded_files:
        st.warning("Please upload at least one PDF to start conversion.")
    else:
        with st.spinner("Processing your PDFs..."):
            excel_files = {}
            for pdf in uploaded_files:
                result = convert_pdf_to_excel(pdf)
                if result:
                    excel_files[pdf.name.replace(".pdf", ".xlsx")] = result.getvalue()

        # ----------------------- #
        # Download Section        #
        # ----------------------- #
        if excel_files:
            st.success("✅ Conversion complete!")

            if len(excel_files) == 1:
                name, data = list(excel_files.items())[0]
                st.download_button(
                    label=f"📥 Download {name}",
                    data=data,
                    file_name=name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for name, data in excel_files.items():
                        zipf.writestr(name, data)
                st.download_button(
                    label="📦 Download All as ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="converted_pdfs.zip",
                    mime="application/zip"
                )


# ----------------------- #
# Footer / Info Section   #
# ----------------------- #
st.markdown("---")
st.markdown(
    """
    ### ℹ️ About
    - Extracts tables from any PDF (bordered or unbordered).  
    - Each table is exported to a new Excel sheet.  
    - Built using **Streamlit** + **tabula-py** + **pandas**.  
    """
)
st.caption("Built with ❤️ by Lahiru • Powered by Streamlit + tabula-py")
