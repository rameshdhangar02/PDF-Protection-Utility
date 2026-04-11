
import streamlit as st
import fitz  # PyMuPDF
import io
import zipfile

st.set_page_config(page_title="PDF Locker Utility", layout="centered")

st.title("📄 PDF Image Converter & Locker")
st.write("Converts PDFs to high-quality images and applies security restrictions to prevent browser copying.")

uploaded_files = st.file_uploader(
    "Choose PDF files", 
    type="pdf", 
    accept_multiple_files=True
)

def convert_and_lock_pdf(file_bytes):
    # Open original PDF
    doc = fitz.open("pdf", file_bytes)
    out_pdf = fitz.open()
    
    for page in doc:
        # 1. Render page to high-res image (300 DPI for crisp text)
        pix = page.get_pixmap(dpi=300)
        
        # 2. Create a new blank page in the output PDF with the exact same dimensions
        new_page = out_pdf.new_page(width=page.rect.width, height=page.rect.height)
        
        # 3. Paste the rendered image exactly over the blank page
        new_page.insert_image(page.rect, pixmap=pix)
        
    # 4. Apply PDF Restrictions (Allow printing, deny copying)
    perms = fitz.PDF_PERM_PRINT 
    
    out_bytes = io.BytesIO()
    
    # Save with strong owner password to enforce the "No Copying" rule
    out_pdf.save(
        out_bytes, 
        encryption=fitz.PDF_ENCRYPT_AES_256, 
        owner_pw="locked_admin_password_123", 
        user_pw="", 
        permissions=perms
    )
    
    return out_bytes.getvalue()

if uploaded_files:
    st.divider()
    st.subheader("Processing Files")
    
    processed_files = []

    for uploaded_file in uploaded_files:
        with st.spinner(f"Processing {uploaded_file.name}..."):
            try:
                # Perform conversion and locking
                output_pdf = convert_and_lock_pdf(uploaded_file.read())
                
                output_name = f"locked_{uploaded_file.name}"
                processed_files.append({"name": output_name, "data": output_pdf})
                
                st.success(f"Done: {uploaded_file.name}")
                st.download_button(
                    label=f"Download {output_name}",
                    data=output_pdf,
                    file_name=output_name,
                    mime="application/pdf",
                    key=output_name
                )
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")

    # Bulk Download for multiple files
    if len(processed_files) > 1:
        st.divider()
        st.subheader("Bulk Download")
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for file in processed_files:
                zf.writestr(file["name"], file["data"])
        
        st.download_button(
            label="Download All as ZIP",
            data=zip_buffer.getvalue(),
            file_name="locked_pdfs.zip",
            mime="application/zip"
        )