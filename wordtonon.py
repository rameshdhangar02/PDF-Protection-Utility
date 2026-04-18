import streamlit as st
import fitz  # PyMuPDF
import io
import zipfile
import os
import tempfile
import platform
import subprocess

st.set_page_config(page_title="PDF Locker Utility", layout="centered")

st.title("📄 Document Image Converter & Locker (Cross-Platform)")
st.write("Converts Word/PDF files to high-quality images and applies strict security restrictions to block browser support.")

input_mode = st.radio("Select Input Mode", ["Web Upload (Files or ZIP)", "Local Folder Path (Direct Processing)"])

def process_file_pipeline(file_name, file_bytes):
    """Handles word-to-pdf conversion and the pdf locker pipeline"""
    # Ignore Word ghost/temporary locking files
    if os.path.basename(file_name).startswith("~$"):
        return None
        
    try:
        file_ext = file_name.split('.')[-1].lower()
        
        # Enforce global extension whitelist to ignore rogue txt, excel, etc.
        if file_ext not in ['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'bmp', 'tiff']:
            return None
        
        # Word docs converting to PDF Stream first
        if file_ext in ['doc', 'docx']:
            file_bytes = convert_word_to_pdf_bytes(file_bytes, file_name)
            if file_bytes is None:
                return None
            file_ext = "pdf"
            
        return convert_and_lock_pdf(file_bytes, file_ext)
    except Exception as e:
        st.error(f"Error processing {file_name}: {e}")
        return None

def convert_word_to_pdf_bytes(file_bytes, filename):
    with tempfile.TemporaryDirectory() as tmpdir:
        # Sanitize filename to prevent nested path crashes when extracting from zip structures
        safe_filename = os.path.basename(filename)
        input_path = os.path.join(tmpdir, safe_filename)
        pdf_name = safe_filename.rsplit('.', 1)[0] + ".pdf"
        output_path = os.path.join(tmpdir, pdf_name)
        
        with open(input_path, "wb") as f:
            f.write(file_bytes)
            
        if os.name == 'nt' or platform.system() == 'Windows':
            # Local Windows execution via COM
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
            except ImportError:
                st.error("pywin32 COM library is missing. Please run 'pip install pywin32'")
                return None
                
            word = None
            try:
                # Force isolated backend execution to perfectly drop Word.Application.Documents collisions
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = False 
                
                # Word backend requires strictly valid absolute directories 
                abs_in = os.path.abspath(input_path)
                abs_out = os.path.abspath(output_path)
                
                doc = word.Documents.Open(abs_in, ReadOnly=True)
                doc.SaveAs(abs_out, FileFormat=17) # 17 translates to wdFormatPDF
                doc.Close(SaveChanges=False)
                
                with open(abs_out, "rb") as pdf_file:
                    return pdf_file.read()
            except Exception as e:
                st.error(f"Failed to convert Word to PDF core: {e}")
                return None
            finally:
                if word:
                    try:
                        word.Quit()
                    except:
                        pass
        else:
            # Linux execution using libreoffice headless
            try:
                # Command: libreoffice --headless --convert-to pdf <input_file> --outdir <output_dir>
                # Streamlit usually aliases to 'libreoffice' or 'soffice'
                process = subprocess.run(
                    ["libreoffice", "--headless", "--nologo", "--nofirststartwizard", "--convert-to", "pdf", input_path, "--outdir", tmpdir],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
                
                if process.returncode != 0:
                    # Some Linux distros use 'soffice'
                    process = subprocess.run(
                        ["soffice", "--headless", "--nologo", "--nofirststartwizard", "--convert-to", "pdf", input_path, "--outdir", tmpdir],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE
                    )
                
                if process.returncode != 0:
                    st.error(f"LibreOffice conversion failed. Error: {process.stderr.decode('utf-8')}")
                    return None
                    
                if os.path.exists(output_path):
                    with open(output_path, "rb") as pdf_file:
                        return pdf_file.read()
                else:
                    st.error("LibreOffice completed but output PDF not found.")
                    return None
            except FileNotFoundError:
                st.error("LibreOffice is not installed on this system. Make sure you added `libreoffice` to your `packages.txt` file on Streamlit Cloud.")
                return None
            except Exception as e:
                st.error(f"Error during Linux Word-to-PDF conversion: {e}")
                return None


def convert_and_lock_pdf(file_bytes, file_ext="pdf"):
    # Open original PDF or Image
    doc = fitz.open(file_ext, file_bytes)
    out_pdf = fitz.open()
    
    total_pages = len(doc)
    # Target DPI calculation to constrain file size under ~5MB.
    # JPEG compression at ~200 DPI gives ~250-400KB/page.
    calculated_dpi = 200
    if total_pages > 10: calculated_dpi = 144
    if total_pages > 30: calculated_dpi = 100
    if total_pages > 50: calculated_dpi = 72
    
    for page in doc:
        # 1. Render page to image with dynamic DPI balancing quality and file size
        pix = page.get_pixmap(dpi=calculated_dpi)
        
        # 2. Create a new blank page in the output PDF with the exact same dimensions
        new_page = out_pdf.new_page(width=page.rect.width, height=page.rect.height)
        
        # 3. Convert image to a highly compressed JPG stream and paste it 
        img_bytes = pix.tobytes("jpg")
        new_page.insert_image(page.rect, stream=img_bytes)
        
    # 4. Apply PDF Restrictions - 0 means strictly deny all options including Print and Copy 
    # to prevent bypasses from internet browsers support
    perms = 0 
    
    # Adding JavaScript to show warning in viewers other than Adobe Reader as an extra anti-browser effort
    js_code = \"\"\"
    var vT = app.viewerType;
    if(vT !== "Reader" && vT !== "Exchange") {
        app.alert("Warning: This secured document is not supported in browser PDF viewers. Please use Adobe Acrobat Reader for full secure viewing.");
    }
    \"\"\"
    try:
        # Some versions of PyMuPDF support setting js actions
        out_pdf.set_open_action(js_code)
    except:
        pass
    
    out_bytes = io.BytesIO()
    
    # Save with strong owner password to enforce strictly "No Copying, No Printing"
    # Also deflating and cleaning unused objects to compress the final PDF
    out_pdf.save(
        out_bytes, 
        deflate=True,
        garbage=3,
        encryption=fitz.PDF_ENCRYPT_AES_256, 
        owner_pw="locked_admin_password_123", 
        user_pw="", 
        permissions=perms
    )
    
    return out_bytes.getvalue()

if input_mode == "Web Upload (Files or ZIP)":
    uploaded_files = st.file_uploader(
        "Choose Document files or a ZIP Folder", 
        type=["pdf", "docx", "doc", "jpg", "jpeg", "png", "bmp", "tiff", "zip"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        st.divider()
        st.subheader("Processing Files")
        
        current_upload_keys = [f"{f.name}_{f.size}" for f in uploaded_files]
        
        if "processed_keys" not in st.session_state or st.session_state["processed_keys"] != current_upload_keys:
            processed_files = []
        
            for uploaded_file in uploaded_files:
                file_ext = uploaded_file.name.split('.')[-1].lower()
                
                if file_ext == "zip":
                    # Process contents of the ZIP folder
                    with st.spinner(f"Unpacking & Processing ZIP folder '{uploaded_file.name}'..."):
                        try:
                            with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
                                valid_files = [n for n in z.namelist() if n.split('.')[-1].lower() in ['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'bmp', 'tiff']]
                                total_files = len(valid_files)
                                
                                if total_files > 0:
                                    progress_bar = st.progress(0, text=f"Processing 0/{total_files} files inside ZIP...")
                                    
                                    for i, z_name in enumerate(valid_files):
                                        progress_bar.progress((i + 1) / total_files, text=f"Processing {i+1}/{total_files}: {os.path.basename(z_name)}")
                                        file_bytes = z.read(z_name)
                                        out_pdf = process_file_pipeline(os.path.basename(z_name), file_bytes)
                                        if out_pdf:
                                            base_name = z_name.rsplit('.', 1)[0]
                                            processed_files.append({"name": f"{base_name}.pdf", "data": out_pdf})
                                    
                                    progress_bar.empty()
                        except Exception as e:
                            st.error(f"Failed to read ZIP contents: {e}")
                else:
                    # Process singular files
                    with st.spinner(f"Processing {uploaded_file.name}..."):
                        out_pdf = process_file_pipeline(uploaded_file.name, uploaded_file.read())
                        if out_pdf:
                            base_name = uploaded_file.name.rsplit('.', 1)[0]
                            output_name = f"{base_name}.pdf"
                            processed_files.append({"name": output_name, "data": out_pdf})
                            
                            st.success(f"Done: {uploaded_file.name}")
            
            st.session_state["processed_keys"] = current_upload_keys
            st.session_state["processed_files"] = processed_files
        else:
            processed_files = st.session_state["processed_files"]
            st.info("Files processed and cached successfully. Ready for download below!")
    
        # Render Download Outputs
        if len(processed_files) > 0:
            st.divider()
            st.subheader("Downloads")
            
            # Non-ZIP individual File Downloads
            if not any(f.name.lower().endswith('.zip') for f in uploaded_files):
                for p_file in processed_files:
                    st.download_button(
                        label=f"⬇️ Download {p_file['name']}",
                        data=p_file["data"],
                        file_name=p_file["name"],
                        mime="application/pdf",
                        key=f"dl_{p_file['name']}"
                    )
            
            # Output ZIP inherits folder/ZIP name logically
            if len(uploaded_files) == 1 and uploaded_files[0].name.lower().endswith(".zip"):
                folder_name = uploaded_files[0].name.rsplit('.', 1)[0]
                output_zip_name = f"{folder_name}.zip"
            else:
                output_zip_name = "documents.zip"
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for file in processed_files:
                    zf.writestr(file["name"], file["data"])
            
            st.download_button(
                label=f"Download All as {output_zip_name}",
                data=zip_buffer.getvalue(),
                file_name=output_zip_name,
                mime="application/zip"
            )

elif input_mode == "Local Folder Path (Direct Processing)":
    st.info("Since this app runs locally, you can paste an absolute path to a folder on your computer. It will process all valid files and store them in an output folder matching your folder name.")
    folder_path = st.text_input("Enter absolute Folder Path (e.g., C:/Users/.../Documents/MyFolder)")
    
    if st.button("Process Folder"):
        import re
        # Clean up the path: remove quotes and whitespaces that users often accidentally copy
        cleaned_path = folder_path.strip(' \t\n\r"\'')
        
        # If the user inputted a network url with forward slashes (e.g. //192.168.1.5/share)
        # convert it strictly to Windows UNC path format (\\\\192.168.1.5\\share)
        if cleaned_path.startswith("//"):
            cleaned_path = "\\\\\\\\" + cleaned_path[2:].replace("/", "\\\\")
        elif re.match(r"^\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}[/\\\\]", cleaned_path):
            # If someone typed an IP address directly without UNC prefixes (e.g. 192.168.1.5/share)
            cleaned_path = "\\\\\\\\" + cleaned_path.replace("/", "\\\\")

        if os.path.exists(cleaned_path) and os.path.isdir(cleaned_path):
            norm_path = os.path.normpath(cleaned_path)
            folder_name = os.path.basename(norm_path)
            if not folder_name: # Handle case where it's a root UNC path like \\\\192.168.1.5\\share
                folder_parts = [p for p in norm_path.split('\\\\') if p]
                folder_name = folder_parts[-1] if folder_parts else "network_folder"
                
            out_folder = os.path.join(os.path.dirname(norm_path), f"new_{folder_name}")
            
            os.makedirs(out_folder, exist_ok=True)
            
            valid_exts = ['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'bmp', 'tiff']
            
            all_valid_paths = []
            for root, dirs, files in os.walk(cleaned_path):
                for file in files:
                    if file.split('.')[-1].lower() in valid_exts:
                        all_valid_paths.append((root, file))
            
            total_files = len(all_valid_paths)
            processed_count = 0
            
            if total_files > 0:
                progress_bar = st.progress(0, text=f"Found {total_files} valid files. Starting processing...")
                
                for i, (root, file) in enumerate(all_valid_paths):
                    file_path = os.path.join(root, file)
                    progress_bar.progress((i + 1) / total_files, text=f"Processing {i+1}/{total_files}: {file}")
                    
                    try:
                        with open(file_path, "rb") as f:
                            file_bytes = f.read()
                            
                        out_pdf = process_file_pipeline(file, file_bytes)
                        
                        if out_pdf:
                            base_name = file.rsplit('.', 1)[0]
                            output_filename = f"{base_name}.pdf"
                            
                            # Replicate the nested directory structure strictly inside the new isolated output folder
                            rel_dir = os.path.relpath(root, cleaned_path)
                            target_dir = os.path.join(out_folder, rel_dir)
                            os.makedirs(target_dir, exist_ok=True)
                            
                            output_path = os.path.join(target_dir, output_filename)
                            
                            with open(output_path, "wb") as f:
                                f.write(out_pdf)
                                
                            processed_count += 1
                    except Exception as e:
                        st.warning(f"Failed to process {file}: {e}")
                
                progress_bar.empty()
            else:
                st.info("No valid files found in this directory.")
            st.success(f"✅ Successfully processed {processed_count} files!")
            st.info(f"📁 Output stored directly at: {out_folder}")
        else:
            st.error("Invalid folder path. Please ensure the directory exists.")
