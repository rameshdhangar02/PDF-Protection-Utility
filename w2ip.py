import streamlit as st
import fitz  # PyMuPDF
import io
import zipfile
import os
import tempfile
import platform
import subprocess
import shutil
import gc
import re

st.set_page_config(page_title="PDF Locker Utility", page_icon="📄", layout="centered")

st.title("📄 Document Image Converter & Locker")
st.write("Converts Word/PDF/Image files to high-quality protected PDFs with restrictions against printing, copying, and browser extraction.")

input_mode = st.radio("Select Input Mode", ["Web Upload (Files or ZIP)", "Local Folder Path (Direct Processing)"])

def format_size(num_bytes):
    """Helper to format bytes to human-readable size"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if num_bytes < 1024.0:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} GB"

def process_file_pipeline(file_name, file_bytes):
    """Handles word-to-pdf conversion and the pdf locker pipeline"""
    if os.path.basename(file_name).startswith("~$"):
        return None
        
    try:
        file_ext = file_name.rsplit('.', 1)[-1].lower() if '.' in file_name else ""
        
        # Whitelist of valid extensions
        if file_ext not in ['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'bmp', 'tiff']:
            return None
        
        # Convert Word docs to PDF bytes first
        if file_ext in ['doc', 'docx']:
            file_bytes = convert_word_to_pdf_bytes(file_bytes, file_name)
            if file_bytes is None:
                return None
            file_ext = "pdf"
            
        if file_bytes is None or len(file_bytes) == 0:
            return None
            
        result = convert_and_lock_pdf(file_bytes, file_ext)
        return result
    except Exception as e:
        st.error(f"Error processing {file_name}: {e}")
        return None
    finally:
        gc.collect()

def convert_word_to_pdf_bytes(file_bytes, filename):
    """Converts Word document to PDF using COM on Windows or headless LibreOffice on Linux"""
    with tempfile.TemporaryDirectory() as tmpdir:
        safe_filename = os.path.basename(filename)
        input_path = os.path.join(tmpdir, safe_filename)
        pdf_name = safe_filename.rsplit('.', 1)[0] + ".pdf"
        output_path = os.path.join(tmpdir, pdf_name)
        
        with open(input_path, "wb") as f:
            f.write(file_bytes)
            
        is_windows = os.name == 'nt' or platform.system() == 'Windows'
        
        if is_windows:
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                
                word = None
                try:
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = False 
                    
                    abs_in = os.path.abspath(input_path)
                    abs_out = os.path.abspath(output_path)
                    
                    doc = word.Documents.Open(abs_in, ReadOnly=True)
                    doc.SaveAs(abs_out, FileFormat=17)  # 17 = wdFormatPDF
                    doc.Close(SaveChanges=False)
                    
                    if os.path.exists(abs_out):
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
            except ImportError:
                st.error("pywin32 COM library is missing. Please run 'pip install pywin32'")
                return None
        else:
            libreoffice_bin = shutil.which("libreoffice") or shutil.which("soffice")
            if not libreoffice_bin:
                st.error("LibreOffice is not installed on this server. Please ensure `libreoffice` is in `packages.txt`.")
                return None
                
            try:
                profile_dir = os.path.join(tmpdir, "libo_profile")
                env = os.environ.copy()
                env["SAL_USE_VCLPLUGIN"] = "svp"
                
                cmd = [
                    libreoffice_bin,
                    "--headless",
                    "--invisible",
                    "--nologo",
                    "--nofirststartwizard",
                    "--norestore",
                    "--nodefault",
                    f"-env:UserInstallation=file://{profile_dir}",
                    "--convert-to", "pdf",
                    input_path,
                    "--outdir", tmpdir
                ]
                process = subprocess.run(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    env=env,
                    timeout=60
                )
                
                if process.returncode != 0:
                    st.error(f"LibreOffice conversion failed. Error: {process.stderr.decode('utf-8', errors='ignore')}")
                    return None
                    
                if os.path.exists(output_path):
                    with open(output_path, "rb") as pdf_file:
                        return pdf_file.read()
                else:
                    candidates = [f for f in os.listdir(tmpdir) if f.endswith(".pdf")]
                    if candidates:
                        with open(os.path.join(tmpdir, candidates[0]), "rb") as pdf_file:
                            return pdf_file.read()
                    st.error("LibreOffice completed but output PDF not found.")
                    return None
            except subprocess.TimeoutExpired:
                st.error("LibreOffice conversion timed out after 60 seconds.")
                return None
            except Exception as e:
                st.error(f"Error during Linux Word-to-PDF conversion: {e}")
                return None

def convert_and_lock_pdf(file_bytes, file_ext="pdf"):
    """Rasterizes all pages to optimized images and locks PDF with strict AES-256 permissions"""
    doc = None
    out_pdf = None
    try:
        doc = fitz.open(stream=file_bytes, filetype=file_ext)
        out_pdf = fitz.open()
        
        total_pages = len(doc)
        if total_pages <= 5:
            calculated_dpi = 144
        elif total_pages <= 20:
            calculated_dpi = 120
        elif total_pages <= 40:
            calculated_dpi = 96
        else:
            calculated_dpi = 72
        
        for page in doc:
            # 1. Render page to image with dynamic DPI balancing crisp quality and memory
            pix = page.get_pixmap(dpi=calculated_dpi)
            
            # 2. Convert to compressed JPEG stream (quality=75 saves ~70% RAM & size vs default 95)
            img_bytes = pix.tobytes("jpg", jpg_quality=75)
            del pix  # free C++ pixmap buffer immediately
            
            # 3. Create a blank page in the output PDF matching original dimensions
            new_page = out_pdf.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=img_bytes)
            del img_bytes  # free JPEG buffer immediately
            
        # 4. Apply strict PDF Restrictions (0 = block all: no print, no copy, no extract)
        perms = 0
        
        # Anti-browser warning javascript
        js_code = """
        var vT = app.viewerType;
        if(vT !== "Reader" && vT !== "Exchange") {
            app.alert("Warning: This secured document is not supported in browser PDF viewers. Please use Adobe Acrobat Reader for full secure viewing.");
        }
        """
        try:
            out_pdf.set_open_action(js_code)
        except:
            pass
        
        out_bytes = io.BytesIO()
        out_pdf.save(
            out_bytes, 
            deflate=True,
            garbage=3,
            clean=True,
            encryption=fitz.PDF_ENCRYPT_AES_256, 
            owner_pw="locked_admin_password_123", 
            user_pw="", 
            permissions=perms
        )
        
        result = out_bytes.getvalue()
        del out_bytes
        return result
    finally:
        if doc is not None:
            try:
                doc.close()
            except:
                pass
        if out_pdf is not None:
            try:
                out_pdf.close()
            except:
                pass
        gc.collect()

if input_mode == "Web Upload (Files or ZIP)":
    # Dedicated temp directory per session on disk (prevents memory blowup on Streamlit Cloud)
    if "work_dir" not in st.session_state or not os.path.exists(st.session_state.get("work_dir", "")):
        st.session_state["work_dir"] = tempfile.mkdtemp(prefix="pdf_locker_")
    work_dir = st.session_state["work_dir"]

    col1, col2 = st.columns([4, 1])
    with col1:
        uploaded_files = st.file_uploader(
            "Choose Document files or a ZIP Folder", 
            type=["pdf", "docx", "doc", "jpg", "jpeg", "png", "bmp", "tiff", "zip"], 
            accept_multiple_files=True
        )
    with col2:
        st.write("")
        st.write("")
        if st.button("🧹 Clear All"):
            shutil.rmtree(work_dir, ignore_errors=True)
            st.session_state["work_dir"] = tempfile.mkdtemp(prefix="pdf_locker_")
            st.session_state.pop("processed_keys", None)
            st.session_state.pop("processed_files", None)
            st.session_state.pop("zip_path", None)
            st.session_state.pop("zip_name", None)
            gc.collect()
            st.rerun()

    if uploaded_files:
        current_upload_keys = [f"{f.name}_{f.size}" for f in uploaded_files]
        
        if "processed_keys" not in st.session_state or st.session_state["processed_keys"] != current_upload_keys:
            st.divider()
            st.subheader("⚙️ Processing Files...")
            
            # Clean work_dir of previous files
            for f in os.listdir(work_dir):
                try:
                    os.remove(os.path.join(work_dir, f))
                except:
                    pass
            
            processed_file_meta = []
            
            for uploaded_file in uploaded_files:
                file_ext = uploaded_file.name.rsplit('.', 1)[-1].lower() if '.' in uploaded_file.name else ""
                
                if file_ext == "zip":
                    with st.spinner(f"Extracting & processing ZIP: '{uploaded_file.name}'..."):
                        try:
                            with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
                                valid_files = [n for n in z.namelist() if not os.path.basename(n).startswith("~$") and '.' in n and n.rsplit('.', 1)[-1].lower() in ['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'bmp', 'tiff']]
                                total_files = len(valid_files)
                                
                                if total_files > 0:
                                    progress_bar = st.progress(0, text=f"Processing 0/{total_files} files inside ZIP...")
                                    for i, z_name in enumerate(valid_files):
                                        safe_name = os.path.basename(z_name)
                                        progress_bar.progress((i + 1) / total_files, text=f"Processing {i+1}/{total_files}: {safe_name}")
                                        file_bytes = z.read(z_name)
                                        out_pdf_bytes = process_file_pipeline(safe_name, file_bytes)
                                        del file_bytes
                                        
                                        if out_pdf_bytes:
                                            base_name = safe_name.rsplit('.', 1)[0]
                                            out_filename = f"{base_name}.pdf"
                                            out_path = os.path.join(work_dir, out_filename)
                                            with open(out_path, "wb") as pf:
                                                pf.write(out_pdf_bytes)
                                            file_size = len(out_pdf_bytes)
                                            del out_pdf_bytes
                                            processed_file_meta.append({
                                                "name": out_filename,
                                                "path": out_path,
                                                "size": format_size(file_size)
                                            })
                                        gc.collect()
                                    progress_bar.empty()
                        except Exception as e:
                            st.error(f"Failed to process ZIP archive: {e}")
                else:
                    with st.spinner(f"Processing '{uploaded_file.name}'..."):
                        file_bytes = uploaded_file.read()
                        out_pdf_bytes = process_file_pipeline(uploaded_file.name, file_bytes)
                        del file_bytes
                        if out_pdf_bytes:
                            base_name = uploaded_file.name.rsplit('.', 1)[0]
                            out_filename = f"{base_name}.pdf"
                            out_path = os.path.join(work_dir, out_filename)
                            with open(out_path, "wb") as pf:
                                pf.write(out_pdf_bytes)
                            file_size = len(out_pdf_bytes)
                            del out_pdf_bytes
                            processed_file_meta.append({
                                "name": out_filename,
                                "path": out_path,
                                "size": format_size(file_size)
                            })
                        gc.collect()

            # Pre-generate ZIP on disk ONCE (avoids rebuilding in RAM on every rerun!)
            if len(processed_file_meta) > 0:
                if len(uploaded_files) == 1 and uploaded_files[0].name.lower().endswith(".zip"):
                    zip_base = uploaded_files[0].name.rsplit('.', 1)[0]
                    zip_name = f"{zip_base}_secured.zip"
                else:
                    zip_name = "secured_documents.zip"
                
                zip_path = os.path.join(work_dir, zip_name)
                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for item in processed_file_meta:
                        zf.write(item["path"], arcname=item["name"])
                
                st.session_state["zip_path"] = zip_path
                st.session_state["zip_name"] = zip_name
            
            st.session_state["processed_keys"] = current_upload_keys
            st.session_state["processed_files"] = processed_file_meta
            gc.collect()
            st.rerun()

        # Render Download Outputs (Lightweight, instant loading)
        processed_files = st.session_state.get("processed_files", [])
        zip_path = st.session_state.get("zip_path", None)
        zip_name = st.session_state.get("zip_name", "secured_documents.zip")

        if processed_files:
            st.divider()
            st.success(f"✅ Successfully converted & locked {len(processed_files)} file(s)!")
            
            # Primary Action: Download All ZIP
            if zip_path and os.path.exists(zip_path):
                zip_size = os.path.getsize(zip_path)
                with open(zip_path, "rb") as zf:
                    st.download_button(
                        label=f"📦 Download All as ZIP ({zip_name} • {format_size(zip_size)})",
                        data=zf.read(),
                        file_name=zip_name,
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
            
            # Secondary Action: Individual file downloads inside a clean expander
            if len(processed_files) > 1:
                with st.expander("📄 Download Individual PDF Files"):
                    for p_file in processed_files:
                        if os.path.exists(p_file["path"]):
                            with open(p_file["path"], "rb") as pf:
                                st.download_button(
                                    label=f"⬇️ {p_file['name']} ({p_file['size']})",
                                    data=pf.read(),
                                    file_name=p_file["name"],
                                    mime="application/pdf",
                                    key=f"dl_{p_file['name']}"
                                )
            elif len(processed_files) == 1:
                p_file = processed_files[0]
                if os.path.exists(p_file["path"]):
                    with open(p_file["path"], "rb") as pf:
                        st.download_button(
                            label=f"⬇️ Download {p_file['name']} ({p_file['size']})",
                            data=pf.read(),
                            file_name=p_file["name"],
                            mime="application/pdf",
                            key=f"dl_{p_file['name']}",
                            use_container_width=True
                        )

elif input_mode == "Local Folder Path (Direct Processing)":
    if os.name != 'nt' and platform.system() != 'Windows':
        st.warning("⚠️ Local folder paths are only accessible when running this app locally on your computer. When running on Streamlit Cloud, please select **'Web Upload (Files or ZIP)'** above.")
    else:
        st.info("Since this app runs locally, you can paste an absolute path to a folder on your computer. It will process all valid files and store them in an output folder matching your folder name.")
    folder_path = st.text_input("Enter absolute Folder Path (e.g., C:/Users/.../Documents/MyFolder)")
    
    if st.button("Process Folder"):
        cleaned_path = folder_path.strip(' \t\n\r"\'')
        
        # UNC path handling for Windows
        if cleaned_path.startswith("//"):
            cleaned_path = "\\\\" + cleaned_path[2:].replace("/", "\\")
        elif re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}[/\\]", cleaned_path):
            cleaned_path = "\\\\" + cleaned_path.replace("/", "\\")

        if os.path.exists(cleaned_path) and os.path.isdir(cleaned_path):
            norm_path = os.path.normpath(cleaned_path)
            folder_name = os.path.basename(norm_path)
            if not folder_name:
                folder_parts = [p for p in norm_path.split('\\') if p]
                folder_name = folder_parts[-1] if folder_parts else "network_folder"
                
            out_folder = os.path.join(os.path.dirname(norm_path), f"new_{folder_name}")
            os.makedirs(out_folder, exist_ok=True)
            
            valid_exts = ['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'bmp', 'tiff']
            
            all_valid_paths = []
            for root, dirs, files in os.walk(cleaned_path):
                for file in files:
                    if not file.startswith("~$") and '.' in file and file.rsplit('.', 1)[-1].lower() in valid_exts:
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
                        del file_bytes
                        
                        if out_pdf:
                            base_name = file.rsplit('.', 1)[0]
                            output_filename = f"{base_name}.pdf"
                            
                            rel_dir = os.path.relpath(root, cleaned_path)
                            target_dir = os.path.join(out_folder, rel_dir)
                            os.makedirs(target_dir, exist_ok=True)
                            
                            output_path = os.path.join(target_dir, output_filename)
                            with open(output_path, "wb") as f:
                                f.write(out_pdf)
                            del out_pdf
                            processed_count += 1
                        gc.collect()
                    except Exception as e:
                        st.warning(f"Failed to process {file}: {e}")
                
                progress_bar.empty()
            else:
                st.info("No valid files found in this directory.")
            st.success(f"✅ Successfully processed {processed_count} files!")
            st.info(f"📁 Output stored directly at: {out_folder}")
        else:
            st.error("Invalid folder path. Please ensure the directory exists.")