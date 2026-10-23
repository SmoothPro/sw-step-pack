import os
import zipfile
import datetime
import win32com.client
import pythoncom

swDocPART = 1
swDocASSEMBLY = 2
swSaveAsCurrentVersion = 0
swSaveAsOptions_Silent = 1
swOpenDocOptions_Silent = 1
swOpenDocOptions_ReadOnly = 2

def connect_solidworks():
    try:
        sw_app = win32com.client.GetActiveObject("SldWorks.Application")
        print("[INFO] Successfully connected to running SOLIDWORKS.")
    except Exception:
        print("[INFO] No running SOLIDWORKS instance detected. Launching a new instance...")
        sw_app = win32com.client.Dispatch("SldWorks.Application")
        sw_app.Visible = True
    return sw_app

def convert_to_stp(sw_app, file_path, output_dir):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".sldprt":
        doc_type = swDocPART
    elif ext == ".sldasm":
        doc_type = swDocASSEMBLY
    else:
        return None

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{file_name}.stp")

    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    open_options = swOpenDocOptions_Silent | swOpenDocOptions_ReadOnly
    model_doc = sw_app.OpenDoc6(file_path, doc_type, open_options, "", errors, warnings)
    
    if not model_doc:
        print(f"[-] Failed to open file: {file_path} (Error code: {errors.value})")
        return None

    try:
        success = model_doc.SaveAs4(output_path, swSaveAsCurrentVersion, swSaveAsOptions_Silent, errors, warnings)
        if success:
            print(f"[+] Successfully exported: {file_name}.stp")
            return output_path
        else:
            print(f"[-] Failed to save: {file_name} (Error code: {errors.value})")
            return None
    finally:
        sw_app.CloseDoc(os.path.basename(file_path))

def create_zip(output_dir, zip_filename, files_to_zip):
    zip_path = os.path.join(output_dir, zip_filename)
    print(f"\n[Compressing] Packing files into: {zip_path}")
    
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zipf:
        for file in files_to_zip:
            if os.path.exists(file):
                zipf.write(file, arcname=os.path.basename(file))
                
    print(f"[Compression Complete] Created {zip_filename} (contains {len(files_to_zip)} files)")

def get_user_inputs():
    print("="*40)
    print(" SOLIDWORKS Part to STP")
    print("="*40)
    
    source_dir = input("[ Source Folder ]:\n>> ").strip()
    source_dir = source_dir.strip('"').strip("'")
    
    while not os.path.isdir(source_dir):
        print("[-] Error: Path not found. Please try again.")
        source_dir = input(">> ").strip().strip('"').strip("'")

    default_output = os.path.join(source_dir, "step")
    output_dir = input(f"\n[ Output Folder ]  Press Enter to use default: '{default_output}'):\n>> ").strip()
    output_dir = output_dir.strip('"').strip("'")
    
    if not output_dir:
        output_dir = default_output

    recursive_input = input("\n3. Include all subfolders? (y/N, press Enter for N):\n>> ").strip().lower()
    recursive = True if recursive_input == 'y' else False

    return source_dir, output_dir, recursive

def main():
    SOURCE_DIR, OUTPUT_DIR, RECURSIVE = get_user_inputs()

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    sw_app = connect_solidworks()

    target_files = []
    if RECURSIVE:
        for root, _, files in os.walk(SOURCE_DIR):
            for f in files:
                if f.lower().endswith((".sldprt", ".sldasm")):
                    target_files.append(os.path.join(root, f))
    else:
        for f in os.listdir(SOURCE_DIR):
            if f.lower().endswith((".sldprt", ".sldasm")):
                target_files.append(os.path.join(SOURCE_DIR, f))

    total = len(target_files)
    if total == 0:
        print("\n[End] No .sldprt or .sldasm files found.")
        return

    print(f"\n[Start Conversion] Found {total} CAD model files...")

    exported_stp_files = []
    for idx, file_path in enumerate(target_files, start=1):
        print(f"({idx}/{total}) Processing: {os.path.basename(file_path)}")
        stp_result = convert_to_stp(sw_app, file_path, OUTPUT_DIR)
        if stp_result:
            exported_stp_files.append(stp_result)

    print(f"\n[Conversion Complete] Successfully exported: {len(exported_stp_files)}/{total}")

    if exported_stp_files:
        default_name = f"STP_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        user_name = input(f"\nEnter [ ZIP file ]S name (Press Enter to use default: '{default_name}'): ").strip()
        
        zip_filename = user_name if user_name else default_name
        if not zip_filename.lower().endswith(".zip"):
            zip_filename += ".zip"

        create_zip(OUTPUT_DIR, zip_filename, exported_stp_files)

if __name__ == "__main__":
    main()