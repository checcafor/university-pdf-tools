import os
from docx2pdf import convert

def convert_folder():
    print("--- Batch Word to PDF Converter (with Subfolder) ---")
    
    # 1. Get Folder Path
    # We ask the user to drag and drop the folder containing the Word files
    folder_path = input("Drag and drop the folder here: ").strip()

    # --- POWERSHELL FIX ---
    # When dragging folders into Windows PowerShell, it creates artifacts 
    # like '& ' at the beginning or extra quotes. We need to clean the string.
    if folder_path.startswith("& "):
        folder_path = folder_path[2:].strip()
    if folder_path.startswith("'") and folder_path.endswith("'"):
        folder_path = folder_path[1:-1]
    if folder_path.startswith('"') and folder_path.endswith('"'):
        folder_path = folder_path[1:-1]
    # ---------------------------------
    
    # 2. Validation
    # Check if the path actually exists and is a directory
    if not os.path.isdir(folder_path):
        print(f"\n[ERROR] The path provided is not a valid folder.")
        return

    # 3. Create Destination Subfolder
    # To keep the main folder clean, we save PDFs in a dedicated subfolder.
    subfolder_name = "Converted_PDFs"
    output_folder = os.path.join(folder_path, subfolder_name)

    # 'makedirs' creates the folder. 'exist_ok=True' prevents errors if it already exists.
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder, exist_ok=True)
            print(f"\n[INFO] Created output folder: {subfolder_name}")
        except Exception as e:
            print(f"\n[ERROR] Could not create output folder: {e}")
            return
    else:
        print(f"\n[INFO] Saving files to existing folder: {subfolder_name}")

    # 4. Find Word Files
    # We list all files and filter them:
    # - Must end with .doc or .docx
    # - Must NOT start with '~$' (these are temporary hidden files created by Word when a doc is open)
    files_word = [
        f for f in os.listdir(folder_path) 
        if f.lower().endswith(('.docx', '.doc')) and not f.startswith('~$')
    ]

    if not files_word:
        print("[WARNING] No Word documents found in the main folder.")
        return

    print(f"Found {len(files_word)} documents to convert.\n")

    count_ok = 0
    count_errors = 0

    for filename in files_word:
        # Full path of the original Word file
        full_path_word = os.path.join(folder_path, filename)
        
        # Build the new PDF path inside the subfolder
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        full_path_pdf = os.path.join(output_folder, pdf_name)
        
        # 'end=\r' allows us to overwrite the current line in the terminal for a cleaner look
        print(f"Converting: {filename} ...", end="\r")
        
        try:
            # The actual conversion process
            convert(full_path_word, full_path_pdf)
            print(f"[OK] Converted: {filename}      ") # Spaces are to clear previous text
            count_ok += 1
        except Exception as e:
            print(f"\n[ERROR] Failed to convert {filename}: {e}")
            count_errors += 1

    # Final Summary
    print("\n" + "="*30)
    print("OPERATION COMPLETED")
    print(f"Files saved in: {output_folder}")
    print(f"Success: {count_ok} | Errors: {count_errors}")
    print("="*30)

if __name__ == "__main__":
    convert_folder()
    input("\nPress ENTER to close...")
