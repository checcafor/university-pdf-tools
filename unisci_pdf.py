import os
import re
from pypdf import PdfWriter

def estrai_numero_lezione(nome_file):
    """
    Helper function to extract the lesson number from the filename.
    This solves the '1, 10, 2' sorting problem. We want '1, 2, 10'.
    """
    # We use Regex to look for the pattern "LEZIONE_" followed by digits (\d+)
    # re.IGNORECASE makes it work for both "LEZIONE" and "lezione"
    match = re.search(r'LEZIONE_(\d+)', nome_file, re.IGNORECASE)
    
    if match:
        # If we found a number, convert it to an integer (so 2 comes before 10)
        return int(match.group(1)) 
    else:
        # If the file doesn't match the pattern (e.g., "Syllabus.pdf"),
        # we assign it 'infinity' so it goes to the very end of the list.
        return float('inf') 

def unisci_pdf():
    print("--- PDF Merger (Batch) ---")
    
    # 1. Get Folder Path
    # We ask the user to drag and drop the folder into the terminal
    folder_path = input("Drag and drop the folder containing PDFs here: ").strip()

    # --- POWERSHELL FIX ---
    # When dragging folders into Windows PowerShell, it adds specific artifacts
    # like '& ' at the start or single quotes. We need to clean this up.
    if folder_path.startswith("& "):
        folder_path = folder_path[2:].strip()
    if folder_path.startswith("'") and folder_path.endswith("'"):
        folder_path = folder_path[1:-1]
    if folder_path.startswith('"') and folder_path.endswith('"'):
        folder_path = folder_path[1:-1]

    # Verify that the cleaned path is actually a valid directory
    if not os.path.isdir(folder_path):
        print(f"\n[ERROR] Invalid folder path.")
        return

    # 2. Find PDF files
    # We scan the directory for anything ending in .pdf
    files_pdf = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    # SAFETY CHECK:
    # If we ran this script before, the output file might already exist.
    # We remove it from the list so we don't try to merge the output into itself!
    nome_output = "RIEPILOGO_LEZIONI_UNITO.pdf"
    if nome_output in files_pdf:
        files_pdf.remove(nome_output)

    if not files_pdf:
        print("[WARNING] No PDF files found in this folder.")
        return

    print(f"\nFound {len(files_pdf)} files. Sorting them intelligently...")

    # 3. Natural Sorting
    # This is the crucial part: we sort using our helper function defined above.
    # It ensures the files are ordered by the actual lesson number.
    files_pdf.sort(key=estrai_numero_lezione)

    print("Detected order:")
    for f in files_pdf:
        print(f" -> {f}")

    # 4. Merging Process
    print(f"\nStarting merge of {len(files_pdf)} files...")
    merger = PdfWriter()

    try:
        # Loop through our sorted list and append each file to the merger
        for filename in files_pdf:
            filepath = os.path.join(folder_path, filename)
            merger.append(filepath)
        
        # 5. Saving the result
        # We save the final big PDF back into the same folder
        output_path = os.path.join(folder_path, nome_output)
        merger.write(output_path)
        merger.close()

        print("\n" + "="*30)
        print("[SUCCESS] Merged PDF created successfully!")
        print(f"Saved as: {output_path}")
        print("="*30)

    except Exception as e:
        print(f"\n[ERROR] Something went wrong during the merge: {e}")

if __name__ == "__main__":
    unisci_pdf()
    # Keep the window open so the user can read the result
    input("\nPress ENTER to close...")
