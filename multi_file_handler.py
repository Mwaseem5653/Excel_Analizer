# multi_file_handler.py
import os
import fitz

async def handle_files(message):
    """
    Return a list of dicts with file names and paths.
    For PDFs, each page gets a virtual path: "file.pdf - Page 1"
    """
    results = []

    for file in message.elements:
        file_path = file.path
        if not file_path:
            continue

        ext = os.path.splitext(file_path)[1].lower()

        if ext in [".jpg", ".jpeg", ".png"]:
            results.append({"file_name": os.path.basename(file_path), "path": file_path})

        elif ext == ".pdf":
            try:
                doc = fitz.open(file_path)
                for page_num in range(len(doc)):
                    results.append({"file_name": f"{os.path.basename(file_path)} - Page {page_num+1}", "path": file_path})
            except:
                continue

    return results
