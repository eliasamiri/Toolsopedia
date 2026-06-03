import os
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.dirname(CURRENT_DIR)
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

from services.converter import ConverterService, convert_range_worker, merge_workbooks

PDF_PATH = "/Users/lzb/Documents/PDFtoExcel/pdf-to-excel-app/uploads/Stats_List___Detailed_0171363.pdf"
OUTPUT_PATH = "/Users/lzb/Documents/PDFtoExcel/pdf-to-excel-app/converted/Stats_List___Detailed_0171363.xlsx"
CHUNK_SIZE = 100
SINGLE_SHEET = True


def main():
    analysis = ConverterService().analyze(PDF_PATH)
    total_pages = analysis["page_count"]

    chunks = []
    start = 1
    while start <= total_pages:
        end = min(start + CHUNK_SIZE - 1, total_pages)
        chunks.append((start, end))
        start = end + 1

    part_paths = []
    args_list = []
    for index, (start_page, end_page) in enumerate(chunks, start=1):
        part_path = f"{OUTPUT_PATH}.part_{index}.xlsx"
        part_paths.append(part_path)
        args_list.append((PDF_PATH, part_path, start_page, end_page, SINGLE_SHEET))

    max_workers = max(2, min(4, os.cpu_count() or 2))
    completed = 0
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(convert_range_worker, args): args for args in args_list}
        for future in as_completed(future_map):
            future.result()
            start_page, end_page = future_map[future][2], future_map[future][3]
            completed += (end_page - start_page + 1)
            print(f"completed {completed}/{total_pages}")

    merge_workbooks(part_paths, OUTPUT_PATH, single_sheet=SINGLE_SHEET)
    for path in part_paths:
        if os.path.exists(path):
            os.remove(path)

    print("done", OUTPUT_PATH)


if __name__ == "__main__":
    main()
