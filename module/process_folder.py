from module.count_page import get_hwp_page_count, get_pdf_page_count
from module.save_excel import save_pages_to_excel


import os
from natsort import natsorted


def process_folder(folder_path, excel_file):
    page_counts = []

    for root, _, files in os.walk(folder_path):
        for filename in natsorted(files):
            file_path = os.path.join(root, filename)

            if filename.lower().endswith('.pdf'):
                num_pages = get_pdf_page_count(file_path)
            elif filename.lower().endswith('.hwp') or filename.lower().endswith('.hwpx'):
                num_pages = get_hwp_page_count(file_path)
            else:
                continue
            page_counts.append({
                '파일명': filename,
                '경로명': file_path,
                '페이지 수': num_pages
            })

    if page_counts:
        save_pages_to_excel(page_counts, excel_file)
        print("PDF, HWP 페이지 수를 엑셀에 저장했습니다.")
    else:
        print("폴더에 PDF나 HWP 파일이 없거나 처리할 수 있는 파일이 없습니다.")
