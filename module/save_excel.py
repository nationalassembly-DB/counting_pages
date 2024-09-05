"""
전달받은 리스트를 이용해 파일명 페이지 수 등을 엑셀에 저장합니다
"""

import os
import openpyxl


def save_pages_to_excel(pages, excel_file):
    """리스트를 전달받아 파일명, 페이지 수, 경로명을 엑셀에 저장합니다"""
    if os.path.exists(excel_file):
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["연번", "파일명", "페이지 수", "경로명"])
    start_no = sheet.max_row if sheet.cell(
        row=1, column=1).value == "No." else 0

    for i, page in enumerate(pages, start=start_no + 1):
        sheet.append([i, page['파일명'], page['페이지 수'], page['경로명']])

    workbook.save(excel_file)
