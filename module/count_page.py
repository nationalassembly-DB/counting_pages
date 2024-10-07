"""
pdf와 hwp 페이지 수를 검색합니다
"""

import fitz
import win32com.client as win32


def get_pdf_page_count(pdf_file_path):
    """pdf의 페이지 수를 가져옵니다"""
    pdf = fitz.open(pdf_file_path)

    pdf_pages = pdf.page_count
    pdf.close()

    return pdf_pages


def get_hwp_page_count(hwp_file_path):
    """hwp의 페이지 수를 가져옵니다"""
    hwp = None
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.Open(hwp_file_path)
        num_pages = hwp.PageCount
    except Exception as e:  # pylint: disable=W0703
        print(f"Error: {e}")
        return None
    finally:
        if hwp:
            hwp.Quit()

    return num_pages
