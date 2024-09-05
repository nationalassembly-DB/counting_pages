"""
pdf와 hwp 페이지 수를 검색합니다
"""

import PyPDF2
import PyPDF2.errors
import win32com.client as win32


def get_pdf_page_count(pdf_file_path):
    """pdf의 페이지 수를 가져옵니다"""
    try:
        with open(pdf_file_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            num_pages = len(pdf_reader.pages)
            return num_pages
    except PyPDF2.errors.PdfReadError as e:
        print(f"Pdf Read Error: {e}")
        return None


def get_hwp_page_count(hwp_file_path):
    """hwp의 페이지 수를 가져옵니다"""
    hwp = None
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.Open(hwp_file_path)
        num_pages = hwp.PageCount
    except Exception as e:
        print(f"Error: {e}")
        return None
    finally:
        if hwp:
            try:
                hwp.ReleaseControl()
                hwp.Quit()
            except:
                pass

    return num_pages
