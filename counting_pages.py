"""
main 함수.
"""

import os

from module.process_folder import process_folder


def main():
    """main 함수. 프로그램을 종료할 때까지 반복합니다"""
    print("\n>>>>>>문서 페이지수 추출기<<<<<<\n")
    input_path = input(
        "PDF, HWP 페이지 수를 가져올 폴더 경로를 입력하세요(종료는 0을 입력) : ").strip()

    if input_path == '0':
        return 0

    output_path = input(
        "엑셀파일 경로를 입력하세요(확장자포함. 파일이 존재하지 않을 경우 새로 생성) : ").strip()

    if not os.path.isdir(input_path):
        print("입력 폴더의 경로를 다시 확인하세요")
        return main()

    process_folder(input_path, output_path)
    print(f"{output_path}에 개인정보목록이 생성되었습니다.")
    print("\n~~~모든 작업이 완료되었습니다~~~")

    return main()


if __name__ == "__main__":
    main()
