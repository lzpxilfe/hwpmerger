import argparse
import sys
from pathlib import Path
from tkinter import Tk, filedialog

from hwp_merge_core import OUTPUT_NAME_DEFAULT, SAVE_INTERVAL_DEFAULT, merge_hwp_files, normalize_output_path


def parse_args():
    parser = argparse.ArgumentParser(description="폴더 안의 .hwp 파일을 순서대로 병합합니다.")
    parser.add_argument(
        "input_dir",
        nargs="?",
        help="병합할 .hwp 파일이 들어 있는 폴더 경로. 비우면 폴더 선택 창이 열립니다.",
    )
    parser.add_argument(
        "-o",
        "--output-file",
        default=OUTPUT_NAME_DEFAULT,
        help=f"생성할 결과 파일 경로 또는 파일 이름 (기본값: {OUTPUT_NAME_DEFAULT})",
    )
    parser.add_argument(
        "--save-interval",
        type=int,
        default=SAVE_INTERVAL_DEFAULT,
        help=f"몇 개마다 중간 저장할지 설정 (기본값: {SAVE_INTERVAL_DEFAULT}, 0이면 사용 안 함)",
    )
    parser.add_argument(
        "--show-hwp",
        action="store_true",
        help="병합 중 한글 창을 화면에 표시합니다.",
    )
    parser.add_argument(
        "--page-break",
        action="store_true",
        help="파일 사이에 새 페이지 구분을 넣습니다. 기본값은 끔입니다.",
    )
    return parser.parse_args()


def choose_input_dir():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askdirectory(title="병합할 한글 파일 폴더를 선택하세요")
    root.destroy()
    return Path(selected) if selected else None


def main():
    args = parse_args()
    input_dir = Path(args.input_dir).expanduser() if args.input_dir else choose_input_dir()

    if not input_dir:
        print("폴더 선택이 취소되었습니다.")
        return 1

    try:
        output_file = Path(args.output_file)
        if output_file.is_absolute():
            output_path = normalize_output_path(output_file)
        else:
            output_path = normalize_output_path(Path(input_dir) / output_file)

        merge_hwp_files(
            input_dir=input_dir,
            output_path=output_path,
            save_interval=args.save_interval,
            visible=args.show_hwp,
            insert_page_break=args.page_break,
        )
        print("완료!")
        return 0
    except Exception as exc:
        print(f"[오류] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
