"""Create README screenshots from the current Tkinter interface.

Run this after a meaningful UI change:
    py tools\\generate_readme_assets.py
"""

from pathlib import Path
import sys
import time

from PIL import ImageGrab
import tkinter as tk


ROOT_DIR = Path(__file__).resolve().parent.parent
ASSETS_DIR = ROOT_DIR / "assets"
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from hwp_merger_gui import HwpMergerApp


def set_log(text_widget, lines):
    text_widget.configure(state="normal")
    text_widget.delete("1.0", "end")
    text_widget.insert("end", "\n".join(lines))
    text_widget.configure(state="disabled")


def capture_window(root, output_path):
    root.update_idletasks()
    root.update()
    time.sleep(0.25)
    x = root.winfo_rootx()
    y = root.winfo_rooty()
    width = root.winfo_width()
    height = root.winfo_height()
    ImageGrab.grab(bbox=(x, y, x + width, y + height)).save(output_path)


def fill_merge_example(app):
    tab = app.merge_tab
    app.notebook.select(tab.frame)
    tab.input_dir_var.set(r"C:\Survey\source_forms")
    tab.output_path_var.set(r"C:\Survey\output\integrated.hwp")
    tab.file_count_var.set("대상 파일: 24개 .hwp 파일")
    tab.save_interval_var.set(10)
    tab.insert_page_break_var.set(True)
    tab.status_var.set("병합 진행 중: 9 / 24")
    tab.progress_bar["maximum"] = 24
    tab.progress_bar["value"] = 9
    set_log(tab.log_text, [
        "새 병합 작업을 시작합니다.",
        "입력 폴더: C:\\Survey\\source_forms",
        "출력 파일: C:\\Survey\\output\\integrated.hwp",
        "총 24개 파일 발견",
        "[1/24] 첫 파일 열기: 01_조사표.hwp",
        "[2/24] 02_조사표.hwp  완료",
        "[9/24] 09_조사표.hwp  병합 중...",
    ])


def fill_split_example(app):
    tab = app.split_tab
    app.notebook.select(tab.frame)
    tab.input_file_var.set(r"C:\Survey\대전 지표조사.hwp")
    tab.output_dir_var.set(r"C:\Survey\분리")
    tab.pattern_var.set("{name}")
    tab.status_var.set("분리 미리보기: 136개 표 묶음")
    tab.progress_bar["maximum"] = 136
    tab.progress_bar["value"] = 114
    for item in tab.tree.get_children():
        tab.tree.delete(item)
    rows = [
        (1, "표 단위", "대전_026 대전 효평동 유물산포지2.hwp"),
        (2, "표 단위", "대전_027 대전 효평동 유물산포지1.hwp"),
        (75, "표 단위", "대전_057 대전 소제동 유적추정지.hwp"),
        (79, "표 단위", "대전_057 대전 소제동 유적추정지 (2).hwp"),
        (136, "표 단위", "대전_068 대전 동구 천동2 주거환경개선 사업부지 내 유적.hwp"),
    ]
    for row in rows:
        tab.tree.insert("", "end", values=row)
    set_log(tab.log_text, [
        "rhwp 표 구조 분석 중: 대전 지표조사.hwp",
        "rhwp 표 구조 확인: 실제 표 207개, 유적 시작 표 136개 검출 완료",
        "A3 표 묶음 저장 완료: 대전_067 대전 문화동 유물산포지1.hwp (표 2개)",
        "기존 검산 통과 파일 유지: 대전_057 대전 소제동 유적추정지.hwp",
        "동명 유적 묶음 재생성: 대전_057 대전 소제동 유적추정지 (2).hwp",
    ])


def main():
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    root = tk.Tk()
    root.geometry("880x740+120+80")
    root.attributes("-topmost", True)
    app = HwpMergerApp(root)

    fill_merge_example(app)
    capture_window(root, ASSETS_DIR / "readme-merge.png")

    fill_split_example(app)
    capture_window(root, ASSETS_DIR / "readme-split.png")

    root.destroy()
    print(ASSETS_DIR / "readme-merge.png")
    print(ASSETS_DIR / "readme-split.png")


if __name__ == "__main__":
    main()
