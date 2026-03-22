import queue
import subprocess
import threading
from pathlib import Path
from tkinter import BooleanVar, IntVar, StringVar, Tk, messagebox, ttk, filedialog
import tkinter as tk

from hwp_merge_core import (
    OUTPUT_NAME_DEFAULT,
    SAVE_INTERVAL_DEFAULT,
    get_hwp_files,
    merge_hwp_files,
    normalize_output_path,
)


class HwpMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HWP 파일 병합기")
        self.root.geometry("780x620")
        self.root.minsize(720, 560)

        self.input_dir_var = StringVar()
        self.output_path_var = StringVar()
        self.file_count_var = StringVar(value="대상 파일: 폴더를 선택해 주세요")
        self.status_var = StringVar(value="대기 중")
        self.show_hwp_var = BooleanVar(value=False)
        self.insert_page_break_var = BooleanVar(value=False)
        self.save_interval_var = IntVar(value=SAVE_INTERVAL_DEFAULT)

        self.last_auto_output = None
        self.running = False
        self.worker = None
        self.message_queue = queue.Queue()
        self.last_output_path = None

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.build_ui()
        self.root.after(100, self.process_queue)

    def build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        file_frame = ttk.LabelFrame(self.root, text="파일 설정", padding=14)
        file_frame.grid(row=0, column=0, padx=16, pady=(16, 10), sticky="nsew")
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="입력 폴더").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.input_entry = ttk.Entry(file_frame, textvariable=self.input_dir_var)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=(10, 8), pady=(0, 8))
        self.input_button = ttk.Button(file_frame, text="폴더 선택", command=self.select_input_dir)
        self.input_button.grid(row=0, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(file_frame, text="출력 파일").grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.output_entry = ttk.Entry(file_frame, textvariable=self.output_path_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=(10, 8), pady=(0, 8))
        self.output_button = ttk.Button(file_frame, text="파일 선택", command=self.select_output_file)
        self.output_button.grid(row=1, column=2, sticky="ew", pady=(0, 8))

        self.file_count_label = ttk.Label(file_frame, textvariable=self.file_count_var)
        self.file_count_label.grid(row=2, column=0, columnspan=3, sticky="w")

        option_frame = ttk.LabelFrame(self.root, text="옵션", padding=14)
        option_frame.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="ew")
        option_frame.columnconfigure(3, weight=1)

        ttk.Label(option_frame, text="중간 저장 간격").grid(row=0, column=0, sticky="w")
        self.save_interval_spinbox = ttk.Spinbox(
            option_frame,
            from_=0,
            to=999,
            textvariable=self.save_interval_var,
            width=8,
        )
        self.save_interval_spinbox.grid(row=0, column=1, sticky="w", padx=(10, 20))
        ttk.Label(option_frame, text="0이면 중간 저장 안 함").grid(row=0, column=2, sticky="w")

        self.show_hwp_check = ttk.Checkbutton(option_frame, text="병합 중 한글 창 표시", variable=self.show_hwp_var)
        self.show_hwp_check.grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 0))

        self.page_break_check = ttk.Checkbutton(
            option_frame,
            text="파일 사이를 새 페이지로 구분",
            variable=self.insert_page_break_var,
        )
        self.page_break_check.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        action_frame = ttk.Frame(option_frame)
        action_frame.grid(row=0, column=3, rowspan=3, sticky="e")
        self.start_button = ttk.Button(action_frame, text="병합 시작", command=self.start_merge)
        self.start_button.grid(row=0, column=0, padx=(10, 0))
        self.open_output_button = ttk.Button(action_frame, text="결과 폴더 열기", command=self.open_output_folder, state="disabled")
        self.open_output_button.grid(row=0, column=1, padx=(8, 0))

        log_frame = ttk.LabelFrame(self.root, text="진행 상태", padding=14)
        log_frame.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(2, weight=1)

        self.status_label = ttk.Label(log_frame, textvariable=self.status_var)
        self.status_label.grid(row=0, column=0, sticky="w")

        self.progressbar = ttk.Progressbar(log_frame, mode="determinate")
        self.progressbar.grid(row=1, column=0, sticky="ew", pady=(10, 12))

        log_inner = ttk.Frame(log_frame)
        log_inner.grid(row=2, column=0, sticky="nsew")
        log_inner.columnconfigure(0, weight=1)
        log_inner.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_inner, wrap="word", height=18, state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_inner, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.input_dir_var.trace_add("write", self.on_input_dir_changed)
        self.output_path_var.trace_add("write", self.on_output_path_changed)

    def select_input_dir(self):
        selected = filedialog.askdirectory(title="병합할 한글 파일 폴더를 선택하세요")
        if not selected:
            return

        self.input_dir_var.set(selected)

    def select_output_file(self):
        initial_dir = self.input_dir_var.get().strip() or str(Path.cwd())
        initial_file = OUTPUT_NAME_DEFAULT

        current_value = self.output_path_var.get().strip()
        if current_value:
            current_path = Path(current_value)
            initial_dir = str(current_path.parent)
            initial_file = current_path.name

        selected = filedialog.asksaveasfilename(
            title="결과 파일 저장 위치 선택",
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".hwp",
            filetypes=[("한글 파일", "*.hwp")],
        )
        if not selected:
            return

        self.last_auto_output = None
        self.output_path_var.set(selected)

    def on_input_dir_changed(self, *_args):
        input_dir = self.input_dir_var.get().strip()
        if not input_dir:
            self.file_count_var.set("대상 파일: 폴더를 선택해 주세요")
            return

        current_output = self.output_path_var.get().strip()
        if not current_output or current_output == self.last_auto_output:
            default_output = Path(input_dir).expanduser() / OUTPUT_NAME_DEFAULT
            self.last_auto_output = str(default_output)
            self.output_path_var.set(self.last_auto_output)
            return

        self.refresh_file_count()

    def on_output_path_changed(self, *_args):
        if not self.output_path_var.get().strip():
            self.last_auto_output = None

        self.refresh_file_count()

    def refresh_file_count(self):
        input_dir = self.input_dir_var.get().strip()
        if not input_dir:
            self.file_count_var.set("대상 파일: 폴더를 선택해 주세요")
            return

        output_candidate = self.output_path_var.get().strip() or None
        try:
            files = get_hwp_files(input_dir, output_candidate)
            self.file_count_var.set(f"대상 파일: {len(files)}개 .hwp 파일")
        except Exception as exc:
            self.file_count_var.set(f"대상 파일: 확인 실패 ({exc})")

    def append_log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def set_controls_state(self, enabled):
        state = "normal" if enabled else "disabled"
        for widget in [
            self.input_entry,
            self.output_entry,
            self.input_button,
            self.output_button,
            self.save_interval_spinbox,
            self.show_hwp_check,
            self.page_break_check,
            self.start_button,
        ]:
            widget.configure(state=state)

        if enabled and self.last_output_path:
            self.open_output_button.configure(state="normal")
        else:
            self.open_output_button.configure(state="disabled")

    def validate_paths(self):
        input_dir = self.input_dir_var.get().strip()
        output_path = self.output_path_var.get().strip()

        if not input_dir:
            raise RuntimeError("입력 폴더를 선택해 주세요.")
        if not output_path:
            raise RuntimeError("출력 파일 위치를 지정해 주세요.")

        normalized_output = normalize_output_path(output_path)
        existing_inputs = get_hwp_files(input_dir)
        conflicting_inputs = [path for path in existing_inputs if path.resolve() == normalized_output]

        if conflicting_inputs:
            proceed = messagebox.askyesno(
                "덮어쓰기 확인",
                "선택한 출력 파일이 입력 폴더 안의 기존 HWP 파일과 같습니다.\n"
                "이 파일은 입력 목록에서 제외되고 덮어써집니다.\n\n계속하시겠습니까?",
            )
            if not proceed:
                raise RuntimeError("다른 출력 파일을 선택해 주세요.")

        return Path(input_dir).expanduser().resolve(), normalized_output

    def start_merge(self):
        if self.running:
            return

        try:
            input_dir, output_path = self.validate_paths()
            save_interval = int(self.save_interval_var.get())
            if save_interval < 0:
                raise RuntimeError("중간 저장 간격은 0 이상이어야 합니다.")
        except Exception as exc:
            messagebox.showerror("실행할 수 없습니다", str(exc))
            return

        self.running = True
        self.last_output_path = None
        self.set_controls_state(False)
        self.progressbar["value"] = 0
        self.progressbar["maximum"] = 1
        self.status_var.set("병합 준비 중")
        self.append_log("")
        self.append_log("=" * 70)
        self.append_log("새 병합 작업을 시작합니다.")

        worker = threading.Thread(
            target=self.run_merge_worker,
            args=(
                input_dir,
                output_path,
                save_interval,
                self.show_hwp_var.get(),
                self.insert_page_break_var.get(),
            ),
        )
        self.worker = worker
        worker.start()

    def run_merge_worker(self, input_dir, output_path, save_interval, visible, insert_page_break):
        try:
            result = merge_hwp_files(
                input_dir=input_dir,
                output_path=output_path,
                save_interval=save_interval,
                visible=visible,
                insert_page_break=insert_page_break,
                logger=lambda message: self.message_queue.put(("log", message)),
                progress_callback=lambda payload: self.message_queue.put(("progress", payload)),
            )
            self.message_queue.put(("done", result))
        except Exception as exc:
            self.message_queue.put(("error", str(exc)))

    def process_queue(self):
        try:
            while True:
                kind, payload = self.message_queue.get_nowait()
                if kind == "log":
                    self.append_log(payload)
                elif kind == "progress":
                    self.handle_progress(payload)
                elif kind == "done":
                    self.handle_done(payload)
                elif kind == "error":
                    self.handle_error(payload)
        except queue.Empty:
            pass

        self.root.after(100, self.process_queue)

    def handle_progress(self, payload):
        event_type = payload.get("type")

        if event_type == "scan_complete":
            total = max(payload.get("total", 1), 1)
            self.progressbar["maximum"] = total
            self.progressbar["value"] = 0
            self.status_var.set(f"파일 확인 완료: 총 {payload.get('total', 0)}개")
            return

        if event_type == "file_start":
            index = payload.get("index", 0)
            total = max(payload.get("total", 1), 1)
            self.progressbar["maximum"] = total
            self.progressbar["value"] = max(index - 1, 0)
            self.status_var.set(f"[{index}/{total}] 처리 중: {payload.get('file_name', '')}")
            return

        if event_type == "file_done":
            index = payload.get("index", 0)
            total = max(payload.get("total", 1), 1)
            self.progressbar["maximum"] = total
            self.progressbar["value"] = index
            if payload.get("success"):
                self.status_var.set(f"[{index}/{total}] 완료: {payload.get('file_name', '')}")
            else:
                self.status_var.set(f"[{index}/{total}] 실패 기록: {payload.get('file_name', '')}")
            return

        if event_type == "intermediate_save":
            self.status_var.set(f"중간 저장 중... ({payload.get('index', 0)}개 완료)")
            return

        if event_type == "final_save":
            self.status_var.set("최종 저장 중...")
            return

        if event_type == "warning":
            self.status_var.set(payload.get("message", "경고가 발생했습니다."))
            return

        if event_type == "completed":
            self.status_var.set(
                f"완료: 성공 {payload.get('success_count', 0)}개, 실패 {payload.get('fail_count', 0)}개"
            )

    def handle_done(self, result):
        self.running = False
        self.worker = None
        self.last_output_path = Path(result["output_path"])
        self.set_controls_state(True)
        self.status_var.set(
            f"병합 완료: {result['success_count']}/{result['total_files']}개 성공, 결과 {self.last_output_path.name}"
        )
        done_message = f"결과 파일이 생성되었습니다.\n\n{self.last_output_path}"
        if not result.get("security_module_registered", True):
            done_message += "\n\n보안 모듈 등록에 실패해 병합 중 한글 창 표시 모드로 전환되었습니다."
        messagebox.showinfo("병합 완료", done_message)

    def handle_error(self, message):
        self.running = False
        self.worker = None
        self.set_controls_state(True)
        self.status_var.set("오류로 중단됨")
        self.append_log(f"[오류] {message}")
        messagebox.showerror("병합 실패", message)

    def open_output_folder(self):
        if not self.last_output_path:
            return

        try:
            subprocess.Popen(["explorer", str(self.last_output_path.parent)])
        except Exception as exc:
            messagebox.showerror("폴더 열기 실패", str(exc))

    def on_close(self):
        if self.running:
            messagebox.showwarning(
                "병합 진행 중",
                "병합이 진행 중일 때는 창을 닫을 수 없습니다.\n완료 후 닫아 주세요.",
            )
            return

        self.root.destroy()


def main():
    root = Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    HwpMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
