import queue
import subprocess
import threading
from pathlib import Path
from tkinter import BooleanVar, IntVar, StringVar, Tk, messagebox, ttk, filedialog
import tkinter as tk
import pythoncom

from hwp_merge_core import (
    OUTPUT_NAME_DEFAULT,
    SAVE_INTERVAL_DEFAULT,
    get_hwp_files,
    merge_hwp_files,
    normalize_output_path,
)
from hwp_split_core import (
    SPLIT_MODE_AUTO,
    SPLIT_MODE_N_PAGES,
    SPLIT_MODE_N_FILES,
    SPLIT_MODE_TABLE_NAME,
    parse_items_from_rhwp,
    build_table_output_plan,
    build_page_ranges_from_starts,
    build_page_ranges_n_pages,
    build_page_ranges_n_files,
    detect_page_breaks_from_open_doc,
    execute_split_by_plan,
)


# ---------------------------------------------------------------------------
# 병합 탭
# ---------------------------------------------------------------------------

class MergeTab:
    def __init__(self, parent_notebook):
        self.frame = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame, text="📄 병합")

        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(2, weight=1)

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

        self._build_ui()

    def _build_ui(self):
        file_frame = ttk.LabelFrame(self.frame, text="파일 설정", padding=14)
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

        option_frame = ttk.LabelFrame(self.frame, text="옵션", padding=14)
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

        self.merge_button = ttk.Button(action_frame, text="병합 시작", command=self.start_merge)
        self.merge_button.pack(side="top", fill="x", pady=(0, 6))

        self.open_button = ttk.Button(action_frame, text="결과 파일 열기", command=self.open_result)
        self.open_button.pack(side="top", fill="x")

        log_frame = ttk.LabelFrame(self.frame, text="진행 상황", padding=14)
        log_frame.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(2, weight=1)

        status_row = ttk.Frame(log_frame)
        status_row.grid(row=0, column=0, sticky="ew")
        status_row.columnconfigure(1, weight=1)

        ttk.Label(status_row, text="상태:").grid(row=0, column=0, sticky="w")
        ttk.Label(status_row, textvariable=self.status_var).grid(row=0, column=1, sticky="w", padx=(6, 0))

        self.progress_bar = ttk.Progressbar(log_frame, mode="determinate")
        self.progress_bar.grid(row=1, column=0, sticky="ew", pady=(8, 8))

        self.log_text = tk.Text(log_frame, height=10, state="disabled", wrap="word")
        self.log_text.grid(row=2, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=2, column=1, sticky="ns")
        self.log_text["yscrollcommand"] = scrollbar.set

    def select_input_dir(self):
        d = filedialog.askdirectory(title="병합할 HWP 파일이 있는 폴더 선택")
        if not d:
            return
        self.input_dir_var.set(d)
        output_path = self.output_path_var.get().strip()
        if not output_path or output_path == str(self.last_auto_output):
            auto = Path(d) / OUTPUT_NAME_DEFAULT
            self.output_path_var.set(str(auto))
            self.last_auto_output = auto
        self._refresh_file_count()

    def select_output_file(self):
        f = filedialog.asksaveasfilename(
            title="저장할 파일 위치 선택",
            defaultextension=".hwp",
            filetypes=[("HWP 파일", "*.hwp")],
        )
        if f:
            self.output_path_var.set(f)

    def _refresh_file_count(self):
        input_dir = self.input_dir_var.get().strip()
        if not input_dir:
            self.file_count_var.set("대상 파일: 폴더를 선택해 주세요")
            return
        try:
            files = get_hwp_files(input_dir, self.output_path_var.get().strip() or None)
            self.file_count_var.set(f"대상 파일: {len(files)}개")
        except Exception as exc:
            self.file_count_var.set(f"오류: {exc}")

    def _set_ui_running(self, running):
        state = "disabled" if running else "normal"
        self.merge_button.config(state=state)
        self.open_button.config(state="disabled" if running else "normal")
        self.input_button.config(state=state)
        self.output_button.config(state=state)
        self.save_interval_spinbox.config(state=state)
        self.show_hwp_check.config(state=state)
        self.page_break_check.config(state=state)
        self.running = running

    def _append_log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def start_merge(self):
        input_dir = self.input_dir_var.get().strip()
        output_path = self.output_path_var.get().strip()

        if not input_dir:
            messagebox.showwarning("경고", "입력 폴더를 선택해 주세요.")
            return
        if not output_path:
            messagebox.showwarning("경고", "출력 파일 경로를 입력해 주세요.")
            return

        self._clear_log()
        self._set_ui_running(True)
        self.status_var.set("병합 준비 중…")
        self.progress_bar["value"] = 0

        params = {
            "input_dir": input_dir,
            "output_path": output_path,
            "visible": self.show_hwp_var.get(),
            "insert_page_break": self.insert_page_break_var.get(),
            "save_interval": self.save_interval_var.get(),
        }
        self.worker = threading.Thread(target=self._merge_worker, kwargs=params, daemon=True)
        self.worker.start()
        self.frame.after(100, self._process_queue)

    def _merge_worker(self, input_dir, output_path, visible, insert_page_break, save_interval):
        pythoncom.CoInitialize()
        q = self.message_queue

        def logger(msg):
            q.put({"type": "log", "message": msg})

        def progress_callback(payload):
            q.put(payload)

        try:
            merge_hwp_files(
                input_dir=input_dir,
                output_path=output_path,
                visible=visible,
                insert_page_break=insert_page_break,
                save_interval=save_interval,
                logger=logger,
                progress_callback=progress_callback,
            )
        except Exception as exc:
            q.put({"type": "error", "message": str(exc)})
        finally:
            pythoncom.CoUninitialize()
            q.put({"type": "done"})

    def _process_queue(self):
        try:
            while True:
                msg = self.message_queue.get_nowait()
                mtype = msg.get("type")

                if mtype == "log":
                    self._append_log(msg["message"])

                elif mtype == "progress":
                    total = msg.get("total", 0)
                    current = msg.get("current", 0)
                    if total:
                        self.progress_bar["maximum"] = total
                        self.progress_bar["value"] = current
                    status = msg.get("status", "")
                    if status:
                        self.status_var.set(status)

                elif mtype == "warning":
                    self._append_log(f"[경고] {msg['message']}")

                elif mtype == "error":
                    self._append_log(f"[오류] {msg['message']}")
                    messagebox.showerror("오류", msg["message"])
                    self._set_ui_running(False)
                    self.status_var.set("오류 발생")

                elif mtype == "done":
                    self._set_ui_running(False)
                    self.status_var.set("완료")
                    output_path = self.output_path_var.get().strip()
                    self.last_output_path = output_path
                    if output_path and Path(output_path).exists():
                        if messagebox.askyesno("완료", "병합이 완료되었습니다.\n결과 파일을 여시겠습니까?"):
                            self.open_result()
        except Exception:
            pass
        finally:
            if self.running:
                self.frame.after(100, self._process_queue)

    def open_result(self):
        path = self.last_output_path or self.output_path_var.get().strip()
        if not path or not Path(path).exists():
            messagebox.showwarning("경고", "열 파일이 없습니다.")
            return
        subprocess.Popen(["start", "", path], shell=True)

    def on_close_check(self):
        if self.running:
            if not messagebox.askyesno("종료 확인", "병합이 진행 중입니다. 종료하시겠습니까?"):
                return False
        return True


# ---------------------------------------------------------------------------
# 분리 탭
# ---------------------------------------------------------------------------

class SplitTab:
    def __init__(self, parent_notebook):
        self.frame = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame, text="✂️ 분리")

        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(3, weight=1)

        self.input_file_var = StringVar()
        self.output_dir_var = StringVar()
        self.status_var = StringVar(value="대기 중")
        self.show_hwp_var = BooleanVar(value=False)

        self.split_mode_var = StringVar(value=SPLIT_MODE_TABLE_NAME)
        self.n_var = IntVar(value=1)
        self.pattern_var = StringVar(value="{name}")

        self._preview_ranges = []
        self._preview_names = []

        self.running = False
        self.worker = None
        self.message_queue = queue.Queue()

        self._build_ui()

    def _build_ui(self):
        file_frame = ttk.LabelFrame(self.frame, text="파일 설정", padding=12)
        file_frame.grid(row=0, column=0, padx=14, pady=(12, 8), sticky="nsew")
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="입력 파일").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.input_entry = ttk.Entry(file_frame, textvariable=self.input_file_var)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=(10, 8), pady=(0, 6))
        self.input_button = ttk.Button(file_frame, text="파일 선택", command=self.select_input_file)
        self.input_button.grid(row=0, column=2, sticky="ew", pady=(0, 6))

        ttk.Label(file_frame, text="출력 폴더").grid(row=1, column=0, sticky="w")
        self.output_entry = ttk.Entry(file_frame, textvariable=self.output_dir_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=(10, 8))
        self.output_dir_button = ttk.Button(file_frame, text="폴더 선택", command=self.select_output_dir)
        self.output_dir_button.grid(row=1, column=2, sticky="ew")

        option_frame = ttk.LabelFrame(self.frame, text="분리 모드", padding=12)
        option_frame.grid(row=1, column=0, padx=14, pady=(0, 8), sticky="ew")
        option_frame.columnconfigure(3, weight=1)

        modes = [
            (SPLIT_MODE_TABLE_NAME, "📋 [도면 명칭 + 유적명] 표 기준 자동 분리 (권장)"),
            (SPLIT_MODE_AUTO,       "📄 페이지 나누기 마커 기준 분리"),
            (SPLIT_MODE_N_PAGES,    "🔢 N페이지마다 균등 분리"),
            (SPLIT_MODE_N_FILES,    "➗ 총 N개 파일로 균등 분리"),
        ]
        for idx, (mode_val, label) in enumerate(modes):
            rb = ttk.Radiobutton(
                option_frame,
                text=label,
                variable=self.split_mode_var,
                value=mode_val,
                command=self._on_mode_change,
            )
            rb.grid(row=idx // 2, column=(idx % 2) * 2, columnspan=2, sticky="w", padx=(0, 10), pady=3)

        self.detail_frame = ttk.Frame(option_frame, padding=(0, 6, 0, 0))
        self.detail_frame.grid(row=2, column=0, columnspan=4, sticky="ew")
        self.detail_frame.columnconfigure(1, weight=1)

        self.n_label = ttk.Label(self.detail_frame, text="페이지/파일 수(N):")
        self.n_label.grid(row=0, column=0, sticky="w", pady=2)
        self.n_spinbox = ttk.Spinbox(self.detail_frame, from_=1, to=9999, textvariable=self.n_var, width=8)
        self.n_spinbox.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=2)

        self.pat_label = ttk.Label(self.detail_frame, text="파일명 서식 패턴:")
        self.pat_label.grid(row=1, column=0, sticky="w", pady=2)
        self.pat_entry = ttk.Entry(self.detail_frame, textvariable=self.pattern_var)
        self.pat_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=2)
        ttk.Label(self.detail_frame, text="기본: {name}", foreground="gray").grid(row=1, column=2, sticky="w", padx=(6, 0))

        btn_box = ttk.Frame(option_frame)
        btn_box.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        btn_box.columnconfigure(0, weight=1)
        btn_box.columnconfigure(1, weight=1)
        btn_box.columnconfigure(2, weight=1)

        self.preview_button = ttk.Button(btn_box, text="🔍 미리보기 / 표 분석", command=self.start_preview)
        self.preview_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.split_button = ttk.Button(btn_box, text="✂️ 분리 시작", command=self.start_split)
        self.split_button.grid(row=0, column=1, sticky="ew", padx=4)

        self.open_dir_button = ttk.Button(btn_box, text="📂 결과 폴더 열기", command=self.open_result_dir)
        self.open_dir_button.grid(row=0, column=2, sticky="ew", padx=(4, 0))

        preview_frame = ttk.LabelFrame(self.frame, text="분리 미리보기 목록", padding=10)
        preview_frame.grid(row=2, column=0, padx=14, pady=(0, 8), sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(preview_frame, columns=("num", "pages", "filename"), show="headings", height=6)
        self.tree.heading("num", text="순번")
        self.tree.heading("pages", text="페이지 범위")
        self.tree.heading("filename", text="생성될 파일명 (.hwp)")
        self.tree.column("num", width=50, anchor="center")
        self.tree.column("pages", width=120, anchor="center")
        self.tree.column("filename", width=550, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")

        tree_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=tree_scroll.set)

        log_frame = ttk.LabelFrame(self.frame, text="진행 상황 / 로그", padding=12)
        log_frame.grid(row=3, column=0, padx=14, pady=(0, 12), sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(2, weight=1)

        status_row = ttk.Frame(log_frame)
        status_row.grid(row=0, column=0, sticky="ew")
        status_row.columnconfigure(1, weight=1)
        ttk.Label(status_row, text="상태:").grid(row=0, column=0, sticky="w")
        ttk.Label(status_row, textvariable=self.status_var).grid(row=0, column=1, sticky="w", padx=(6, 0))

        self.progress_bar = ttk.Progressbar(log_frame, mode="determinate")
        self.progress_bar.grid(row=1, column=0, sticky="ew", pady=(6, 6))

        self.log_text = tk.Text(log_frame, height=5, state="disabled", wrap="word")
        self.log_text.grid(row=2, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=2, column=1, sticky="ns")
        self.log_text["yscrollcommand"] = scrollbar.set

        self._on_mode_change()

    def _on_mode_change(self):
        mode = self.split_mode_var.get()
        if mode in (SPLIT_MODE_N_PAGES, SPLIT_MODE_N_FILES):
            self.n_label.grid()
            self.n_spinbox.grid()
            self.pat_label.grid_remove()
            self.pat_entry.grid_remove()
        else:
            self.n_label.grid_remove()
            self.n_spinbox.grid_remove()
            self.pat_label.grid()
            self.pat_entry.grid()

    def select_input_file(self):
        f = filedialog.askopenfilename(
            title="분리할 HWP 파일 선택",
            filetypes=[("HWP 파일", "*.hwp")],
        )
        if not f:
            return
        self.input_file_var.set(f)
        if not self.output_dir_var.get().strip():
            self.output_dir_var.set(str(Path(f).parent / (Path(f).stem + "_분리")))
        for item in self.tree.get_children():
            self.tree.delete(item)

    def select_output_dir(self):
        d = filedialog.askdirectory(title="분리 파일을 저장할 폴더 선택")
        if d:
            self.output_dir_var.set(d)

    def _append_log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _set_ui_running(self, running):
        state = "disabled" if running else "normal"
        self.split_button.config(state=state)
        self.preview_button.config(state=state)
        self.input_button.config(state=state)
        self.output_dir_button.config(state=state)
        self.running = running

    def start_preview(self):
        input_file = self.input_file_var.get().strip()
        if not input_file:
            messagebox.showwarning("경고", "입력 파일을 먼저 선택해 주세요.")
            return

        self._clear_log()
        self._set_ui_running(True)
        self.status_var.set("문서 분석 및 미리보기 추출 중…")
        self.progress_bar["value"] = 0

        for item in self.tree.get_children():
            self.tree.delete(item)

        mode = self.split_mode_var.get()
        n = self.n_var.get()
        pattern = self.pattern_var.get().strip() or "{name}"

        t = threading.Thread(
            target=self._preview_worker,
            args=(input_file, mode, n, pattern),
            daemon=True,
        )
        t.start()
        self.frame.after(100, self._process_queue)

    def _preview_worker(self, input_file, mode, n, pattern):
        pythoncom.CoInitialize()
        q = self.message_queue
        q.put({"type": "log", "message": "[빌드] rhwp 표 구조 + 실제 A3 바탕 탭 검증 v35 (2026-08-19)"})

        def logger(msg):
            q.put({"type": "log", "message": msg})

        try:
            input_path = Path(input_file)
            if mode == SPLIT_MODE_TABLE_NAME:
                q.put({"type": "log", "message": f"rhwp 표 구조 분석 중: {input_path.name}"})
                ranges, table_names = parse_items_from_rhwp(input_path, logger=logger)
                desc = f"rhwp 표 구조 기준: 총 {len(table_names)}개 유적 항목 감지"
                q.put({"type": "preview_done", "ranges": ranges, "desc": desc, "names": table_names, "pattern": pattern})
                return

            import win32com.client, time as _t
            hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            try:
                hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
            except Exception:
                pass
            
            q.put({"type": "log", "message": f"HWP 파일 여는 중: {input_path.name}"})
            hwp.Open(str(input_path), "HWP", "forceopen:true")
            _t.sleep(0.5)
            total_pages = hwp.PageCount
            q.put({"type": "log", "message": f"총 페이지 수: {total_pages}"})

            if mode == SPLIT_MODE_AUTO:
                segment_starts, _ = detect_page_breaks_from_open_doc(hwp)
                try:
                    hwp.Quit()
                except Exception:
                    pass
                ranges = build_page_ranges_from_starts(segment_starts, total_pages)
                desc = f"페이지 나누기 자동 감지: {len(ranges)}개 세그먼트"
                q.put({"type": "preview_done", "ranges": ranges, "desc": desc, "names": None, "pattern": pattern})

            elif mode == SPLIT_MODE_N_PAGES:
                try:
                    hwp.Quit()
                except Exception:
                    pass
                ranges = build_page_ranges_n_pages(total_pages, n)
                desc = f"{n}페이지마다 분리: 총 {len(ranges)}개 파일"
                q.put({"type": "preview_done", "ranges": ranges, "desc": desc, "names": None, "pattern": pattern})

            else: # N_FILES
                try:
                    hwp.Quit()
                except Exception:
                    pass
                ranges = build_page_ranges_n_files(total_pages, n)
                desc = f"총 {n}개로 균등 분리: 총 {len(ranges)}개 파일"
                q.put({"type": "preview_done", "ranges": ranges, "desc": desc, "names": None, "pattern": pattern})

        except Exception as exc:
            q.put({"type": "error", "message": str(exc)})
        finally:
            pythoncom.CoUninitialize()
            q.put({"type": "done", "is_preview": True})

    def start_split(self):
        input_file = self.input_file_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not input_file:
            messagebox.showwarning("경고", "입력 파일을 선택해 주세요.")
            return
        if not output_dir:
            messagebox.showwarning("경고", "출력 폴더를 선택해 주세요.")
            return

        if not self._preview_ranges or not self._preview_names:
            messagebox.showinfo("안내", "먼저 [🔍 미리보기 / 표 분석]을 실행하여 분리 계획을 확인합니다.")
            self.start_preview()
            return

        pattern = self.pattern_var.get().strip() or "{name}"

        self._clear_log()
        self._set_ui_running(True)
        self.status_var.set("분리 작업 진행 중…")
        self.progress_bar["value"] = 0

        params = dict(
            input_path=input_file,
            output_dir=output_dir,
            page_ranges=self._preview_ranges,
            names=self._preview_names,
            pattern=pattern,
            visible=self.show_hwp_var.get(),
            split_by_table=(self.split_mode_var.get() == SPLIT_MODE_TABLE_NAME),
        )
        self.worker = threading.Thread(target=self._split_worker, kwargs=params, daemon=True)
        self.worker.start()
        self.frame.after(100, self._process_queue)

    def _split_worker(self, input_path, output_dir, page_ranges, names, pattern, visible, split_by_table):
        pythoncom.CoInitialize()
        q = self.message_queue

        def logger(msg):
            q.put({"type": "log", "message": msg})

        def progress_callback(payload):
            q.put(payload)

        try:
            saved = execute_split_by_plan(
                input_path=input_path,
                output_dir=output_dir,
                page_ranges=page_ranges,
                names=names,
                pattern=pattern,
                visible=visible,
                logger=logger,
                progress_callback=progress_callback,
                split_by_table=split_by_table,
            )
            q.put({"type": "split_complete", "count": len(saved), "output_dir": output_dir})
        except Exception as exc:
            q.put({"type": "error", "message": str(exc)})
        finally:
            pythoncom.CoUninitialize()
            q.put({"type": "done", "is_preview": False})

    def _process_queue(self):
        try:
            while True:
                msg = self.message_queue.get_nowait()
                mtype = msg.get("type")

                if mtype == "log":
                    self._append_log(msg["message"])

                elif mtype == "progress":
                    total = msg.get("total", 0)
                    current = msg.get("current", 0)
                    if total:
                        self.progress_bar["maximum"] = total
                        self.progress_bar["value"] = current
                    status = msg.get("status", "")
                    if status:
                        self.status_var.set(status)

                elif mtype == "warning":
                    self._append_log(f"[경고] {msg['message']}")

                elif mtype == "error":
                    self._append_log(f"[오류] {msg['message']}")
                    messagebox.showerror("오류", msg["message"])
                    self._set_ui_running(False)
                    self.status_var.set("오류 발생")

                elif mtype == "preview_done":
                    ranges = msg["ranges"]
                    desc = msg["desc"]
                    names = msg.get("names")
                    pattern = msg.get("pattern", "{name}")
                    
                    self._preview_ranges = ranges
                    self._preview_names = names
                    
                    for item in self.tree.get_children():
                        self.tree.delete(item)
                        
                    table_output_plan = build_table_output_plan(names, pattern) if names else []
                    digits = len(str(len(ranges))) if ranges else 1
                    stem = Path(self.input_file_var.get()).stem
                    
                    for idx, (s, e) in enumerate(ranges, start=1):
                        num_str = str(idx).zfill(digits)
                        if table_output_plan and idx - 1 < len(table_output_plan):
                            fname = table_output_plan[idx - 1]["filename"]
                        else:
                            fname = f"{stem}_{num_str}.hwp"
                            
                        page_label = "표 단위" if (s, e) == (0, 0) else f"p.{s} ~ p.{e}"
                        self.tree.insert("", "end", values=(idx, page_label, fname))
                        
                    self.status_var.set(f"{desc} (총 {len(ranges)}개)")

                elif mtype == "split_complete":
                    count = msg["count"]
                    out_dir = msg["output_dir"]
                    self.status_var.set(f"분리 완료 ({count}개 파일)")
                    if messagebox.askyesno("완료", f"분리가 완료되었습니다!\n총 {count}개 파일 → {out_dir}\n\n결과 폴더를 여시겠습니까?"):
                        self.open_result_dir()

                elif mtype == "done":
                    self._set_ui_running(False)
                    if self.status_var.get() not in ("오류 발생",) and not self.status_var.get().startswith("분리 완료"):
                        self.status_var.set("대기 중")

        except Exception:
            pass
        finally:
            if self.running:
                self.frame.after(100, self._process_queue)

    def open_result_dir(self):
        d = self.output_dir_var.get().strip()
        if not d:
            messagebox.showwarning("경고", "출력 폴더가 설정되지 않았습니다.")
            return
        subprocess.Popen(["explorer", d])

    def on_close_check(self):
        if self.running:
            if not messagebox.askyesno("종료 확인", "분리가 진행 중입니다. 종료하시겠습니까?"):
                return False
        return True


# ---------------------------------------------------------------------------
# 앱 루트
# ---------------------------------------------------------------------------

class HwpMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HWP 파일 병합/분리기 - rhwp 표 구조 · 실제 A3 바탕 탭 v35")
        self.root.geometry("880x740")
        self.root.minsize(760, 620)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        self.merge_tab = MergeTab(self.notebook)
        self.split_tab = SplitTab(self.notebook)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        if not self.merge_tab.on_close_check():
            return
        if not self.split_tab.on_close_check():
            return
        self.root.destroy()


def main():
    root = Tk()
    app = HwpMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
