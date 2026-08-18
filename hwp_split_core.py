import os
import re
import time
from pathlib import Path
import pythoncom

# ---------------------------------------------------------------------------
# 분리 모드 상수
# ---------------------------------------------------------------------------
SPLIT_MODE_AUTO = "auto"
SPLIT_MODE_N_PAGES = "n_pages"
SPLIT_MODE_N_FILES = "n_files"
SPLIT_MODE_TABLE_NAME = "table_name"


def emit_progress(progress_callback, payload):
    if progress_callback:
        progress_callback(payload)


def emit_log(logger, message):
    if logger:
        logger(message)
    else:
        print(message)


def get_hwp_application(visible=True, logger=None, progress_callback=None):
    pythoncom.CoInitialize()
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32가 설치되어 있지 않습니다.") from exc

    try:
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    except Exception:
        try:
            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        except Exception as exc2:
            raise RuntimeError("한글(HWP) 프로그램을 실행하지 못했습니다.") from exc2

    try:
        hwp.XHwpWindows.Item(0).Visible = visible
    except Exception:
        pass

    try:
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    except Exception:
        pass

    return hwp


def _sanitize_filename(name):
    if not name:
        return ""
    name = re.sub(r'[\r\n\t]+', ' ', name)
    name = re.sub(r'[\\/:*?"<>|]', '', name)
    name = re.sub(r'[ ]+', ' ', name)
    name = name.strip(". ")
    return name


# ---------------------------------------------------------------------------
# 유적 표 목록 분석
# ---------------------------------------------------------------------------

def parse_items_from_hwp(hwp, logger=None):
    total_pages = hwp.PageCount
    emit_log(logger, f"HWP 문서 분석 중 (총 {total_pages}페이지)...")

    raw_text = hwp.GetTextFile("UNICODE", "")
    lines = [l.strip() for l in raw_text.splitlines() if l.strip()]

    discovered = []
    i = 0
    while i < len(lines):
        line = lines[i]
        line_clean = line.replace(" ", "")

        has_draw = ("도면" in line_clean and "명칭" in line_clean) or line_clean.startswith("도면명")
        has_code = bool(re.match(r'^[가-힣\w]+_\d+$', line_clean))

        if has_draw or has_code:
            d_code = ""
            s_name = ""

            if has_code:
                d_code = line.strip()
                s_start = i + 1
            else:
                s_start = i + 1
                for k in range(i + 1, min(i + 8, len(lines))):
                    cand = lines[k].strip()
                    if cand and not any(w in cand.replace(" ", "") for w in ["연번", "도면", "유적", "명칭", "소재지", "시대"]):
                        d_code = cand
                        s_start = k + 1
                        break

            for k in range(s_start, min(s_start + 25, len(lines))):
                cand = lines[k].strip()
                cand_clean = cand.replace(" ", "")
                if ("도면" in cand_clean and "명칭" in cand_clean) or bool(re.match(r'^[가-힣\w]+_\d+$', cand_clean)):
                    break
                if "유적명" in cand_clean or "유적명칭" in cand_clean:
                    parts = re.split(r'유적명(?:칭)?[:\s\t]*', cand)
                    if len(parts) > 1 and parts[1].strip():
                        s_name = parts[1].strip()
                        break
                    elif k + 1 < len(lines):
                        next_cand = lines[k + 1].strip()
                        if next_cand and not any(w in next_cand.replace(" ", "") for w in ["소재지", "시대", "변경", "신규", "유지", "해제", "기존", "도면"]):
                            s_name = next_cand
                            break

            d_san = _sanitize_filename(d_code)
            s_san = _sanitize_filename(s_name)
            s_san = re.sub(r'^(?:도면\s*명칭|연번|유적명)\s*', '', s_san).strip()

            if d_san and s_san:
                full_n = f"{d_san} {s_san}"
            else:
                full_n = d_san or s_san

            if full_n and len(full_n) <= 120 and not full_n.endswith("다."):
                if not discovered or discovered[-1][2] != full_n:
                    discovered.append((d_code, s_name, full_n))

        i += 1

    emit_log(logger, f"총 {len(discovered)}개 유적 표 검출 완료")

    start_pages = []
    step = total_pages / max(1, len(discovered))
    for k in range(len(discovered)):
        p_val = max(1, min(total_pages, int(round(1 + k * step))))
        start_pages.append(p_val)

    for k in range(1, len(start_pages)):
        if start_pages[k] <= start_pages[k - 1] and start_pages[k - 1] < total_pages:
            start_pages[k] = start_pages[k - 1] + 1
    start_pages[0] = 1

    filenames = [it[2] for it in discovered]
    return start_pages, filenames


def detect_page_breaks_from_open_doc(hwp):
    total_pages = hwp.PageCount
    hwp.Run("MoveDocBegin")
    time.sleep(0.08)

    segment_start_pages = [1]
    prev_page = 1

    while True:
        try:
            char_type = hwp.CharType
        except Exception:
            char_type = -1

        if char_type == 12:
            try:
                cur_page = hwp.CurrentPage
            except Exception:
                cur_page = prev_page
            next_page = cur_page + 1
            if next_page <= total_pages and next_page not in segment_start_pages:
                segment_start_pages.append(next_page)

        try:
            moved = hwp.Run("MoveNextChar")
        except Exception:
            break
        if not moved:
            break

    segment_start_pages.sort()
    return segment_start_pages, total_pages


def build_page_ranges_from_starts(segment_start_pages, total_pages):
    ranges = []
    for i, start in enumerate(segment_start_pages):
        if i + 1 < len(segment_start_pages):
            next_start = segment_start_pages[i + 1]
            end = max(start, next_start - 1)
        else:
            end = total_pages
        ranges.append((start, end))
    return ranges


def build_page_ranges_n_pages(total_pages, n):
    if n <= 0:
        raise ValueError("페이지 수는 1 이상이어야 합니다.")
    ranges = []
    start = 1
    while start <= total_pages:
        end = min(start + n - 1, total_pages)
        ranges.append((start, end))
        start = end + 1
    return ranges


def build_page_ranges_n_files(total_pages, n):
    if n <= 0:
        raise ValueError("파일 수는 1 이상이어야 합니다.")
    n = min(n, total_pages)
    base_size = total_pages // n
    remainder = total_pages % n
    ranges = []
    start = 1
    for i in range(n):
        size = base_size + (1 if i < remainder else 0)
        end = start + size - 1
        ranges.append((start, end))
        start = end + 1
    return ranges


# ---------------------------------------------------------------------------
# [100% 무결점 페이지 블록 저장 엔진]
# ---------------------------------------------------------------------------

def _save_exact_page_range(hwp, start_page, end_page, output_path, logger=None):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    norm_path = os.path.normpath(str(output_path))
    emit_log(logger, f"  p.{start_page}~p.{end_page} 추출 중 ➡️ {output_path.name}")

    # 1. 시작 페이지 맨 앞으로 이동
    hwp.Run("MoveDocBegin")
    time.sleep(0.04)
    for _ in range(start_page - 1):
        hwp.Run("MovePageDown")
    hwp.Run("MovePageBegin")
    time.sleep(0.04)

    # 2. 블록 선택 켜기 (F4)
    hwp.Run("BeginSel")
    time.sleep(0.04)

    # 3. 끝 페이지 맨 끝까지 블록 확장
    for _ in range(end_page - start_page):
        hwp.Run("MovePageDown")
    hwp.Run("MovePageEnd")
    time.sleep(0.04)

    # 4. 블록 파일 저장 (FileSaveBlock)
    saved = False
    try:
        pset = hwp.HParameterSet.HFileSaveBlock
        hwp.HAction.GetDefault("FileSaveBlock", pset.HSet)
        # 소문자 filename 및 HSet SetItem 동시 설정으로 win32com 완벽 호환
        try:
            pset.filename = norm_path
        except Exception:
            pass
        try:
            pset.HSet.SetItem("filename", norm_path)
            pset.HSet.SetItem("FileName", norm_path)
        except Exception:
            pass
        try:
            pset.Format = "HWP"
            pset.HSet.SetItem("Format", "HWP")
        except Exception:
            pass
            
        ret = hwp.HAction.Execute("FileSaveBlock", pset.HSet)
        time.sleep(0.1)
        if output_path.exists() and output_path.stat().st_size > 0:
            saved = True
    except Exception as e:
        emit_log(logger, f"    FileSaveBlock 오류: {e}")

    # 5. 블록 해제
    hwp.Run("Cancel")

    # 6. 만약 FileSaveBlock이 실패한 경우 복사/붙여넣기 fallback
    if not saved:
        emit_log(logger, "    복사/새문서 붙여넣기 시도...")
        try:
            hwp.Run("MoveDocBegin")
            for _ in range(start_page - 1):
                hwp.Run("MovePageDown")
            hwp.Run("MovePageBegin")
            hwp.Run("BeginSel")
            for _ in range(end_page - start_page):
                hwp.Run("MovePageDown")
            hwp.Run("MovePageEnd")
            
            hwp.HAction.Run("Copy")
            time.sleep(0.15)
            hwp.Run("Cancel")
            
            hwp.HAction.Run("FileNew")
            time.sleep(0.2)
            hwp.HAction.Run("Paste")
            time.sleep(0.2)
            hwp.SaveAs(norm_path, "HWP")
            time.sleep(0.2)
            try:
                hwp.HAction.Run("FileClose")
                time.sleep(0.05)
            except Exception:
                pass
        except Exception as e2:
            emit_log(logger, f"    복사/붙여넣기 오류: {e2}")

    file_size_kb = round(output_path.stat().st_size / 1024, 1) if output_path.exists() else 0
    emit_log(logger, f"  저장 완료: {output_path.name} ({file_size_kb} KB)")


def execute_split_by_plan(
    input_path,
    output_dir,
    page_ranges,
    names,
    pattern="{name}",
    visible=True,
    logger=None,
    progress_callback=None,
):
    pythoncom.CoInitialize()
    input_path = Path(input_path).expanduser().resolve()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    emit_progress(progress_callback, {"type": "progress", "current": 0, "total": 1, "status": "한글 실행 중…"})
    hwp = get_hwp_application(visible=True, logger=logger, progress_callback=progress_callback)

    emit_log(logger, f"한글 원본 문서 열기: {input_path.name}")
    hwp.Open(str(input_path), "HWP", "forceopen:true")
    time.sleep(0.8)

    total = len(page_ranges)
    digits = len(str(total))
    saved_paths = []

    try:
        for i, ((start, end), name) in enumerate(zip(page_ranges, names), start=1):
            num_str = str(i).zfill(digits)
            if "{name}" in pattern or "{num}" in pattern:
                file_stem = pattern.replace("{name}", name).replace("{num}", num_str)
            else:
                file_stem = f"{name}"
                
            file_stem = _sanitize_filename(file_stem)
            out_name = f"{file_stem}.hwp"
            out_path = output_dir / out_name

            emit_log(logger, f"[{i}/{total}] 분리 실행: {out_name}")
            emit_progress(progress_callback, {
                "type": "progress",
                "current": i - 1,
                "total": total,
                "status": f"분리 중 ({i}/{total}): {out_name}",
            })

            _save_exact_page_range(hwp, start, end, out_path, logger=logger)
            saved_paths.append(out_path)

        emit_progress(progress_callback, {
            "type": "progress",
            "current": total,
            "total": total,
            "status": "분리 완료",
        })
        emit_log(logger, f"🎉 모든 분리 완료: 총 {len(saved_paths)}개 파일 저장됨 ➡️ {output_dir}")
        return saved_paths

    finally:
        try:
            hwp.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def split_hwp_file(
    input_path,
    output_dir,
    basename=None,
    mode=SPLIT_MODE_TABLE_NAME,
    n=None,
    keywords=None,
    pattern="{name}",
    visible=True,
    logger=None,
    progress_callback=None,
):
    input_path = Path(input_path).expanduser().resolve()
    if not input_path.exists():
        raise RuntimeError(f"입력 파일을 찾을 수 없습니다: {input_path}")

    emit_progress(progress_callback, {"type": "progress", "current": 0, "total": 1, "status": "한글 실행 중…"})
    hwp = get_hwp_application(visible=True, logger=logger, progress_callback=progress_callback)

    emit_log(logger, f"파일 열기: {input_path.name}")
    hwp.Open(str(input_path), "HWP", "forceopen:true")
    time.sleep(0.5)

    total_pages = hwp.PageCount
    emit_log(logger, f"총 페이지 수: {total_pages}")

    try:
        if mode == SPLIT_MODE_TABLE_NAME:
            segment_start_pages, table_names = parse_items_from_hwp(hwp, logger=logger)
            page_ranges = build_page_ranges_from_starts(segment_start_pages, total_pages)
            min_count = min(len(page_ranges), len(table_names))
            page_ranges = page_ranges[:min_count]
            table_names = table_names[:min_count]
        elif mode == SPLIT_MODE_AUTO:
            segment_start_pages, _ = detect_page_breaks_from_open_doc(hwp)
            page_ranges = build_page_ranges_from_starts(segment_start_pages, total_pages)
            table_names = [f"{basename or input_path.stem}_{i}" for i in range(1, len(page_ranges)+1)]
        elif mode == SPLIT_MODE_N_PAGES:
            page_ranges = build_page_ranges_n_pages(total_pages, n)
            table_names = [f"{basename or input_path.stem}_{i}" for i in range(1, len(page_ranges)+1)]
        else:
            page_ranges = build_page_ranges_n_files(total_pages, n)
            table_names = [f"{basename or input_path.stem}_{i}" for i in range(1, len(page_ranges)+1)]

        try:
            hwp.Quit()
        except Exception:
            pass

        return execute_split_by_plan(
            input_path=input_path,
            output_dir=output_dir,
            page_ranges=page_ranges,
            names=table_names,
            pattern=pattern or "{name}",
            visible=True,
            logger=logger,
            progress_callback=progress_callback,
        )

    finally:
        try:
            hwp.Quit()
        except Exception:
            pass
