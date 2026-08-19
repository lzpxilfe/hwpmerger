import os
import re
import time
import os
import json
import subprocess
import sys
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ElementTree
import pythoncom

# ---------------------------------------------------------------------------
# 분리 모드 상수
# ---------------------------------------------------------------------------
SPLIT_MODE_AUTO = "auto"
SPLIT_MODE_N_PAGES = "n_pages"
SPLIT_MODE_N_FILES = "n_files"
SPLIT_MODE_TABLE_NAME = "table_name"

# This source document's final form title is stored as a non-text object, so
# it is absent from GetTextFile even though the table itself is present.
MANUAL_FINAL_TABLE_NAME = "대전_068 대전 동구 천동2 주거환경개선 사업부지 내 유적"

# The native rhwp engine reads HWP's actual paragraph/table model.  The old
# GetTextFile + find-from-caret route could map two headings to one outer HWP
# control and then incorrectly report that the headings were out of order.
_APP_ROOTS = (
    Path(getattr(sys, "_MEIPASS", "")) if getattr(sys, "_MEIPASS", "") else None,
    Path(__file__).resolve().parent,
)
_RHWP_RELATIVE_PATH = Path("tools") / "rhwp" / "rhwp" / "rhwp.exe"
_A3_TEMPLATE_HWP_RELATIVE_PATH = Path("resources") / "a3_blank.hwp"
_A3_TEMPLATE_HWPX_RELATIVE_PATH = Path("resources") / "a3_blank.hwpx"


def emit_progress(progress_callback, payload):
    if progress_callback:
        progress_callback(payload)


def emit_log(logger, message):
    if logger:
        logger(message)
    else:
        print(message)


def _bundled_file(relative_path):
    for root in _APP_ROOTS:
        if root is None:
            continue
        candidate = root / relative_path
        if candidate.is_file():
            return candidate
    return None


def _rhwp_executable():
    executable = _bundled_file(_RHWP_RELATIVE_PATH)
    if executable is None:
        raise RuntimeError(
            "rhwp 분석 엔진을 찾지 못했습니다. tools\\rhwp\\rhwp\\rhwp.exe가 있어야 합니다."
        )
    return executable


def _a3_template_dimensions():
    template = _bundled_file(_A3_TEMPLATE_HWPX_RELATIVE_PATH)
    if template is None:
        raise RuntimeError(
            "A3 바탕 설정을 찾지 못했습니다. resources\\a3_blank.hwpx가 있어야 합니다."
        )
    try:
        with ZipFile(template) as archive:
            root = ElementTree.fromstring(archive.read("Contents/section0.xml"))
        for element in root.iter():
            if element.tag.endswith("pagePr"):
                return int(element.attrib["width"]), int(element.attrib["height"])
    except Exception as exc:
        raise RuntimeError("프로그램의 A3 바탕 설정을 읽지 못했습니다.") from exc
    raise RuntimeError("프로그램의 A3 바탕 설정에 용지 크기가 없습니다.")


def _a3_template_hwp_path():
    template = _bundled_file(_A3_TEMPLATE_HWP_RELATIVE_PATH)
    if template is None:
        raise RuntimeError(
            "A3 바탕 문서를 찾지 못했습니다. resources\\a3_blank.hwp가 있어야 합니다."
        )
    return template


def _run_rhwp_json(arguments, timeout=180):
    """Run rhwp without sending document text through the GUI/clipboard."""
    command = [str(_rhwp_executable()), *map(str, arguments)]
    try:
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("rhwp 구조 분석 시간이 초과되었습니다.") from exc
    except OSError as exc:
        raise RuntimeError(f"rhwp를 실행하지 못했습니다: {exc}") from exc

    if result.returncode:
        detail = (result.stderr or result.stdout or "알 수 없는 오류").strip()
        raise RuntimeError(f"rhwp 구조 분석 실패: {detail}")
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError("rhwp가 읽을 수 없는 분석 결과를 돌려주었습니다.") from exc


_RHWP_CODE_RE = re.compile(r"^[가-힣A-Za-z0-9]+_\d+(?:[·~\-]\d+)?$")


def _rhwp_record_from_table(table, table_index):
    """Recognise a survey header table directly from rhwp's table cells."""
    cells = [_clean_table_line(cell.get("text", "")) for cell in table.get("cells", [])]
    code_index = next(
        (index for index, value in enumerate(cells)
         if _RHWP_CODE_RE.fullmatch(value.replace(" ", ""))),
        None,
    )
    if code_index is None:
        return None

    site_name = ""
    for index, value in enumerate(cells):
        compact = value.replace(" ", "")
        if "유적명" not in compact and "유적명칭" not in compact:
            continue
        for candidate in cells[index + 1:index + 5]:
            candidate_compact = candidate.replace(" ", "")
            if (
                candidate
                and "유적" not in candidate_compact[:4]
                and not _is_label(candidate)
                and not _RHWP_CODE_RE.fullmatch(candidate_compact)
            ):
                site_name = candidate
                break
        if site_name:
            break

    code = _sanitize_filename(cells[code_index])
    site_name = _sanitize_filename(site_name)
    if not code or not site_name:
        return None
    return {
        "table_index": table_index,
        "name": f"{code} {site_name}",
        "section": table.get("section"),
        "paragraph": table.get("paragraph"),
    }


def analyze_items_with_rhwp(input_path, logger=None):
    """Return ordered logical records from HWP structure, without HWP UI text scans."""
    input_path = Path(input_path).expanduser().resolve()
    emit_log(logger, "rhwp로 표 구조를 분석하는 중 (본문 전체를 읽지 않습니다)...")
    payload = _run_rhwp_json(["export-tables", input_path, "--json"])
    tables = payload.get("tables")
    if not isinstance(tables, list) or not tables:
        raise RuntimeError("rhwp가 문서의 표 구조를 읽지 못했습니다.")

    records = []
    for table_index, table in enumerate(tables):
        record = _rhwp_record_from_table(table, table_index)
        if record:
            records.append(record)

    if not records:
        raise RuntimeError("rhwp가 [도면 명칭 + 유적명] 표를 찾지 못했습니다.")
    if any(right["table_index"] <= left["table_index"] for left, right in zip(records, records[1:])):
        raise RuntimeError("rhwp 표 인덱스가 역순으로 나와 자동 분리를 중단했습니다.")

    emit_log(
        logger,
        f"rhwp 표 구조 확인: 실제 표 {len(tables)}개, 유적 시작 표 {len(records)}개 검출 완료",
    )
    return tables, records


def parse_items_from_rhwp(input_path, logger=None):
    """Compatibility helper used by the preview UI for table-name mode."""
    _tables, records = analyze_items_with_rhwp(input_path, logger=logger)
    return [(0, 0)] * len(records), [record["name"] for record in records]


def get_hwp_application(visible=True, logger=None, progress_callback=None):
    pythoncom.CoInitialize()
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32가 설치되어 있지 않습니다.") from exc

    try:
        # Use a dedicated HWP process.  EnsureDispatch may attach to an
        # unrelated, modal HWP window and then the splitter appears frozen.
        hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
    except Exception:
        try:
            hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
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


def build_table_output_plan(names, pattern="{name}"):
    """Make deterministic, collision-free output names for recognised tables.

    Identical drawing-code/site-name pairs can legitimately occur when a
    cancelled record name is later reused.  The first keeps its ordinary name;
    later occurrences receive `` (2)``, `` (3)``, and so on.
    """
    digits = len(str(len(names))) if names else 1
    raw_stems = []
    base_counts = {}
    for index, name in enumerate(names, start=1):
        stem = pattern.replace("{name}", name).replace("{num}", str(index).zfill(digits))
        stem = _sanitize_filename(stem) or f"record-{index:0{digits}d}"
        key = stem.casefold()
        raw_stems.append((stem, key))
        base_counts[key] = base_counts.get(key, 0) + 1

    occurrences = {}
    used_keys = set()
    plan = []
    for index, (name, (base_stem, base_key)) in enumerate(zip(names, raw_stems), start=1):
        occurrence = occurrences.get(base_key, 0) + 1
        occurrences[base_key] = occurrence
        stem = base_stem if occurrence == 1 else f"{base_stem} ({occurrence})"
        suffix = occurrence
        while stem.casefold() in used_keys:
            suffix += 1
            stem = f"{base_stem} ({suffix})"
        used_keys.add(stem.casefold())
        plan.append({
            "index": index,
            "name": name,
            "stem": stem,
            "filename": f"{stem}.hwp",
            "base_key": base_key,
            "is_duplicate": base_counts[base_key] > 1,
        })
    return plan


# ---------------------------------------------------------------------------
# 유적 표 목록 분석
# ---------------------------------------------------------------------------

_TABLE_LABELS = (
    "도면명칭", "유적명", "유적명칭", "연번", "소재지", "시대", "정보변경",
    "현황", "변경", "신규", "유지", "해제", "기존", "유적설명", "참고문헌",
)
_DRAWING_CODE_RE = re.compile(r"^[가-힣A-Za-z0-9]+_\d+(?:[·ㆍ,-]\d+)?$")


def _clean_table_line(value):
    return re.sub(r"\s+", " ", (value or "").replace("\x00", " ")).strip()


def _is_label(value):
    compact = value.replace(" ", "")
    return not compact or any(label in compact for label in _TABLE_LABELS)


def _item_from_table_text(raw_text):
    """Return (drawing_code, site_name, filename) from one actual table/page.

    HWP tables are exported in slightly different row orders depending on how a
    table was authored.  This intentionally relies on the cell labels instead
    of a fixed line offset such as "the next 25 lines".
    """
    lines = [_clean_table_line(line) for line in (raw_text or "").splitlines()]
    lines = [line for line in lines if line]
    if not lines:
        return None

    drawing_code = ""
    for line in lines:
        compact = line.replace(" ", "")
        if _DRAWING_CODE_RE.fullmatch(compact):
            drawing_code = compact
            break

    site_name = ""
    for index, line in enumerate(lines):
        compact = line.replace(" ", "")
        label_match = re.match(r"^유적명(?:칭)?\s*[:：]?\s*(.*)$", line)
        if label_match and label_match.group(1).strip():
            site_name = label_match.group(1).strip()
            break
        if compact in {"유적명", "유적명칭"}:
            for candidate in lines[index + 1:index + 8]:
                if not _is_label(candidate) and not _DRAWING_CODE_RE.fullmatch(candidate.replace(" ", "")):
                    site_name = candidate
                    break
        if site_name:
            break

    # A few forms have no literal "유적명" cell, but put the item name right
    # after a drawing code.  Use that only as a guarded fallback.
    if not site_name and drawing_code:
        code_index = next(
            (i for i, line in enumerate(lines)
             if line.replace(" ", "") == drawing_code),
            -1,
        )
        for candidate in lines[code_index + 1:code_index + 18]:
            if not _is_label(candidate) and len(candidate) >= 3:
                site_name = candidate
                break

    drawing_code = _sanitize_filename(drawing_code)
    site_name = _sanitize_filename(site_name)
    if not site_name:
        return None
    filename = f"{drawing_code} {site_name}".strip()
    return drawing_code, site_name, filename


def _selected_page_text(hwp):
    """Read only the current page, not the entire document text stream."""
    hwp.Run("MovePageBegin")
    hwp.Run("BeginSel")
    try:
        hwp.Run("MovePageEnd")
        return hwp.GetTextFile("TEXT", "saveblock")
    finally:
        try:
            hwp.Run("Cancel")
        except Exception:
            pass


def _current_page_number(hwp):
    """Return the physical page where the caret is placed.

    ``CurrentPage`` is not exposed by the desktop HWP automation object.
    KeyIndicator's fourth value is the physical page number (1-based).
    """
    indicator = hwp.KeyIndicator()
    if not indicator or not indicator[0]:
        raise RuntimeError("한글에서 현재 쪽 번호를 읽지 못했습니다.")
    return int(indicator[3])


def _table_control_positions(hwp, keep_control=False):
    """Return source-table positions, optionally retaining their HWP anchors.

    Some HWP files have floating/legacy tables whose three-number cursor
    position is not sufficient after switching away from and back to the
    source tab.  The live control reference lets the saving phase ask HWP for
    the original anchor again.  Legacy callers keep receiving plain tuples.
    """
    positions = []
    ctrl = hwp.HeadCtrl
    while ctrl:
        if str(ctrl.CtrlID) == "tbl":
            try:
                anchor = ctrl.GetAnchorPos(0)
                hwp.SetPosBySet(anchor)
                position = tuple(hwp.GetPos())
                if keep_control:
                    positions.append({"position": position, "anchor": anchor})
                else:
                    positions.append(position)
            except Exception:
                pass
        ctrl = ctrl.Next
    return positions


def _select_table_control(hwp, table_position):
    """Select a table, retrying HWP's alternate control-selection route."""
    if isinstance(table_position, dict):
        plain_position = table_position["position"]
    else:
        plain_position = table_position

    hwp.SetPos(*plain_position)
    found = hwp.FindCtrl()
    if found == "tbl" or not isinstance(table_position, dict):
        return found

    # The ordinary position works for inline tables.  A floating or legacy
    # table needs the exact anchor supplied by HWP and then a direct control
    # selection.  Do not call FindCtrl() between SetPosBySet() and
    # SelectCtrlReverse(): that moves HWP away from the anchor in some files.
    try:
        try:
            hwp.Run("Cancel")
        except Exception:
            pass
        hwp.SetPosBySet(table_position["anchor"])
        # Some documents place the anchor immediately before a floating or
        # legacy table.  In that case FindCtrl() reports nothing, while HWP's
        # documented SelectCtrlReverse action selects the object behind the
        # anchor directly.
        hwp.Run("SelectCtrlReverse")
        selected = hwp.CurSelectedCtrl
        selected_id = str(getattr(selected, "CtrlID", "")) if selected else ""
        return selected_id or found
    except Exception:
        return found


def _table_text_from_position(hwp, position):
    """Return header-cell text from one table without reading document body."""
    hwp.SetPos(*position)
    page = _current_page_number(hwp)
    # MoveRight/TableCellBlock does not enter floating tables reliably.  Find
    # the control at its anchor and explicitly enter its first cell instead.
    hwp.FindCtrl()
    hwp.Run("ShapeObjTableSelCell")
    hwp.Run("Cancel")

    cells = []
    seen_cells = set()
    for _ in range(40):
        cell_id = hwp.KeyIndicator()[-1]
        if cell_id in seen_cells:
            break
        seen_cells.add(cell_id)
        hwp.Run("SelectAll")
        hwp.InitScan(0, 0x00ff, 0, 0, 0, 0)
        # The drawing code and site name are in the form header.  Do not scan
        # the long site-description cells that follow it: they are unrelated
        # to naming and make a 200-table document unnecessarily slow.
        chunks = []
        text_length = 0
        for _ in range(64):
            state, chunk = hwp.GetText()
            if chunk:
                chunks.append(chunk)
                text_length += len(chunk)
            if text_length >= 4096:
                break
            if state in (0, 1):
                break
        try:
            hwp.ReleaseScan()
        except Exception:
            pass
        hwp.Run("Cancel")
        cells.append("".join(chunks))
        if not hwp.HAction.Run("TableRightCell"):
            break
        hwp.Run("Cancel")
    return page, "\n".join(cells)


def _save_table_at_position(hwp, position, output_path):
    """Copy one table control at ``position`` into its own HWP document."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    hwp.SetPos(*position)
    hwp.Run("MoveRight")
    hwp.Run("TableCellBlock")
    hwp.Run("TableCellBlockExtend")
    hwp.Run("TableCellBlockExtend")
    hwp.Run("TableCellBlockExtend")
    try:
        hwp.HAction.Run("Copy")
    finally:
        hwp.Run("Cancel")

    try:
        hwp.HAction.Run("FileNew")
        time.sleep(0.12)
        hwp.HAction.Run("Paste")
        time.sleep(0.2)
        hwp.SaveAs(os.path.normpath(str(output_path)), "HWP")
        time.sleep(0.2)
    finally:
        hwp.HAction.Run("FileClose")
        time.sleep(0.06)

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError(f"표 저장 결과를 찾지 못했습니다: {output_path.name}")


def _read_current_page_setup(hwp):
    """Read the full paper size/orientation/margin setup at the source table."""
    action = hwp.CreateAction("PageSetup")
    settings = hwp.CreateSet("SecDef")
    action.GetDefault(settings)
    page = settings.Item("PageDef")
    fields = (
        "PaperWidth", "PaperHeight", "Landscape", "TopMargin", "BottomMargin",
        "LeftMargin", "RightMargin", "HeaderLen", "FooterLen", "GutterLen", "GutterType",
    )
    return {field: page.Item(field) for field in fields}


def _apply_page_setup(hwp, values):
    """Apply a source form's paper settings to the temporary output document."""
    action = hwp.CreateAction("PageSetup")
    settings = hwp.CreateSet("SecDef")
    action.GetDefault(settings)
    page = settings.Item("PageDef")
    for field, value in values.items():
        page.SetItem(field, value)
    # Apply the setting to the current section, not the whole source document.
    settings.SetItem("ApplyClass", 24)
    settings.SetItem("ApplyTo", 3)
    action.Execute(settings)


def _save_selected_table_control(hwp, output_path, source_size=0, logger=None):
    """Copy the selected HWP table object into a fresh document.

    This intentionally avoids ``TableCellBlock``.  That action is unreliable
    for floating/legacy tables and can cause HWP to copy the entire document.
    ``FindCtrl`` selects the actual ``tbl`` object, which is what Copy needs.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    selected = hwp.CurSelectedCtrl
    if selected is None or str(getattr(selected, "CtrlID", "")) != "tbl":
        raise RuntimeError("현재 위치에서 표 개체를 선택하지 못했습니다.")

    source_page_setup = _read_current_page_setup(hwp)

    try:
        emit_log(logger, "  1/4 표 개체를 클립보드로 복사 중...")
        hwp.HAction.Run("Copy")
        time.sleep(0.15)
        # XHwpDocuments.Add can block indefinitely on this installed HWP
        # build.  FileNew is less elegant but it reliably activates a blank
        # document for each output; it is closed immediately below.
        emit_log(logger, "  2/4 새 문서 준비 중...")
        hwp.HAction.Run("FileNew")
        time.sleep(0.15)
        emit_log(logger, "  3/4 원본 용지(A3 포함) 설정 적용 중...")
        _apply_page_setup(hwp, source_page_setup)
        emit_log(logger, "  4/4 표 붙여넣기 및 저장 중...")
        time.sleep(0.15)
        hwp.HAction.Run("Paste")
        time.sleep(0.25)
        hwp.SaveAs(os.path.normpath(str(output_path)), "HWP")
        time.sleep(0.25)
    except Exception as exc:
        raise RuntimeError(f"표 개체를 새 문서로 저장하지 못했습니다: {exc}") from exc
    finally:
        try:
            hwp.HAction.Run("FileClose")
            time.sleep(0.1)
        except Exception:
            pass

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError(f"표 결과 파일을 찾지 못했습니다: {output_path.name}")
    saved_size = output_path.stat().st_size
    if source_size and saved_size >= source_size * 0.25:
        raise RuntimeError(
            f"원본 전체가 저장된 것으로 보입니다 ({saved_size / 1024 / 1024:.1f} MB). "
            "이 결과는 사용하지 마세요."
        )
    paper_width_mm = source_page_setup["PaperWidth"] / 283.465
    paper_height_mm = source_page_setup["PaperHeight"] / 283.465
    emit_log(
        logger,
        f"  표 개체 저장 완료: {output_path.name} ({saved_size / 1024:.1f} KB, "
        f"용지 {paper_width_mm:.0f}×{paper_height_mm:.0f} mm)",
    )
    return output_path


def _save_table_object_at_position(hwp, position, output_path, source_size=0, logger=None):
    hwp.SetPos(*position)
    if hwp.FindCtrl() != "tbl":
        raise RuntimeError("마지막 표 개체를 찾지 못했습니다.")
    return _save_selected_table_control(hwp, output_path, source_size, logger)


def _save_table_object_by_site_name(hwp, site_name, output_path, source_size=0, logger=None):
    if not _find_first_text(hwp, site_name):
        raise RuntimeError(f"유적명을 문서에서 찾지 못했습니다: {site_name}")
    hwp.Run("Cancel")
    # At text inside a floating table, FindCtrl() alone does not select the
    # containing object.  ParentCtrl does: it exposes the table's anchor.
    parent = hwp.ParentCtrl
    if parent is None or str(getattr(parent, "CtrlID", "")) != "tbl":
        raise RuntimeError(f"'{site_name}'가 들어 있는 표 개체를 찾지 못했습니다.")
    hwp.SetPosBySet(parent.GetAnchorPos(0))
    if hwp.FindCtrl() != "tbl":
        raise RuntimeError(f"'{site_name}' 표 개체를 선택하지 못했습니다.")
    return _save_selected_table_control(hwp, output_path, source_size, logger)


def _execute_split_by_table_controls_legacy(input_path, output_dir, names, pattern="{name}", visible=False, logger=None, progress_callback=None):
    """Save the table containing each recognised site name as an HWP file."""
    input_path = Path(input_path).expanduser().resolve()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    targets = {}
    total = len(names)
    digits = len(str(total))
    for index, name in enumerate(names, start=1):
        stem = pattern.replace("{name}", name).replace("{num}", str(index).zfill(digits))
        targets[_sanitize_filename(name)] = output_dir / f"{_sanitize_filename(stem)}.hwp"

    emit_log(logger, "한글 전용 작업 창을 시작하는 중...")
    # Saving must be visible: an HWP version/security dialog in a hidden
    # window otherwise makes the GUI look permanently stuck.
    hwp = get_hwp_application(visible=True, logger=logger, progress_callback=progress_callback)
    saved = []
    try:
        emit_log(logger, "원본 HWP를 여는 중...")
        hwp.Open(str(input_path), "HWP", "forceopen:true")
        emit_log(logger, f"인식된 표 {total}개를 이름으로 찾아 분리합니다...")
        final_table_position = None
        for index, name in enumerate(names, start=1):
            # Filenames are built as "drawing_code site_name".  The code
            # never contains a space, so the remainder is the exact table
            # value used to place the caret before copying the whole table.
            parts = name.split(" ", 1)
            site_name = parts[1] if len(parts) == 2 else name
            key = _sanitize_filename(name)
            output_path = targets.pop(key)
            emit_progress(progress_callback, {
                "type": "progress", "current": index - 1, "total": total,
                "status": f"표 분리 중 ({index}/{total}): {output_path.name}",
            })
            if name == MANUAL_FINAL_TABLE_NAME:
                if final_table_position is None:
                    positions = _table_control_positions(hwp)
                    if not positions:
                        raise RuntimeError("마지막 표 개체를 찾지 못했습니다.")
                    final_table_position = positions[-1]
                emit_log(logger, "  마지막 비텍스트 표를 천동2 항목으로 저장합니다.")
                _save_table_object_at_position(
                    hwp, final_table_position, output_path, input_path.stat().st_size, logger
                )
            else:
                _save_table_object_by_site_name(
                    hwp, site_name, output_path, input_path.stat().st_size, logger
                )
            saved.append(output_path)
        emit_progress(progress_callback, {"type": "progress", "current": total, "total": total, "status": "분리 완료"})
        return saved
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass


def _table_items_from_controls(hwp, logger=None):
    """Return candidates from actual table controls, including nonstandard forms."""
    found = []
    skipped = 0
    first_error = None
    try:
        positions = _table_control_positions(hwp)
        for index, position in enumerate(positions, start=1):
            if index == 1 or index % 25 == 0:
                emit_log(logger, f"표 개체 확인 중: {index}/{len(positions)}")
            try:
                page, table_text = _table_text_from_position(hwp, position)
                item = _item_from_table_text(table_text)
                if item:
                    found.append((page, *item))
            except Exception as exc:
                skipped += 1
                first_error = first_error or str(exc)
    except Exception as exc:
        emit_log(logger, f"표 개체 목록을 읽지 못했습니다. 페이지 분석으로 계속합니다: {exc}")
    if skipped:
        emit_log(logger, f"표 개체 {skipped}개는 직접 읽지 못했습니다. 검색 방식으로 보완합니다: {first_error}")
    return found


def _items_from_document_text(raw_text):
    """Read table names from the complete text stream in document order."""
    lines = [_clean_table_line(line) for line in (raw_text or "").splitlines()]
    header_indexes = [
        index for index, line in enumerate(lines)
        if "도면" in line.replace(" ", "") and "명칭" in line.replace(" ", "")
    ]
    items = []
    for number, start in enumerate(header_indexes):
        end = header_indexes[number + 1] if number + 1 < len(header_indexes) else len(lines)
        item = _item_from_table_text("\n".join(lines[start:end]))
        if item and (not items or items[-1][2] != item[2]):
            items.append(item)
    return items


def _find_pages_for_text(hwp, find_text):
    """Find every occurrence of text and return its real, physical page."""
    pages = []
    positions = set()
    hwp.Run("MoveDocBegin")
    while True:
        pset = hwp.HParameterSet.HFindReplace
        hwp.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.ReplaceString = ""
        pset.FindString = find_text
        pset.Direction = hwp.FindDir("Forward")
        pset.IgnoreMessage = 1
        pset.FindType = 1
        if not hwp.HAction.Execute("RepeatFind", pset.HSet):
            break
        position = tuple(hwp.GetPos())
        if position in positions:
            break
        positions.add(position)
        pages.append(_current_page_number(hwp))
        # RepeatFind normally advances from the selected match.  Moving one
        # character guarantees progress on HWP versions that keep it selected.
        hwp.Run("MoveRight")
    return pages


def _find_first_text(hwp, find_text):
    """Place the caret on the first matching text without showing a dialog."""
    hwp.Run("MoveDocBegin")
    pset = hwp.HParameterSet.HFindReplace
    hwp.HAction.GetDefault("RepeatFind", pset.HSet)
    pset.ReplaceString = ""
    pset.FindString = find_text
    pset.Direction = hwp.FindDir("Forward")
    pset.IgnoreMessage = 1
    pset.FindType = 1
    return bool(hwp.HAction.Execute("RepeatFind", pset.HSet))


def _save_table_by_site_name(hwp, site_name, output_path, logger=None):
    """Copy the complete table containing ``site_name`` into a new HWP file.

    This deliberately does not use a page range.  Several forms can begin on
    the same page, and FileSaveBlock may save the entire source document.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    norm_path = os.path.normpath(str(output_path))

    if not _find_first_text(hwp, site_name):
        raise RuntimeError(f"유적명을 문서에서 찾지 못했습니다: {site_name}")

    try:
        # RepeatFind leaves the matching text selected.  Clear that selection
        # while keeping the caret in its cell, then extend to the whole table.
        hwp.Run("Cancel")
        hwp.Run("TableCellBlock")
        hwp.Run("TableCellBlockExtend")
        hwp.Run("TableCellBlockExtend")
        hwp.Run("TableCellBlockExtend")
        hwp.HAction.Run("Copy")
        time.sleep(0.15)
    except Exception as exc:
        raise RuntimeError(f"'{site_name}' 표 전체를 선택하지 못했습니다: {exc}") from exc
    finally:
        try:
            hwp.Run("Cancel")
        except Exception:
            pass

    try:
        hwp.HAction.Run("FileNew")
        time.sleep(0.15)
        hwp.HAction.Run("Paste")
        time.sleep(0.25)
        hwp.SaveAs(norm_path, "HWP")
        time.sleep(0.25)
    except Exception as exc:
        raise RuntimeError(f"'{site_name}' 새 문서 저장 실패: {exc}") from exc
    finally:
        try:
            hwp.HAction.Run("FileClose")
            time.sleep(0.1)
        except Exception:
            pass

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError(f"'{site_name}' 결과 HWP 파일을 찾지 못했습니다.")
    emit_log(logger, f"  표 저장 완료: {output_path.name} ({round(output_path.stat().st_size / 1024, 1)} KB)")


def build_table_split_plan(hwp, logger=None):
    """Build a split plan from real table/page positions.

    The old implementation found names in a flat text dump and then spread
    them evenly over the document pages.  That makes a plausible-looking but
    wrong result.  This routine records the page where each actual table lives
    and never fabricates a page number.
    """
    total_pages = hwp.PageCount
    emit_log(logger, "[빌드] 다중 표 묶음 + A3·용량 검증 v18 (2026-08-19)")
    emit_log(logger, f"HWP 표 이름 분석 중 (총 {total_pages}페이지)...")

    # This legacy HWP exposes form text in the document text stream but not
    # through individual table controls.  Read it once for names only; page
    # ranges and body text are never used for the actual split.
    items = _items_from_document_text(hwp.GetTextFile("UNICODE", ""))
    if not any(filename == MANUAL_FINAL_TABLE_NAME for _, _, filename in items):
        items.append(("대전_068", "대전 동구 천동2 주거환경개선 사업부지 내 유적", MANUAL_FINAL_TABLE_NAME))
        emit_log(logger, "[확인 필요] 마지막 천동2 제목은 문서 텍스트에 없어 사용자가 지정한 이름으로 추가했습니다.")
    plan = [
        {"page": 0, "drawing_code": code, "site_name": site, "filename": filename}
        for code, site, filename in items
    ]

    if not plan:
        raise RuntimeError(
            "표 이름을 찾지 못했습니다. 이 문서는 표를 이미지/그리기 개체로 저장했을 수 있습니다."
        )

    emit_log(logger, f"표 이름 {len(plan)}개 인식 완료 (실제 저장은 표 단위)")
    return plan


def parse_items_from_hwp(hwp, logger=None):
    """Compatibility wrapper used by the GUI and CLI."""
    plan = build_table_split_plan(hwp, logger=logger)
    return [item["page"] for item in plan], [item["filename"] for item in plan]


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

    # 4. Copy the selected pages, create a new document, paste, and save it.
    # FileSaveBlock and FileSaveBlock_S can ignore the selection on some HWP
    # builds and save the complete source document instead.
    try:
        hwp.HAction.Run("Copy")
        time.sleep(0.15)
    finally:
        hwp.Run("Cancel")

    try:
        hwp.HAction.Run("FileNew")
        time.sleep(0.2)
        hwp.HAction.Run("Paste")
        time.sleep(0.25)
        hwp.SaveAs(norm_path, "HWP")
        time.sleep(0.25)
    except Exception as exc:
        raise RuntimeError(f"새 문서로 저장하지 못했습니다: {exc}") from exc
    finally:
        # FileClose returns to the still-open source document, ready for the
        # next range.  It also prevents the next copy from using pasted output.
        try:
            hwp.HAction.Run("FileClose")
            time.sleep(0.1)
        except Exception:
            pass

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError("새 문서 저장 후 결과 HWP 파일을 찾지 못했습니다.")

    file_size_kb = round(output_path.stat().st_size / 1024, 1) if output_path.exists() else 0
    emit_log(logger, f"  저장 완료: {output_path.name} ({file_size_kb} KB)")


def _repeat_find_from_caret(hwp, find_text):
    """Find text forward from the current caret without opening a dialog."""
    pset = hwp.HParameterSet.HFindReplace
    hwp.HAction.GetDefault("RepeatFind", pset.HSet)
    pset.ReplaceString = ""
    pset.FindString = find_text
    pset.Direction = hwp.FindDir("Forward")
    pset.IgnoreMessage = 1
    pset.FindType = 1
    return bool(hwp.HAction.Execute("RepeatFind", pset.HSet))


def _logical_record_starts(hwp, names, logger=None):
    """Locate logical record starts in document order, without page guessing.

    A record begins at its recognised form title and ends immediately before
    the next form title.  It may contain more than one physical HWP table.
    """
    starts = []
    hwp.Run("MoveDocBegin")
    for name in names:
        if name == MANUAL_FINAL_TABLE_NAME:
            positions = _table_control_positions(hwp)
            if not positions:
                raise RuntimeError("텍스트 밖의 마지막 표 개체를 찾지 못했습니다.")
            starts.append(positions[-1])
            continue
        parts = name.split(" ", 1)
        site_name = parts[1] if len(parts) == 2 else name
        if not _repeat_find_from_caret(hwp, site_name):
            raise RuntimeError(f"분리 시작점 유적명을 찾지 못했습니다: {site_name}")
        hwp.Run("Cancel")
        starts.append(tuple(hwp.GetPos()))
        # Start the next search after this title; do not repeatedly find the
        # current form's first occurrence.
        hwp.Run("MoveRight")
    if len(starts) != len(names):
        raise RuntimeError("표 시작점 개수와 파일명 개수가 다릅니다.")
    emit_log(logger, f"논리 표 시작점 {len(starts)}개 확인 완료")
    return starts


def _save_logical_record(hwp, start, end, output_path, source_size, logger=None):
    """Save the source block directly, without using the system clipboard."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    hwp.SetPos(*start)
    source_page_setup = _read_current_page_setup(hwp)
    emit_log(logger, "  1/2 다음 제목 전까지의 표 묶음 선택 중...")
    hwp.Run("BeginSel")
    hwp.SetPos(*end)
    # The next record's first character must not be included in this output.
    hwp.Run("MoveLeft")
    try:
        # FileSaveBlock_S is HWP's own "Save Block" command.  Unlike
        # Copy/FileNew/Paste, it never touches the Windows clipboard and
        # writes only the active selection to a standalone HWP file.
        emit_log(logger, "  2/2 선택된 표 묶음을 별도 HWP로 저장 중...")
        pset = hwp.HParameterSet.HFileOpenSave
        hwp.HAction.GetDefault("FileSaveBlock_S", pset.HSet)
        pset.filename = os.path.normpath(str(output_path))
        pset.Format = "HWP"
        pset.Attributes = 1
        if not hwp.HAction.Execute("FileSaveBlock_S", pset.HSet):
            raise RuntimeError("한글 블록 저장 명령이 실행되지 않았습니다.")
    finally:
        hwp.Run("Cancel")

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError(f"분리 결과 파일을 찾지 못했습니다: {output_path.name}")
    saved_size = output_path.stat().st_size
    if source_size and saved_size >= source_size * 0.25:
        raise RuntimeError(
            f"원본 전체가 저장된 것으로 보입니다 ({saved_size / 1024 / 1024:.1f} MB). "
            "이 결과는 사용하지 마세요."
        )
    paper_width_mm = source_page_setup["PaperWidth"] / 283.465
    paper_height_mm = source_page_setup["PaperHeight"] / 283.465
    emit_log(
        logger,
        f"  표 묶음 저장 완료: {output_path.name} ({saved_size / 1024:.1f} KB, "
        f"용지 {paper_width_mm:.0f}×{paper_height_mm:.0f} mm)",
    )


def _execute_split_by_logical_blocks_unsupported(input_path, output_dir, names, pattern="{name}", visible=False, logger=None, progress_callback=None):
    """Split logical form blocks: title start through the next title start."""
    input_path = Path(input_path).expanduser().resolve()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    total = len(names)
    digits = len(str(total))
    emit_log(logger, "한글 전용 작업 창을 시작하는 중...")
    hwp = get_hwp_application(visible=True, logger=logger, progress_callback=progress_callback)
    saved = []
    try:
        emit_log(logger, "원본 HWP를 여는 중...")
        hwp.Open(str(input_path), "HWP", "forceopen:true")
        emit_log(logger, f"인식된 표 {total}개를 시작점-다음 시작점 범위로 분리합니다...")
        starts = _logical_record_starts(hwp, names, logger)
        hwp.Run("MoveDocEnd")
        document_end = tuple(hwp.GetPos())
        for index, name in enumerate(names, start=1):
            stem = pattern.replace("{name}", name).replace("{num}", str(index).zfill(digits))
            output_path = output_dir / f"{_sanitize_filename(stem)}.hwp"
            end = starts[index] if index < total else document_end
            emit_progress(progress_callback, {
                "type": "progress", "current": index - 1, "total": total,
                "status": f"분리 중 ({index}/{total}): {output_path.name}",
            })
            _save_logical_record(hwp, starts[index - 1], end, output_path, input_path.stat().st_size, logger)
            saved.append(output_path)
        emit_progress(progress_callback, {"type": "progress", "current": total, "total": total, "status": "분리 완료"})
        return saved
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass


def _select_table_by_site_name(hwp, site_name):
    """Select the one real table control that contains ``site_name``."""
    if not _find_first_text(hwp, site_name):
        raise RuntimeError(f"유적명을 문서에서 찾지 못했습니다: {site_name}")
    hwp.Run("Cancel")
    parent = hwp.ParentCtrl
    if parent is None or str(getattr(parent, "CtrlID", "")) != "tbl":
        raise RuntimeError(f"'{site_name}'가 들어 있는 표 개체를 찾지 못했습니다.")
    hwp.SetPosBySet(parent.GetAnchorPos(0))
    if hwp.FindCtrl() != "tbl":
        raise RuntimeError(f"'{site_name}' 표 개체를 선택하지 못했습니다.")


def _save_selected_table_in_fresh_hwp(source_hwp, output_path, source_size, logger=None):
    """Paste one selected table into a temporary tab in the same HWP process."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    selected = source_hwp.CurSelectedCtrl
    if selected is None or str(getattr(selected, "CtrlID", "")) != "tbl":
        raise RuntimeError("저장할 표 개체가 선택되지 않았습니다.")
    page_setup = _read_current_page_setup(source_hwp)
    # Some HWP 2018 installations block a second HWPFrame instance.  Use a
    # temporary tab in the same process, with explicit document activation.
    # The tab is created before Copy, keeping the clipboard handoff atomic.
    source_document = source_hwp.XHwpDocuments.Active_XHwpDocument
    source_document_id = source_document.DocumentID
    emit_log(logger, "  임시 저장 탭 준비 중...")
    try:
        destination_document = source_hwp.XHwpDocuments.Add(True)
        destination_document.SetActive_XHwpDocument()
        _apply_page_setup(source_hwp, page_setup)
        source_document.SetActive_XHwpDocument()
        # Do not yield, log, or create a document between these two calls.
        source_hwp.HAction.Run("Copy")
        destination_document.SetActive_XHwpDocument()
        source_hwp.HAction.Run("Paste")
        source_hwp.SaveAs(os.path.normpath(str(output_path)), "HWP")
        time.sleep(0.2)
    finally:
        try:
            # Close the exact temporary document.  Closing the collection's
            # active item can close the source instead on HWP 2018.
            destination_document.Modified = False
            destination_document.Close(False)
            source_hwp.XHwpDocuments.FindItem(source_document_id).SetActive_XHwpDocument()
        except Exception:
            pass

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError(f"표 결과 파일을 찾지 못했습니다: {output_path.name}")
    saved_size = output_path.stat().st_size
    if source_size and saved_size >= source_size * 0.25:
        raise RuntimeError(
            f"원본 전체가 저장된 것으로 보입니다 ({saved_size / 1024 / 1024:.1f} MB). "
            "이 결과는 사용하지 마세요."
        )
    paper_width_mm = page_setup["PaperWidth"] / 283.465
    paper_height_mm = page_setup["PaperHeight"] / 283.465
    emit_log(
        logger,
        f"  표 개체 저장 완료: {output_path.name} ({saved_size / 1024:.1f} KB, "
        f"용지 {paper_width_mm:.0f}×{paper_height_mm:.0f} mm)",
    )


def _execute_single_table_legacy(input_path, output_dir, names, pattern="{name}", visible=False, logger=None, progress_callback=None):
    """Save actual HWP table controls, one standalone HWP per recognised name."""
    input_path = Path(input_path).expanduser().resolve()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    emit_log(logger, "한글 원본 작업 창을 시작하는 중...")
    source_hwp = get_hwp_application(visible=True, logger=logger, progress_callback=progress_callback)
    saved = []
    try:
        emit_log(logger, "원본 HWP를 여는 중...")
        source_hwp.Open(str(input_path), "HWP", "forceopen:true")
        emit_log(logger, f"인식된 표 {total}개를 실제 표 개체로 저장합니다...")
        last_table_position = None
        for index, name in enumerate(names, start=1):
            stem = pattern.replace("{name}", name).replace("{num}", str(index).zfill(digits))
            output_path = output_dir / f"{_sanitize_filename(stem)}.hwp"
            emit_progress(progress_callback, {
                "type": "progress", "current": index - 1, "total": total,
                "status": f"표 저장 중 ({index}/{total}): {output_path.name}",
            })
            if name == MANUAL_FINAL_TABLE_NAME:
                if last_table_position is None:
                    positions = _table_control_positions(source_hwp)
                    if not positions:
                        raise RuntimeError("텍스트 밖의 마지막 표 개체를 찾지 못했습니다.")
                    last_table_position = positions[-1]
                source_hwp.SetPos(*last_table_position)
                if source_hwp.FindCtrl() != "tbl":
                    raise RuntimeError("마지막 표 개체를 선택하지 못했습니다.")
            else:
                parts = name.split(" ", 1)
                _select_table_by_site_name(source_hwp, parts[1] if len(parts) == 2 else name)
            _save_selected_table_in_fresh_hwp(
                source_hwp, output_path, input_path.stat().st_size, logger
            )
            saved.append(output_path)
        emit_progress(progress_callback, {"type": "progress", "current": total, "total": total, "status": "분리 완료"})
        return saved
    finally:
        try:
            source_hwp.Quit()
        except Exception:
            pass


def _control_anchor_key(ctrl):
    anchor = ctrl.GetAnchorPos(0)
    return (int(anchor.Item("List")), int(anchor.Item("Para")), int(anchor.Item("Pos")))


def _table_groups_for_names_legacy(hwp, names, logger=None):
    """Map every recognised title to all consecutive physical tables it owns."""
    table_positions = []
    table_index = {}
    ctrl = hwp.HeadCtrl
    while ctrl:
        if str(ctrl.CtrlID) == "tbl":
            key = _control_anchor_key(ctrl)
            table_index.setdefault(key, len(table_positions))
            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            table_positions.append(tuple(hwp.GetPos()))
        ctrl = ctrl.Next
    if not table_positions:
        raise RuntimeError("문서 안의 표 개체를 찾지 못했습니다.")

    title_indexes = []
    hwp.Run("MoveDocBegin")
    for name in names:
        if name == MANUAL_FINAL_TABLE_NAME:
            title_indexes.append(len(table_positions) - 1)
            continue
        parts = name.split(" ", 1)
        site_name = parts[1] if len(parts) == 2 else name
        if not _repeat_find_from_caret(hwp, site_name):
            raise RuntimeError(f"분리 시작 유적명을 찾지 못했습니다: {site_name}")
        hwp.Run("Cancel")
        parent = hwp.ParentCtrl
        if parent is None or str(getattr(parent, "CtrlID", "")) != "tbl":
            raise RuntimeError(f"'{site_name}'의 시작 표 개체를 찾지 못했습니다.")
        index = table_index.get(_control_anchor_key(parent))
        if index is None:
            raise RuntimeError(f"'{site_name}' 시작 표의 문서 위치를 찾지 못했습니다.")
        title_indexes.append(index)
        hwp.Run("MoveRight")

    if any(right <= left for left, right in zip(title_indexes, title_indexes[1:])):
        raise RuntimeError("표 시작점 순서가 뒤섞여 자동 분리를 중단했습니다.")

    groups = []
    for index, start in enumerate(title_indexes):
        end = title_indexes[index + 1] if index + 1 < len(title_indexes) else len(table_positions)
        groups.append(table_positions[start:end])
    emit_log(logger, f"제목 {len(names)}개 ↔ 실제 표 {len(table_positions)}개를 {len(groups)}개 묶음으로 연결했습니다.")
    return groups


def _table_groups_for_names(input_path, hwp, preview_names=None, logger=None):
    """Map all record groups using rhwp's true physical table stream.

    A heading search in the HWP UI may return the same outer control for two
    nested tables.  That is not a reversed record order.  rhwp assigns every
    table a stable ordinal, so titles and photo tables can be grouped without
    scanning or selecting the document body text.
    """
    rhwp_tables, records = analyze_items_with_rhwp(input_path, logger=logger)
    table_positions = _table_control_positions(hwp, keep_control=True)
    if not table_positions:
        raise RuntimeError("한글에서 표 개체를 찾지 못했습니다.")
    if len(table_positions) != len(rhwp_tables):
        raise RuntimeError(
            "rhwp와 한글이 인식한 실제 표 개수가 다릅니다 "
            f"(rhwp {len(rhwp_tables)}개 / 한글 {len(table_positions)}개). "
            "이 상태에서는 빠진 표가 생길 수 있어 저장하지 않았습니다."
        )

    planned_names = [record["name"] for record in records]
    if preview_names and list(preview_names) != planned_names:
        emit_log(logger, "[안내] 이전 미리보기 대신 rhwp가 방금 읽은 표 순서로 분리합니다.")

    groups = []
    for index, record in enumerate(records):
        start = record["table_index"]
        end = records[index + 1]["table_index"] if index + 1 < len(records) else len(table_positions)
        positions = table_positions[start:end]
        if not positions:
            raise RuntimeError(f"'{record['name']}'의 표 묶음이 비어 있습니다.")
        groups.append(positions)

    emit_log(
        logger,
        f"rhwp 기준 유적 시작점 {len(records)}개를 실제 표 {len(table_positions)}개의 묶음으로 연결했습니다.",
    )
    return planned_names, groups


def _count_top_level_tables(hwp):
    count = 0
    ctrl = hwp.HeadCtrl
    while ctrl:
        if str(ctrl.CtrlID) == "tbl":
            count += 1
        ctrl = ctrl.Next
    return count


def _save_table_group_in_tab(source_hwp, positions, output_path, source_size, logger=None):
    """Save all table controls in one logical record into one A3 HWP file."""
    if not positions:
        raise RuntimeError("저장할 표 개체 묶음이 비어 있습니다.")
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    source_document = source_hwp.XHwpDocuments.Active_XHwpDocument
    source_document_id = source_document.DocumentID
    source_hwp.SetPos(*positions[0])
    source_page_setup = _read_current_page_setup(source_hwp)
    emit_log(logger, f"  표 개체 {len(positions)}개를 같은 파일에 저장 중...")

    destination_document = None
    try:
        destination_document = source_hwp.XHwpDocuments.Add(True)
        destination_document.SetActive_XHwpDocument()
        for number, position in enumerate(positions, start=1):
            source_hwp.XHwpDocuments.FindItem(source_document_id).SetActive_XHwpDocument()
            source_hwp.SetPos(*position)
            if source_hwp.FindCtrl() != "tbl":
                raise RuntimeError(f"묶음의 {number}번째 표 개체를 선택하지 못했습니다.")
            source_hwp.HAction.Run("Copy")
            destination_document.SetActive_XHwpDocument()
            if number > 1:
                source_hwp.Run("MoveDocEnd")
                source_hwp.Run("BreakPara")
            source_hwp.HAction.Run("Paste")

        destination_document.SetActive_XHwpDocument()
        # Apply the source A3 setup after pasting so pasted content cannot
        # replace the destination section's paper definition with A4.
        _apply_page_setup(source_hwp, source_page_setup)
        applied_page_setup = _read_current_page_setup(source_hwp)
        if (
            applied_page_setup["PaperWidth"] != source_page_setup["PaperWidth"]
            or applied_page_setup["PaperHeight"] != source_page_setup["PaperHeight"]
        ):
            raise RuntimeError("출력 탭의 용지 크기가 원본 A3 설정과 다릅니다. 저장하지 않습니다.")
        saved_table_count = _count_top_level_tables(source_hwp)
        if saved_table_count < len(positions):
            raise RuntimeError(
                f"표 {len(positions)}개 중 {saved_table_count}개만 결과 탭에 들어갔습니다. 저장하지 않습니다."
            )
        source_hwp.SaveAs(os.path.normpath(str(output_path)), "HWP")
        time.sleep(0.2)
    finally:
        try:
            if destination_document is not None:
                destination_document.Modified = False
                destination_document.Close(False)
            source_hwp.XHwpDocuments.FindItem(source_document_id).SetActive_XHwpDocument()
        except Exception:
            pass

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError(f"표 묶음 결과 파일을 찾지 못했습니다: {output_path.name}")
    saved_size = output_path.stat().st_size
    if source_size and saved_size >= source_size * 0.25:
        raise RuntimeError(f"원본 전체가 저장된 것으로 보입니다 ({saved_size / 1024 / 1024:.1f} MB).")
    if saved_size < 96 * 1024:
        raise RuntimeError(
            f"결과가 {saved_size / 1024:.1f} KB뿐이라 표 내용 또는 사진이 빠진 것으로 보입니다. "
            "정상 분리로 처리하지 않습니다."
        )
    paper_width_mm = source_page_setup["PaperWidth"] / 283.465
    paper_height_mm = source_page_setup["PaperHeight"] / 283.465
    emit_log(
        logger,
        f"  표 묶음 저장 완료: {output_path.name} ({saved_size / 1024:.1f} KB, "
        f"표 {len(positions)}개, 용지 {paper_width_mm:.0f}×{paper_height_mm:.0f} mm)",
    )


def _verify_saved_table_bundle(output_path, expected_name, expected_table_count):
    """Fail closed when a saved file lost its header or a photo table."""
    payload = _run_rhwp_json(["export-tables", output_path, "--json"])
    tables = payload.get("tables") or []
    found_names = {
        record["name"]
        for index, table in enumerate(tables)
        for record in [_rhwp_record_from_table(table, index)]
        if record
    }
    if expected_name not in found_names:
        raise RuntimeError(f"저장본에 '{expected_name}' 제목 표가 없습니다.")
    if len(tables) < expected_table_count:
        raise RuntimeError(
            f"저장본 표가 {len(tables)}개뿐입니다. "
            f"원래 묶음의 {expected_table_count}개 표가 모두 저장되지 않았습니다."
        )
    return len(tables)


def _has_valid_existing_table_bundle(output_path, expected_name, expected_table_count):
    """Return true only for an existing output that still passes rhwp checks."""
    output_path = Path(output_path)
    if not output_path.is_file() or output_path.stat().st_size == 0:
        return False
    try:
        _verify_saved_table_bundle(output_path, expected_name, expected_table_count)
    except Exception:
        return False
    return True


def _save_table_group_in_tab(source_hwp, positions, output_path, source_size, expected_name, logger=None):
    """Paste a group into an A3-configured tab, then validate the saved HWP."""
    if not positions:
        raise RuntimeError("저장할 표 묶음이 비어 있습니다.")
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    working_path = output_path.with_name(f".{output_path.stem}.partial{output_path.suffix}")
    if working_path.exists():
        working_path.unlink()

    source_document = source_hwp.XHwpDocuments.Active_XHwpDocument
    source_document_id = source_document.DocumentID
    destination_document = None
    a3_page_setup = None
    try:
        # HWP 2018 reuses the active tab for Open().  First create a blank tab,
        # then load the real A3 template into that tab.  The source therefore
        # remains a separate document and the destination has an actual A3
        # section before any table is pasted.
        source_document.SetActive_XHwpDocument()
        a3_width, a3_height = _a3_template_dimensions()
        source_hwp.HAction.Run("FileNew")
        time.sleep(0.15)
        blank_document = source_hwp.XHwpDocuments.Active_XHwpDocument
        if blank_document.DocumentID == source_document_id:
            raise RuntimeError("A3 출력용 새 탭을 만들지 못했습니다.")
        source_hwp.Open(str(_a3_template_hwp_path()), "HWP", "forceopen:true")
        time.sleep(0.15)
        destination_document = source_hwp.XHwpDocuments.Active_XHwpDocument
        if destination_document.DocumentID == source_document_id:
            raise RuntimeError("A3 출력용 새 탭을 별도로 만들지 못했습니다.")
        a3_page_setup = _read_current_page_setup(source_hwp)
        if (
            abs(a3_page_setup["PaperWidth"] - a3_width) > 10
            or abs(a3_page_setup["PaperHeight"] - a3_height) > 10
        ):
            raise RuntimeError("A3 바탕 문서가 297×420 mm로 열리지 않았습니다.")
        if blank_document.DocumentID != destination_document.DocumentID:
            blank_document.Modified = False
            blank_document.Close(False)

        emit_log(logger, f"  A3 바탕 탭에 표 {len(positions)}개를 넣는 중...")
        for number, position in enumerate(positions, start=1):
            source_document.SetActive_XHwpDocument()
            found_control = _select_table_control(source_hwp, position)
            if found_control != "tbl":
                raise RuntimeError(
                    f"묶음의 {number}번째 표 개체를 선택하지 못했습니다 "
                    f"(한글 응답: {found_control or '없음'})."
                )
            source_hwp.HAction.Run("Copy")
            destination_document.SetActive_XHwpDocument()
            source_hwp.Run("MoveDocEnd")
            if number > 1:
                source_hwp.Run("BreakPara")
            source_hwp.HAction.Run("Paste")

        destination_document.SetActive_XHwpDocument()
        applied_page_setup = _read_current_page_setup(source_hwp)
        if (
            abs(applied_page_setup["PaperWidth"] - a3_width) > 10
            or abs(applied_page_setup["PaperHeight"] - a3_height) > 10
        ):
            # A whole-section paste is unusual for a selected tbl object, but
            # recover once if this HWP build carried a page definition across.
            _apply_page_setup(source_hwp, a3_page_setup)
            applied_page_setup = _read_current_page_setup(source_hwp)
            if (
                abs(applied_page_setup["PaperWidth"] - a3_width) > 10
                or abs(applied_page_setup["PaperHeight"] - a3_height) > 10
            ):
                raise RuntimeError("표를 붙인 뒤 출력 문서가 A3가 아니어서 저장하지 않았습니다.")
        # HeadCtrl omits nested photo tables on some forms.  The saved file is
        # verified below with rhwp, which counts every real table and rejects
        # an incomplete bundle without this false early failure.
        source_hwp.SaveAs(os.path.normpath(str(working_path)), "HWP")
        time.sleep(0.2)
    except Exception:
        try:
            if working_path.exists():
                working_path.unlink()
        except OSError:
            pass
        raise
    finally:
        try:
            if destination_document is not None:
                destination_document.Modified = False
                destination_document.Close(False)
            source_hwp.XHwpDocuments.FindItem(source_document_id).SetActive_XHwpDocument()
        except Exception:
            pass

    try:
        if not working_path.exists() or working_path.stat().st_size == 0:
            raise RuntimeError("A3 임시 저장본을 찾지 못했습니다.")
        saved_size = working_path.stat().st_size
        if source_size and saved_size >= source_size * 0.25:
            raise RuntimeError(
                f"원본 전체가 복사된 것으로 보입니다 ({saved_size / 1024 / 1024:.1f} MB)."
            )
        saved_table_count = _verify_saved_table_bundle(
            working_path, expected_name, len(positions)
        )
        os.replace(working_path, output_path)
    except Exception:
        try:
            if working_path.exists():
                working_path.unlink()
        except OSError:
            pass
        raise

    emit_log(
        logger,
        f"  A3 표 묶음 저장 완료: {output_path.name} ({saved_size / 1024:.1f} KB, "
        f"표 {saved_table_count}개, 용지 297×420 mm)",
    )
    return output_path


def execute_split_by_table_controls(input_path, output_dir, names, pattern="{name}", visible=False, logger=None, progress_callback=None):
    """Split every logical record, including its 2–3 constituent tables."""
    input_path = Path(input_path).expanduser().resolve()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    total = len(names)
    digits = len(str(total))
    emit_log(logger, "한글 원본 작업 창을 시작하는 중...")
    hwp = get_hwp_application(visible=True, logger=logger, progress_callback=progress_callback)
    saved = []
    try:
        emit_log(logger, "원본 HWP를 여는 중...")
        hwp.Open(str(input_path), "HWP", "forceopen:true")
        planned_names, groups = _table_groups_for_names(
            input_path, hwp, preview_names=names, logger=logger
        )
        total = len(planned_names)
        output_plan = build_table_output_plan(planned_names, pattern)
        duplicate_groups_missing_output = {
            item["base_key"]
            for item in output_plan
            if item["is_duplicate"]
            and not all(
                (output_dir / candidate["filename"]).is_file()
                for candidate in output_plan
                if candidate["base_key"] == item["base_key"]
            )
        }
        for item, positions in zip(output_plan, groups):
            index = item["index"]
            name = item["name"]
            output_path = output_dir / item["filename"]
            emit_progress(progress_callback, {
                "type": "progress", "current": index - 1, "total": total,
                "status": f"표 묶음 저장 중 ({index}/{total}, 표 {len(positions)}개): {output_path.name}",
            })
            rebuild_duplicate_group = item["base_key"] in duplicate_groups_missing_output
            if not rebuild_duplicate_group and _has_valid_existing_table_bundle(output_path, name, len(positions)):
                emit_log(logger, f"  기존 검산 통과 파일 유지: {output_path.name}")
                saved.append(output_path)
                continue
            if rebuild_duplicate_group:
                emit_log(logger, f"  동명 유적 묶음 재생성: {output_path.name}")
            _save_table_group_in_tab(
                hwp,
                positions,
                output_path,
                input_path.stat().st_size,
                expected_name=name,
                logger=logger,
            )
            saved.append(output_path)
        emit_progress(progress_callback, {"type": "progress", "current": total, "total": total, "status": "분리 완료"})
        return saved
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass


def execute_split_by_plan(
    input_path,
    output_dir,
    page_ranges,
    names,
    pattern="{name}",
    visible=True,
    logger=None,
    progress_callback=None,
    split_by_table=False,
):
    if split_by_table:
        return execute_split_by_table_controls(
            input_path=input_path,
            output_dir=output_dir,
            names=names,
            pattern=pattern,
            visible=visible,
            logger=logger,
            progress_callback=progress_callback,
        )
    pythoncom.CoInitialize()
    input_path = Path(input_path).expanduser().resolve()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    emit_progress(progress_callback, {"type": "progress", "current": 0, "total": 1, "status": "한글 실행 중…"})
    # Saving is intentionally visible: HWP may show an overwrite/security
    # confirmation, and a hidden modal dialog otherwise looks like a failed
    # split with no output files.
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

    # Table-name mode is completely rhwp-driven: do not open HWP merely to
    # scan its body text before the real table-copy phase begins.
    if mode == SPLIT_MODE_TABLE_NAME:
        page_ranges, table_names = parse_items_from_rhwp(input_path, logger=logger)
        return execute_split_by_plan(
            input_path=input_path,
            output_dir=output_dir,
            page_ranges=page_ranges,
            names=table_names,
            pattern=pattern or "{name}",
            visible=visible,
            logger=logger,
            progress_callback=progress_callback,
            split_by_table=True,
        )

    emit_progress(progress_callback, {"type": "progress", "current": 0, "total": 1, "status": "한글 실행 중…"})
    hwp = get_hwp_application(visible=visible, logger=logger, progress_callback=progress_callback)

    emit_log(logger, f"파일 열기: {input_path.name}")
    hwp.Open(str(input_path), "HWP", "forceopen:true")
    time.sleep(0.5)

    total_pages = hwp.PageCount
    emit_log(logger, f"총 페이지 수: {total_pages}")

    try:
        if mode == SPLIT_MODE_AUTO:
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
            visible=visible,
            logger=logger,
            progress_callback=progress_callback,
            split_by_table=False,
        )

    finally:
        try:
            hwp.Quit()
        except Exception:
            pass
