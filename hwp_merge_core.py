import re
import time
from pathlib import Path


OUTPUT_NAME_DEFAULT = "병합결과.hwp"
SAVE_INTERVAL_DEFAULT = 10


def natural_key(text):
    return [int(token) if token.isdigit() else token.lower() for token in re.split(r"(\d+)", text)]


def normalize_output_path(output_path):
    output_path = Path(output_path).expanduser().resolve()
    if output_path.suffix.lower() != ".hwp":
        output_path = output_path.with_suffix(".hwp")
    return output_path


def scan_hwp_files(input_dir):
    input_dir = Path(input_dir).expanduser().resolve()
    if not input_dir.exists() or not input_dir.is_dir():
        raise RuntimeError(f"입력 폴더를 찾을 수 없습니다: {input_dir}")

    return sorted(
        [path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() == ".hwp"],
        key=lambda path: natural_key(path.name),
    )


def get_hwp_files(input_dir, output_path=None):
    files = scan_hwp_files(input_dir)
    if output_path is None:
        return files

    normalized_output = normalize_output_path(output_path)
    return [path for path in files if path.resolve() != normalized_output]


def get_hwp_application(visible):
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32가 설치되어 있지 않습니다. `pip install pywin32` 후 다시 실행하세요.") from exc

    try:
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    except Exception as exc:
        raise RuntimeError("한글(HWP) COM 객체를 열지 못했습니다. 한글이 설치되어 있는지 확인해 주세요.") from exc

    try:
        hwp.XHwpWindows.Item(0).Visible = visible
    except Exception:
        pass

    try:
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    except Exception:
        pass

    return hwp


def emit_progress(progress_callback, payload):
    if progress_callback:
        progress_callback(payload)


def emit_log(logger, message):
    if logger:
        logger(message)
    else:
        print(message)


def prepare_boundary_for_insert(hwp, insert_page_break):
    hwp.Run("MoveDocEnd")
    time.sleep(0.15)

    if insert_page_break:
        hwp.HAction.Run("BreakPage")
        time.sleep(0.1)


def merge_hwp_files(
    input_dir,
    output_path,
    save_interval=SAVE_INTERVAL_DEFAULT,
    visible=False,
    insert_page_break=False,
    logger=None,
    progress_callback=None,
):
    input_dir = Path(input_dir).expanduser().resolve()
    output_path = normalize_output_path(output_path)

    if not input_dir.exists() or not input_dir.is_dir():
        raise RuntimeError(f"입력 폴더를 찾을 수 없습니다: {input_dir}")

    if not output_path.parent.exists():
        raise RuntimeError(f"출력 폴더를 찾을 수 없습니다: {output_path.parent}")

    all_hwp_files = scan_hwp_files(input_dir)
    matching_output = [path for path in all_hwp_files if path.resolve() == output_path]
    file_list = [path for path in all_hwp_files if path.resolve() != output_path]

    if not file_list:
        raise RuntimeError(f"병합할 .hwp 파일이 없습니다: {input_dir}")

    total_files = len(file_list)
    emit_progress(
        progress_callback,
        {
            "type": "scan_complete",
            "total": total_files,
            "input_dir": str(input_dir),
            "output_path": str(output_path),
        },
    )

    emit_log(logger, f"입력 폴더: {input_dir}")
    emit_log(logger, f"출력 파일: {output_path}")
    emit_log(logger, f"총 {total_files}개 파일 발견")

    if matching_output:
        emit_log(
            logger,
            "[주의] 출력 파일이 입력 폴더 안에 이미 있습니다. 이 파일은 입력 목록에서 제외하고 덮어씁니다.",
        )

    pythoncom = None
    try:
        import pythoncom as _pythoncom

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
    except ImportError:
        pythoncom = None

    hwp = get_hwp_application(visible=visible)
    success_count = 0
    fail_list = []

    try:
        first_file = file_list[0]
        emit_progress(
            progress_callback,
            {"type": "file_start", "index": 1, "total": total_files, "file_name": first_file.name},
        )
        emit_log(logger, f"\n[1/{total_files}] 첫 파일 열기: {first_file.name}")
        hwp.Open(str(first_file))
        time.sleep(1.5)

        if output_path.exists():
            output_path.unlink()

        hwp.SaveAs(str(output_path))
        time.sleep(0.8)
        success_count = 1
        emit_log(logger, "  결과 파일 생성 완료")
        emit_progress(
            progress_callback,
            {"type": "file_done", "index": 1, "total": total_files, "file_name": first_file.name, "success": True},
        )

        for idx, file_path in enumerate(file_list[1:], start=2):
            emit_progress(
                progress_callback,
                {"type": "file_start", "index": idx, "total": total_files, "file_name": file_path.name},
            )
            emit_log(logger, f"\n[{idx}/{total_files}] {file_path.name}")

            success = False
            try:
                prepare_boundary_for_insert(hwp, insert_page_break=insert_page_break)

                hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
                hwp.HParameterSet.HInsertFile.filename = str(file_path)
                hwp.HParameterSet.HInsertFile.KeepSection = 1
                result = hwp.HAction.Execute("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
                time.sleep(0.2)

                if result is False:
                    emit_log(logger, "  삽입 실패")
                    fail_list.append(file_path.name)
                else:
                    emit_log(logger, "  성공")
                    success_count += 1
                    success = True

                if save_interval > 0 and idx % save_interval == 0:
                    emit_log(logger, f"  중간 저장... ({idx}개 완료)")
                    emit_progress(progress_callback, {"type": "intermediate_save", "index": idx, "total": total_files})
                    hwp.Save()
                    time.sleep(0.6)

            except Exception as exc:
                emit_log(logger, f"  오류: {exc}")
                fail_list.append(file_path.name)

            emit_progress(
                progress_callback,
                {
                    "type": "file_done",
                    "index": idx,
                    "total": total_files,
                    "file_name": file_path.name,
                    "success": success,
                },
            )

        emit_log(logger, f"\n{'=' * 60}")
        emit_log(logger, "최종 저장 중...")
        emit_progress(progress_callback, {"type": "final_save", "total": total_files})
        hwp.Save()
        time.sleep(1.0)
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass

        if pythoncom is not None:
            pythoncom.CoUninitialize()

    if not output_path.exists():
        raise RuntimeError("결과 파일 저장에 실패했습니다.")

    file_size_mb = output_path.stat().st_size / (1024 * 1024)
    emit_log(logger, "병합 완료!")
    emit_log(logger, f"저장: {output_path}")
    emit_log(logger, f"크기: {file_size_mb:.2f} MB")
    emit_log(logger, f"성공: {success_count}/{total_files}개")

    if fail_list:
        emit_log(logger, f"\n실패 ({len(fail_list)}개):")
        for file_name in fail_list:
            emit_log(logger, f"  - {file_name}")

    result = {
        "output_path": output_path,
        "success_count": success_count,
        "total_files": total_files,
        "fail_list": fail_list,
        "file_size_mb": file_size_mb,
    }
    emit_progress(
        progress_callback,
        {
            "type": "completed",
            "total": total_files,
            "output_path": str(output_path),
            "success_count": success_count,
            "fail_count": len(fail_list),
        },
    )
    return result
