from pathlib import Path
import sys
import time

import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageGrab

ROOT_DIR = Path(__file__).resolve().parent.parent
ASSETS_DIR = ROOT_DIR / "assets"
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from hwp_merger_gui import HwpMergerApp

def load_font(size, bold=False):
    candidates = []
    if bold:
        candidates.extend(["malgunbd.ttf", "seguisb.ttf", "arialbd.ttf"])
    candidates.extend(["malgun.ttf", "segoeui.ttf", "arial.ttf"])

    for name in candidates:
        try:
            return ImageFont.truetype(name, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


TITLE_FONT = load_font(34, bold=True)
SUBTITLE_FONT = load_font(18)
CAPTION_FONT = load_font(16)


def ensure_assets_dir():
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)


def clear_log(app):
    app.log_text.configure(state="normal")
    app.log_text.delete("1.0", "end")
    app.log_text.configure(state="disabled")


def fill_demo_state(app, stage):
    app.input_dir_var.set(r"C:\Users\nuri9\Desktop\daejeon2-1")
    app.output_path_var.set(r"C:\Users\nuri9\Desktop\daejeon2-1\병합결과.hwp")
    app.file_count_var.set("대상 파일: 24개 .hwp 파일")
    app.save_interval_var.set(10)
    app.show_hwp_var.set(False)
    app.insert_page_break_var.set(False)
    app.last_output_path = Path(r"C:\Users\nuri9\Desktop\daejeon2-1\병합결과.hwp")
    clear_log(app)

    if stage == "setup":
        app.status_var.set("병합 준비 중")
        app.progressbar["maximum"] = 24
        app.progressbar["value"] = 0
        app.append_log("======================================================================")
        app.append_log("새 병합 작업을 시작합니다.")
        app.append_log("입력 폴더: C:\\Users\\nuri9\\Desktop\\daejeon2-1")
        app.append_log("출력 파일: C:\\Users\\nuri9\\Desktop\\daejeon2-1\\병합결과.hwp")
        app.append_log("총 24개 파일 발견")
        return "폴더와 출력 파일을 고른 뒤 바로 병합을 시작할 수 있습니다."

    if stage == "progress":
        app.status_var.set("[9/24] 처리 중: 계약서9.hwp")
        app.progressbar["maximum"] = 24
        app.progressbar["value"] = 9
        app.append_log("======================================================================")
        app.append_log("새 병합 작업을 시작합니다.")
        app.append_log("입력 폴더: C:\\Users\\nuri9\\Desktop\\daejeon2-1")
        app.append_log("출력 파일: C:\\Users\\nuri9\\Desktop\\daejeon2-1\\병합결과.hwp")
        app.append_log("총 24개 파일 발견")
        app.append_log("")
        app.append_log("[1/24] 첫 파일 열기: 계약서1.hwp")
        app.append_log("  결과 파일 생성 완료")
        app.append_log("")
        app.append_log("[2/24] 계약서2.hwp")
        app.append_log("  성공")
        app.append_log("")
        app.append_log("[9/24] 계약서9.hwp")
        return "진행 로그와 상태를 한 화면에서 확인할 수 있습니다."

    app.status_var.set("병합 완료: 24/24개 성공, 결과 병합결과.hwp")
    app.progressbar["maximum"] = 24
    app.progressbar["value"] = 24
    app.open_output_button.configure(state="normal")
    app.append_log("======================================================================")
    app.append_log("새 병합 작업을 시작합니다.")
    app.append_log("입력 폴더: C:\\Users\\nuri9\\Desktop\\daejeon2-1")
    app.append_log("출력 파일: C:\\Users\\nuri9\\Desktop\\daejeon2-1\\병합결과.hwp")
    app.append_log("총 24개 파일 발견")
    app.append_log("")
    app.append_log("[24/24] 계약서24.hwp")
    app.append_log("  성공")
    app.append_log("")
    app.append_log("최종 저장 중...")
    app.append_log("병합 완료!")
    app.append_log("저장: C:\\Users\\nuri9\\Desktop\\daejeon2-1\\병합결과.hwp")
    app.append_log("성공: 24/24개")
    return "완료 후에는 결과 폴더를 바로 열 수 있습니다."


def capture_window(root):
    root.update_idletasks()
    root.update()
    time.sleep(0.2)
    x = root.winfo_rootx()
    y = root.winfo_rooty()
    width = root.winfo_width()
    height = root.winfo_height()
    return ImageGrab.grab(bbox=(x, y, x + width, y + height)).convert("RGBA")


def rounded_mask(size, radius):
    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, size[0], size[1]), radius=radius, fill=255)
    return mask


def decorate_capture(capture, caption):
    canvas = Image.new("RGBA", (1100, 860), "#f6efe5")
    draw = ImageDraw.Draw(canvas)

    for bbox, color in [
        ((-40, -20, 420, 300), "#f7c9a9"),
        ((760, 80, 1120, 410), "#c9ddd2"),
        ((720, 620, 1120, 920), "#f3d4be"),
    ]:
        overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
        overlay_draw = ImageDraw.Draw(overlay)
        overlay_draw.ellipse(bbox, fill=color)
        overlay = overlay.filter(ImageFilter.GaussianBlur(28))
        canvas.alpha_composite(overlay)

    draw.text((72, 56), "HWP Merger", font=TITLE_FONT, fill="#2b241f")
    draw.text((72, 106), "한글 문서를 한 번에, 더 보기 좋게.", font=SUBTITLE_FONT, fill="#65584d")

    shadow = Image.new("RGBA", (capture.width + 36, capture.height + 36), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle((18, 18, capture.width + 18, capture.height + 18), radius=26, fill=(58, 41, 28, 90))
    shadow = shadow.filter(ImageFilter.GaussianBlur(16))
    canvas.alpha_composite(shadow, (74, 162))

    card = Image.new("RGBA", (capture.width, capture.height), (255, 255, 255, 255))
    card.putalpha(rounded_mask(card.size, 22))
    framed_capture = Image.new("RGBA", card.size, (255, 255, 255, 0))
    framed_capture.paste(capture, (0, 0))
    framed_capture.putalpha(rounded_mask(card.size, 22))
    card.alpha_composite(framed_capture)
    canvas.alpha_composite(card, (92, 180))

    draw.rounded_rectangle((72, 736, 1028, 804), radius=20, fill=(255, 250, 245, 205))
    draw.text((96, 756), caption, font=CAPTION_FONT, fill="#4b3f36")
    return canvas.convert("P")


def generate_assets():
    ensure_assets_dir()

    root = tk.Tk()
    style = ttk.Style()
    try:
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    root.geometry("780x620+120+80")
    root.attributes("-topmost", True)
    app = HwpMergerApp(root)

    frames = []
    preview_frame = None
    for stage in ["setup", "progress", "done"]:
        caption = fill_demo_state(app, stage)
        capture = capture_window(root)
        decorated = decorate_capture(capture, caption)
        frames.append(decorated)
        if stage == "progress":
            preview_frame = decorated

    preview_path = ASSETS_DIR / "readme-preview.png"
    gif_path = ASSETS_DIR / "readme-demo.gif"
    preview_frame.save(preview_path)
    frames[0].save(
        gif_path,
        save_all=True,
        append_images=frames[1:],
        duration=[1000, 1200, 1600],
        loop=0,
        optimize=False,
    )
    root.destroy()
    print(preview_path)
    print(gif_path)


if __name__ == "__main__":
    generate_assets()
