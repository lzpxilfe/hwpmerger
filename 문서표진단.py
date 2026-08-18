"""
HWP 문서 내 '천동' 및 '대전_068' 실제 원본 위치 및 표 텍스트 정밀 덤프 도구
"""
import sys
import re
import win32com.client
from tkinter import filedialog, Tk

root = Tk()
root.withdraw()

print("="*70)
print(" HWP 문서 실제 '천동' / '도면명칭' 원본 텍스트 구조 정밀 덤프")
print("="*70)

fpath = filedialog.askopenfilename(title="진단할 HWP 파일 선택", filetypes=[("HWP 파일", "*.hwp")])
if not fpath:
    print("파일이 선택되지 않았습니다.")
    sys.exit()

print(f"\n[1] 파일 열기: {fpath}")
hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
try:
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
except Exception:
    pass

hwp.Open(fpath, "HWP", "forceopen:true")
print(f"[2] 총 페이지: {hwp.PageCount}p")

raw_text = hwp.GetTextFile("UNICODE", "")
lines = raw_text.splitlines()
print(f"[3] 총 {len(lines)}줄 수집됨\n")

print("="*70)
print(" [4] 문서 내 '천동' 단어가 들어간 실제 줄과 앞뒤 전후 줄 덤프:")
print("="*70)

matched_indices = [i for i, l in enumerate(lines) if "천동" in l]

print(f"'천동' 발견 횟수: 총 {len(matched_indices)}개 줄\n")

for count, idx in enumerate(matched_indices, 1):
    print(f"--- [발견 #{count} (라인 {idx+1})] ---")
    # 앞 3줄부터 뒤 7줄까지 출력
    start = max(0, idx - 3)
    end = min(len(lines), idx + 8)
    for k in range(start, end):
        prefix = "👉 " if k == idx else "   "
        print(f"{prefix}줄 {k+1:04d}: {repr(lines[k])}")
    print()

hwp.Quit()
print("\n진단 완료. 터미널의 내용을 복사해서 알려주세요!")
input()
