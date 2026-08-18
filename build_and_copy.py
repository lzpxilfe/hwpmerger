import os
import shutil
import sys

target_dir = r"C:\Users\nuri9\Desktop\hwpmerger"
dist_exe = os.path.join(target_dir, "dist", "HWP병합분리기.exe")
root_exe = os.path.join(target_dir, "HWP병합분리기.exe")
hwpmerger_exe = os.path.join(target_dir, "HwpMerger.exe")

print("Checking dist_exe:", os.path.exists(dist_exe))

if os.path.exists(dist_exe):
    try:
        shutil.copy2(dist_exe, root_exe)
        shutil.copy2(dist_exe, hwpmerger_exe)
        print("SUCCESS: Copied dist_exe to root folder!")
    except Exception as e:
        print("Copy failed:", e)
else:
    print("dist_exe not found, running PyInstaller directly via Python API...")
    import PyInstaller.__main__
    
    gui_script = os.path.join(target_dir, "hwp_merger_gui.py")
    PyInstaller.__main__.run([
        '--onefile',
        '--windowed',
        '--noconfirm',
        '--name=HWP병합분리기',
        f'--distpath={target_dir}',
        gui_script
    ])
    if os.path.exists(root_exe):
        shutil.copy2(root_exe, hwpmerger_exe)
        print("SUCCESS: Built directly to root folder!")

print("\n--- Current files in target_dir ---")
for f in os.listdir(target_dir):
    p = os.path.join(target_dir, f)
    if os.path.isfile(p):
        print(f"{f} ({os.path.getsize(p)} bytes)")
