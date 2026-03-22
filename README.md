# HWP 폴더 병합 프로그램

입력 폴더 안의 `.hwp` 파일을 자연 정렬로 병합하는 도구입니다. GUI와 CLI 둘 다 사용할 수 있습니다.

## 준비

1. Windows에 한글(HWP)이 설치되어 있어야 합니다.
2. Python이 설치되어 있어야 합니다.
3. `pywin32`가 필요합니다.

```powershell
pip install pywin32
```

## GUI 실행

일반적인 데스크톱 UI로 입력 폴더와 출력 파일을 고를 수 있습니다.

```powershell
python hwp_merger_gui.py
```

또는 [run_merge_hwp.bat](C:/Users/nuri9/Desktop/hwpmerger/run_merge_hwp.bat)를 더블클릭해도 됩니다.

GUI 기능:

- 입력 폴더 선택
- 출력 파일 위치와 이름 지정
- 대상 `.hwp` 파일 개수 확인
- 중간 저장 간격 설정
- 병합 중 한글 창 표시 여부 선택
- 파일 사이 새 페이지 구분 여부 선택
- 진행 로그와 진행 상태 표시

## EXE 파일

빌드된 실행 파일:

- [HwpMerger.exe](C:/Users/nuri9/Desktop/hwpmerger/dist/HwpMerger.exe)

현재 빌드는 완료되어 있고, exe는 `dist` 폴더에 생성됩니다.

직접 다시 빌드하려면:

```powershell
pyinstaller --noconfirm HwpMerger.spec
```

또는 [build_exe.bat](C:/Users/nuri9/Desktop/hwpmerger/build_exe.bat)를 실행하면 됩니다.

## CLI 실행

기존처럼 명령줄에서도 실행할 수 있습니다.

```powershell
python merge_hwp.py "C:\Users\nuri9\Desktop\daejeon2-1"
```

출력 파일을 직접 지정:

```powershell
python merge_hwp.py "C:\Users\nuri9\Desktop\daejeon2-1" --output-file "D:\merged\전체병합.hwp"
```

한글 창을 보면서 실행:

```powershell
python merge_hwp.py "C:\Users\nuri9\Desktop\daejeon2-1" --show-hwp
```

파일 사이를 새 페이지로 구분:

```powershell
python merge_hwp.py "C:\Users\nuri9\Desktop\daejeon2-1" --page-break
```

## 파일 구성

- [hwp_merger_gui.py](C:/Users/nuri9/Desktop/hwpmerger/hwp_merger_gui.py): GUI 실행 파일
- [merge_hwp.py](C:/Users/nuri9/Desktop/hwpmerger/merge_hwp.py): CLI 실행 파일
- [hwp_merge_core.py](C:/Users/nuri9/Desktop/hwpmerger/hwp_merge_core.py): 공통 병합 로직
- [HwpMerger.spec](C:/Users/nuri9/Desktop/hwpmerger/HwpMerger.spec): PyInstaller 빌드 설정
- [build_exe.bat](C:/Users/nuri9/Desktop/hwpmerger/build_exe.bat): exe 재빌드 배치 파일

## 동작 방식

- `.hwp` 파일만 대상으로 합니다.
- `1`, `2`, `10` 같은 숫자 포함 파일명도 자연 정렬합니다.
- 첫 원본 파일을 결과 파일로 즉시 저장해서 원본이 덮어써지지 않게 처리합니다.
- 기본값은 파일을 그대로 이어 붙입니다.
- 필요하면 옵션으로 파일 사이에만 새 페이지 구분을 넣을 수 있습니다.
- 기본적으로 10개마다 중간 저장합니다.
