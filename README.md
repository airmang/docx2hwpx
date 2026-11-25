# 워드 → 한글 변환기

워드(.docx) 문서를 한글(.hwpx) 문서로 바꾸는 윈도우 프로그램입니다. 변환 엔진은 python-hwpx 7 입니다.

## 쓰는 법 (사용자)
1. `워드한글변환기.exe` 를 실행합니다. 설치가 필요 없습니다.
2. 워드 파일이나 폴더를 창에 끌어다 놓거나 [파일 추가]를 누릅니다.
3. 저장 위치를 고르고(기본: 원본과 같은 폴더) [변환]을 누릅니다.
4. 목록에서 파일을 누르면 한글로 옮기지 못한 부분이 아래 칸에 나옵니다. [한/글로 열기], [폴더 열기]로 바로 확인합니다.

같은 이름의 .hwpx 가 있으면 덮어쓰지 않고 `이름 (1).hwpx` 처럼 새 이름으로 저장합니다.

## 명령줄
```
python -m docx2hwpx                      # 창 열기
python -m docx2hwpx in.docx              # 변환 (같은 폴더, 이름 겹치면 (1), (2)…)
python -m docx2hwpx in.docx -o out.hwpx [--overwrite]
python -m docx2hwpx --version
```

## 개발
```
.venv\Scripts\python -m pytest            # 시험
.venv\Scripts\python packaging\screenshot.py   # out\gui.png (시연용 가짜 변환; --real 이면 실제 엔진)
powershell -ExecutionPolicy Bypass -File packaging\build.ps1   # dist\워드한글변환기.exe 만들기 + 자체 시험
```
- 창 없는 exe 의 점검: `워드한글변환기.exe --selftest 결과.txt [--selftest-docx 시험.docx]` 는 창을 만들었다 닫고, 엔진과 hwpx 하위 모듈을 불러와 결과를 파일에 적습니다.
- `src/docx2hwpx/api.py` 의 `convert()` 계약만 GUI 가 압니다. 변환 엔진은 `core.py`.
