# 워드 → 한글 변환기 (docx2hwpx)

워드(.docx) 문서를 한글(.hwpx) 문서로 바꾸는 윈도우 프로그램입니다. 한컴오피스가 없어도 바꿀 수 있고, 바꾼 파일은 한/글에서 바로 열립니다. 변환 엔진은 [python-hwpx](https://github.com/airmang/python-hwpx) 7입니다.

## 내려받기
[Releases](https://github.com/airmang/docx2hwpx/releases)에서 `워드한글변환기.exe`를 받아 실행합니다. 설치할 것이 없습니다.

- Windows 10/11 64비트에서 돌아갑니다.
- 서명하지 않은 프로그램이라, 처음 실행할 때 "Windows의 PC 보호" 창이 뜰 수 있습니다. [추가 정보] → [실행]을 누르면 됩니다.
- 실행 파일 하나가 압축을 풀면서 켜지므로 처음 뜨는 데 몇 초 걸립니다.

## 쓰는 법
1. 워드 파일이나 폴더를 창에 끌어다 놓거나 [파일 추가]·[폴더 추가]를 누릅니다.
2. 저장 위치를 고릅니다(기본: 원본과 같은 폴더).
3. [변환]을 누릅니다. 목록에서 파일을 누르면 한글로 옮기지 못한 부분이 아래 칸에 나옵니다.
4. [한/글로 열기], [폴더 열기]로 결과를 확인합니다.

원본 워드 파일은 건드리지 않습니다. 같은 이름의 .hwpx가 이미 있으면 덮어쓰지 않고 `이름 (1).hwpx`처럼 새 이름으로 저장합니다. 문서는 내 PC 안에서만 바뀌고 어디로도 보내지 않습니다.

## 옮겨지는 것
- 글자 모양: 글꼴, 크기, 굵게·기울임·밑줄·취소선, 색, 형광펜, 위·아래 첨자
- 문단 모양: 정렬, 들여쓰기·내어쓰기, 줄 간격, 문단 위·아래 간격, 쪽 나눔, 문단 테두리·배경, 탭
- 제목 스타일(개요 수준), 글머리표·번호 목록
- 표: 칸 합치기, 칸 배경, 테두리(없는 표 포함), 열 너비, 표 안의 표
- 그림(본문에 넣은 그림, 떠 있는 그림의 위치)
- 쪽: 용지 크기·방향(가로 구역 포함), 여백, 다단, 구역 나누기
- 머리말·꼬리말, 쪽 번호, 각주·미주, 하이퍼링크

## 옮기지 못하는 것
아래는 한글 문서로 옮기지 못하거나 모양이 바뀝니다. 그런 부분이 있으면 프로그램이 파일마다 알려 줍니다. 말없이 사라지지는 않습니다.
- 글상자(글은 본문 문단으로 꺼냄), 도형, 차트, SmartArt, 메모
- 수식(글자로 남음)
- 변경 내용 추적(받아들인 최종본으로 바뀜)
- 첫 쪽만 다른 머리말, 쪽 테두리, 줄 번호, 세로 쓰기 칸

한/글이 워드 문서를 직접 열었을 때와 쪽 수·모양이 같게 나오도록 맞췄습니다. 그래도 문서마다 다를 수 있으니 중요한 문서는 한/글에서 열어 확인해 주세요. 이상한 결과가 나오면 [Issues](https://github.com/airmang/docx2hwpx/issues)에 알려 주세요. 개인정보가 든 문서는 올리지 마세요.

## 명령줄 (개발자용)
```
python -m docx2hwpx                      # 창 열기
python -m docx2hwpx in.docx              # 변환 (같은 폴더, 이름 겹치면 (1), (2)…)
python -m docx2hwpx in.docx -o out.hwpx [--overwrite]
python -m docx2hwpx --version
```

## 개발
```
uv venv -p 3.12 .venv
uv pip install -p .venv -e . pytest pyinstaller
.venv\Scripts\python -m pytest                                     # 시험
powershell -ExecutionPolicy Bypass -File packaging\build.ps1       # dist\워드한글변환기.exe 만들기 + 자체 시험
```
- 창 없는 exe의 점검: `워드한글변환기.exe --selftest 결과.txt [--selftest-docx 시험.docx]`
- GUI는 `src/docx2hwpx/api.py`의 `convert()` 계약만 압니다. 변환은 `core.py` → `docx_reader.py`(읽기) → `writer.py`(쓰기)입니다.
- 시험 문서(`tests/fixtures`)는 `tools/make_samples.py`로 만든 것이고, 글은 모두 지어낸 것입니다.

## 사용권
Apache License 2.0 (`LICENSE`). 함께 들어간 소프트웨어는 `THIRD_PARTY_NOTICES.md`를 보세요.
