# 깨끗한 상태에서 dist\워드한글변환기.exe 를 만든다.
#   powershell -ExecutionPolicy Bypass -File packaging\build.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
$py = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) { throw ".venv 가 없습니다. 먼저 가상환경을 만드세요." }

foreach ($d in "build", "dist") { if (Test-Path $d) { Remove-Item $d -Recurse -Force } }
& $py packaging\make_assets.py
if ($LASTEXITCODE) { throw "아이콘/버전 정보 생성 실패" }
& $py -m PyInstaller packaging\docx2hwpx.spec --noconfirm --clean --distpath dist --workpath build
if ($LASTEXITCODE) { throw "PyInstaller 실패" }

$exe = Get-ChildItem dist -Filter *.exe | Select-Object -First 1
if (-not $exe) { throw "exe 가 만들어지지 않았습니다." }
"{0}  {1:N1} MB" -f $exe.FullName, ($exe.Length / 1MB)

# 점검: 창을 만들었다 닫고 엔진을 불러와 보는 자체 시험 (창 없는 exe 라서 결과를 파일로 받는다)
$res = Join-Path $env:TEMP "docx2hwpx_selftest.txt"
if (Test-Path $res) { Remove-Item $res }
$sw = [Diagnostics.Stopwatch]::StartNew()
$p = Start-Process -FilePath $exe.FullName -ArgumentList "--selftest", "`"$res`"" -PassThru -Wait
$sw.Stop()
"자체 시험 종료 코드 $($p.ExitCode), $([int]$sw.ElapsedMilliseconds) ms"
if (Test-Path $res) { Get-Content $res -Encoding UTF8 } else { throw "자체 시험 결과 파일이 없습니다." }
if ($p.ExitCode -ne 0) { throw "자체 시험 실패" }
