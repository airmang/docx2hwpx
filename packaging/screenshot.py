"""실행 중인 창을 out/gui.png 로 저장한다 (개발용). 인자 --real 이면 실제 변환 엔진을 쓴다."""
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from PIL import ImageGrab

from docx2hwpx import gui
from docx2hwpx.api import ConversionResult, Loss

if "--real" not in sys.argv:
    def fake(source, output, *, overwrite=False, progress=None):
        time.sleep(0.15)
        n = source.name
        if "s3" in n:
            return ConversionResult(source, output, ok=False, error="표 구조를 읽지 못했습니다. (시연용 가짜 결과)")
        if "s5" in n:
            return ConversionResult(source, output, ok=True, losses=[
                Loss("textbox", "dropped", 2, "글상자 2개가 빠졌습니다"),
                Loss("equation", "changed", 1, "수식 1개가 그림으로 바뀌었습니다")])
        return ConversionResult(source, output, ok=True)
    gui.convert = fake

def grab_window(root):
    import ctypes
    from ctypes import wintypes
    from PIL import Image
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    hwnd = user32.GetParent(root.winfo_id()) or root.winfo_id()
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    w, h = r.right - r.left, r.bottom - r.top
    hdc = user32.GetWindowDC(hwnd)
    mdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    gdi32.SelectObject(mdc, bmp)
    user32.PrintWindow(hwnd, mdc, 2)  # PW_RENDERFULLCONTENT
    class BIH(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("a", wintypes.LONG), ("b", wintypes.LONG),
                    ("c", wintypes.DWORD), ("d", wintypes.DWORD)]
    bih = BIH(ctypes.sizeof(BIH), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bih), 0)
    gdi32.DeleteObject(bmp); gdi32.DeleteDC(mdc); user32.ReleaseDC(hwnd, hdc)
    return Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1).convert("RGB")


app = gui.App()
root = app.root
root.geometry("+40+40")
app.add_paths([str(ROOT / "tests" / "fixtures")])
root.update()
app.start()
t = time.time()
while app.running and time.time() - t < 60:
    root.update()
    time.sleep(0.02)
for _ in range(20):
    root.update()
    time.sleep(0.02)
sel = app.order[-1]
app.tree.selection_set(sel)
root.update()
time.sleep(0.4)
root.update()
x, y, w, h = root.winfo_rootx(), root.winfo_rooty(), root.winfo_width(), root.winfo_height()
(ROOT / "out").mkdir(exist_ok=True)
try:
    img = ImageGrab.grab((x - 8, y - 40, x + w + 8, y + h + 8))
except OSError:  # 화면을 직접 찍을 수 없는 세션이면 창 자체를 그려 받는다
    img = grab_window(root)
img.save(ROOT / "out" / "gui.png")
print("saved", w, h)
root.destroy()
