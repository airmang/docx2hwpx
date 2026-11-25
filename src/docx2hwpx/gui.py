"""워드 → 한글 변환기: 창 화면 (tkinter + ttk + tkinterdnd2).

화면 문구는 모두 한국어다. 변환은 별도 스레드에서 돌리고, 화면 갱신은 메인 스레드에서만 한다.
"""
from __future__ import annotations

import ctypes
import os
import queue
import sys
import threading
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Iterable

from . import __version__
from .api import ConversionResult, convert

APP_TITLE = "워드 → 한글 변환기"
FONT_FAMILY = "Malgun Gothic"

ST_WAIT = "대기"
ST_RUN = "변환 중"
ST_OK = "완료"
ST_PART = "일부 바뀜"
ST_FAIL = "실패"

SAME_FOLDER_LABEL = "원본과 같은 폴더"


# --------------------------------------------------------------------------- 화면과 무관한 도우미

def is_word_file(path: Path) -> bool:
    """변환 대상 .docx 인가. 워드가 열어 둔 임시(잠금) 파일 `~$...` 은 제외."""
    name = path.name
    return path.suffix.lower() == ".docx" and not name.startswith("~$")


def expand_inputs(paths: Iterable[str | os.PathLike]) -> list[Path]:
    """끌어다 놓은/고른 경로를 .docx 파일 목록으로 펼친다.

    폴더는 안쪽(하위 폴더 포함)의 .docx 를 이름순으로 가져온다. `~$` 잠금 파일과
    .docx 가 아닌 파일은 건너뛰고, 같은 파일이 두 번 나오지 않게 한다.
    """
    seen: set[str] = set()
    result: list[Path] = []

    def add(p: Path) -> None:
        key = os.path.normcase(os.path.abspath(p))
        if key not in seen and is_word_file(p):
            seen.add(key)
            result.append(Path(os.path.abspath(p)))

    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            found: list[Path] = []
            for root, dirs, files in os.walk(p):
                dirs.sort()
                found.extend(Path(root) / f for f in sorted(files))
            for f in found:
                add(f)
        elif p.is_file():
            add(p)
    return result


def unique_output_path(source: Path, out_dir: Path | None = None,
                       reserved: Iterable[Path] = ()) -> Path:
    """`이름.hwpx`, 이미 있으면 `이름 (1).hwpx`, `이름 (2).hwpx` … 로 비어 있는 경로를 돌려준다.

    out_dir 이 None 이면 원본과 같은 폴더. reserved 는 이번 작업에서 이미 쓰기로 한 경로.
    """
    source = Path(source)
    folder = Path(out_dir) if out_dir else source.parent
    taken = {os.path.normcase(os.path.abspath(r)) for r in reserved}
    stem = source.stem

    def free(p: Path) -> bool:
        return not p.exists() and os.path.normcase(os.path.abspath(p)) not in taken

    cand = folder / f"{stem}.hwpx"
    n = 0
    while not free(cand):
        n += 1
        cand = folder / f"{stem} ({n}).hwpx"
    return cand


def status_for(result: ConversionResult) -> str:
    if not result.ok:
        return ST_FAIL
    return ST_PART if result.losses else ST_OK


def detail_text(result: ConversionResult | None, status: str) -> str:
    """오른쪽 아래 '자세한 내용' 칸에 보여 줄 글."""
    if result is None:
        if status == ST_WAIT:
            return "변환을 기다리는 파일입니다. [변환] 단추를 누르세요."
        if status == ST_RUN:
            return "변환하고 있습니다…"
        return ""
    if not result.ok:
        return "변환하지 못했습니다.\n" + (result.error or "알 수 없는 오류입니다.")
    lines = [f"저장했습니다: {result.output}"]
    if result.losses:
        lines.append("")
        lines.append("한글로 옮기지 못했거나 달라진 부분이 있습니다. 한글에서 열어 확인해 보세요.")
        for loss in result.losses:
            mark = "빠짐" if loss.kind == "dropped" else "달라짐"
            lines.append(f"  - [{mark}] {loss.label}")
    else:
        lines.append("")
        lines.append("빠지거나 달라진 부분 없이 옮겼습니다.")
    return "\n".join(lines)


def enable_dpi_awareness() -> None:
    """윈도우 배율(125~150%)에서도 글자가 흐려지지 않게 한다."""
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # 모니터별 배율
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _make_root() -> tk.Tk:
    try:
        from tkinterdnd2 import TkinterDnD
        return TkinterDnD.Tk()
    except Exception:
        return tk.Tk()


def resource_path(rel: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


# --------------------------------------------------------------------------- 창

@dataclass
class Item:
    source: Path
    iid: str = ""
    status: str = ST_WAIT
    result: ConversionResult | None = None
    output: Path | None = None
    extra: dict = field(default_factory=dict)


class App:
    def __init__(self, root: tk.Tk | None = None) -> None:
        enable_dpi_awareness()
        self.root = root or _make_root()
        self.root.title(APP_TITLE)
        self.items: dict[str, Item] = {}
        self.order: list[str] = []
        self.msgs: "queue.Queue[tuple]" = queue.Queue()
        self.worker: threading.Thread | None = None
        self.running = False
        self.dnd_ok = False
        self.out_dir: Path | None = None

        self._setup_style()
        self._build()
        self._set_geometry()
        self._enable_drop()
        self._bind_keys()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll)
        self._refresh_buttons()

    # ---- 겉모양
    def _setup_style(self) -> None:
        import tkinter.font as tkfont

        families = set(tkfont.families(self.root))
        family = FONT_FAMILY if FONT_FAMILY in families else "Segoe UI"
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkFixedFont"):
            try:
                tkfont.nametofont(name).configure(family=family, size=10)
            except tk.TclError:
                pass
        self.font = tkfont.Font(family=family, size=10)
        self.font_bold = tkfont.Font(family=family, size=10, weight="bold")
        self.font_big = tkfont.Font(family=family, size=12, weight="bold")
        self.font_hint = tkfont.Font(family=family, size=13)
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        line = self.font.metrics("linespace")
        style.configure("Treeview", font=self.font, rowheight=line + 12)
        style.configure("Treeview.Heading", font=self.font_bold)
        style.configure("TButton", font=self.font, padding=(10, 5))
        style.configure("Big.TButton", font=self.font_big, padding=(22, 9))
        style.configure("TLabel", font=self.font)
        style.configure("TRadiobutton", font=self.font)
        style.configure("Hint.TLabel", font=self.font, foreground="#555555")
        style.configure("Status.TLabel", font=self.font)
        icon = resource_path("app.ico")
        if not icon.exists():
            icon = Path(__file__).resolve().parents[2] / "packaging" / "app.ico"
        if icon.exists() and sys.platform == "win32":
            try:
                self.root.iconbitmap(default=str(icon))
            except tk.TclError:
                pass

    def _set_geometry(self) -> None:
        self.root.minsize(760, 560)
        self.root.update_idletasks()
        w = max(self.root.winfo_reqwidth(), 900)
        h = max(self.root.winfo_reqheight(), 660)
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = min(w, sw - 80), min(h, sh - 120)
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{max((sh - h) // 3, 0)}")

    def _build(self) -> None:
        pad = 12
        outer = ttk.Frame(self.root, padding=pad)
        outer.pack(fill="both", expand=True)
        self.outer = outer

        # 위: 파일 추가 / 지우기
        top = ttk.Frame(outer)
        top.pack(fill="x")
        self.btn_add = ttk.Button(top, text="파일 추가", command=self.add_files)
        self.btn_add.pack(side="left")
        self.btn_add_dir = ttk.Button(top, text="폴더 추가", command=self.add_folder)
        self.btn_add_dir.pack(side="left", padx=(6, 0))
        self.btn_clear = ttk.Button(top, text="목록 비우기", command=self.clear_all)
        self.btn_clear.pack(side="right")
        self.btn_remove = ttk.Button(top, text="선택 지우기", command=self.remove_selected)
        self.btn_remove.pack(side="right", padx=(0, 6))

        # 목록
        mid = ttk.Frame(outer)
        mid.pack(fill="both", expand=True, pady=(10, 0))
        cols = ("file", "status", "where")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="extended", height=8)
        self.tree.heading("file", text="워드 파일")
        self.tree.heading("status", text="상태")
        self.tree.heading("where", text="저장된 한글 파일(다른 폴더면 전체 경로)")
        self.tree.column("file", width=330, anchor="w")
        self.tree.column("status", width=90, anchor="center", stretch=False)
        self.tree.column("where", width=330, anchor="w")
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("ok", foreground="#1a6b2f")
        self.tree.tag_configure("part", foreground="#9a5b00")
        self.tree.tag_configure("fail", foreground="#b3261e")
        self.tree.tag_configure("run", foreground="#1f5fbf")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_select())
        self.tree.bind("<Double-1>", lambda e: self.open_hwpx())
        self.hint = tk.Label(mid, text="여기에 워드 파일(.docx)이나 폴더를\n끌어다 놓거나, [파일 추가]를 누르세요.",
                             font=self.font_hint, fg="#8a8a8a", justify="center")
        self.hint.place(relx=0.5, rely=0.5, anchor="center")

        # 저장 위치
        dest = ttk.LabelFrame(outer, text="저장 위치", padding=(10, 6))
        dest.pack(fill="x", pady=(10, 0))
        self.dest_mode = tk.StringVar(value="same")
        ttk.Radiobutton(dest, text=SAME_FOLDER_LABEL, value="same", variable=self.dest_mode,
                        command=self._on_dest_mode).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(dest, text="다른 폴더:", value="other", variable=self.dest_mode,
                        command=self._on_dest_mode).grid(row=0, column=1, sticky="w", padx=(18, 0))
        self.dest_var = tk.StringVar(value="")
        self.dest_entry = ttk.Entry(dest, textvariable=self.dest_var, state="readonly")
        self.dest_entry.grid(row=0, column=2, sticky="ew", padx=(6, 6))
        self.btn_browse = ttk.Button(dest, text="찾아보기", command=self.choose_folder)
        self.btn_browse.grid(row=0, column=3)
        dest.columnconfigure(2, weight=1)
        ttk.Label(dest, style="Hint.TLabel",
                  text="이름은 같고 .hwpx 로 저장됩니다. 같은 이름이 이미 있으면 '이름 (1).hwpx' 처럼 새 이름을 씁니다. 기존 파일은 덮어쓰지 않습니다."
                  ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(4, 0))

        # 자세한 내용
        det = ttk.LabelFrame(outer, text="자세한 내용", padding=(8, 6))
        det.pack(fill="both", pady=(10, 0))
        self.detail = tk.Text(det, height=7, wrap="word", font=self.font, relief="flat",
                              background="#fafafa", padx=6, pady=4, state="disabled")
        dsb = ttk.Scrollbar(det, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.pack(side="left", fill="both", expand=True)
        dsb.pack(side="right", fill="y")

        # 아래: 진행, 단추
        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(10, 0))
        self.progress = ttk.Progressbar(bottom, maximum=100, mode="determinate")
        self.progress.pack(fill="x")
        self.status_var = tk.StringVar(value="파일을 추가하고 [변환]을 누르세요.")
        ttk.Label(bottom, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w", pady=(4, 0))

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(10, 0))
        self.btn_convert = ttk.Button(btns, text="변환", style="Big.TButton", command=self.start)
        self.btn_convert.pack(side="left")
        self.btn_open_folder = ttk.Button(btns, text="폴더 열기", command=self.open_folder)
        self.btn_open_folder.pack(side="left", padx=(14, 0))
        self.btn_open_hwpx = ttk.Button(btns, text="한/글로 열기", command=self.open_hwpx)
        self.btn_open_hwpx.pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="정보", command=self.show_about).pack(side="right")

    # ---- 끌어다 놓기, 키보드
    def _enable_drop(self) -> None:
        try:
            from tkinterdnd2 import DND_FILES
        except Exception:
            return
        if not hasattr(self.root, "drop_target_register"):
            return

        def reg(w: tk.Misc) -> None:
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                return
            for c in w.winfo_children():
                reg(c)

        reg(self.root)
        self.dnd_ok = True

    def _on_drop(self, event) -> str:
        try:
            paths = self.root.tk.splitlist(event.data)
        except Exception:
            paths = [event.data]
        self.add_paths(paths)
        return "copy"

    def _bind_keys(self) -> None:
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.root.bind("<Delete>", lambda e: self.remove_selected() if self.root.focus_get() is self.tree else None)
        self.root.bind("<Control-o>", lambda e: self.add_files())

    # ---- 목록 관리
    def add_files(self) -> None:
        files = filedialog.askopenfilenames(
            parent=self.root, title="변환할 워드 파일 고르기",
            filetypes=[("워드 문서", "*.docx"), ("모든 파일", "*.*")])
        if files:
            self.add_paths(self.root.tk.splitlist(files) if isinstance(files, str) else files)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(parent=self.root, title="워드 파일이 들어 있는 폴더 고르기")
        if folder:
            self.add_paths([folder])

    def add_paths(self, paths: Iterable[str]) -> int:
        found = expand_inputs(paths)
        have = {os.path.normcase(str(i.source)) for i in self.items.values()}
        added = 0
        for p in found:
            if os.path.normcase(str(p)) in have:
                continue
            self._add_item(p)
            added += 1
        if added:
            self.status_var.set(f"{added}개를 추가했습니다. [변환]을 누르세요.")
        elif not found:
            self.status_var.set("추가할 .docx 파일이 없습니다. (워드 파일이나 그 파일이 든 폴더를 놓아 주세요)")
        else:
            self.status_var.set("이미 목록에 있는 파일입니다.")
        self._refresh_buttons()
        return added

    def _add_item(self, path: Path) -> Item:
        item = Item(source=path)
        iid = self.tree.insert("", "end", values=(path.name, ST_WAIT, ""))
        item.iid = iid
        self.items[iid] = item
        self.order.append(iid)
        return item

    def remove_selected(self) -> None:
        if self.running:
            return
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.items.pop(iid, None)
            if iid in self.order:
                self.order.remove(iid)
        self._show_detail(None)
        self._refresh_buttons()

    def clear_all(self) -> None:
        if self.running:
            return
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        self.order.clear()
        self.progress["value"] = 0
        self.status_var.set("목록을 비웠습니다.")
        self._show_detail(None)
        self._refresh_buttons()

    # ---- 저장 위치
    def _on_dest_mode(self) -> None:
        if self.dest_mode.get() == "other" and self.out_dir is None:
            self.choose_folder()
        self._refresh_buttons()

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(parent=self.root, title="한글 파일을 저장할 폴더 고르기")
        if folder:
            self.out_dir = Path(folder)
            self._set_entry(str(self.out_dir))
            self.dest_mode.set("other")
        elif self.out_dir is None:
            self.dest_mode.set("same")
        self._refresh_buttons()

    def _set_entry(self, text: str) -> None:
        self.dest_entry.configure(state="normal")
        self.dest_var.set(text)
        self.dest_entry.configure(state="readonly")

    def _target_dir(self) -> Path | None:
        if self.dest_mode.get() == "other" and self.out_dir is not None:
            return self.out_dir
        return None

    # ---- 변환
    def start(self) -> None:
        if self.running:
            return
        todo = [self.items[i] for i in self.order if self.items[i].status in (ST_WAIT, ST_FAIL)]
        if not todo:
            messagebox.showinfo(APP_TITLE, "변환할 파일이 없습니다.\n[파일 추가]로 워드 파일을 넣어 주세요.", parent=self.root)
            return
        target = self._target_dir()
        if target is not None:
            try:
                target.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                messagebox.showerror(APP_TITLE, f"저장 폴더를 만들 수 없습니다.\n{e}", parent=self.root)
                return
        for it in todo:
            self._set_status(it, ST_WAIT)
        self.running = True
        self.progress["value"] = 0
        self.status_var.set("변환을 시작합니다…")
        self._refresh_buttons()
        self.worker = threading.Thread(target=self._work, args=(todo, target), daemon=True)
        self.worker.start()

    def _work(self, todo: list[Item], target: Path | None) -> None:
        total = len(todo)
        reserved: list[Path] = []
        for n, it in enumerate(todo):
            self.msgs.put(("start", it, n, total))
            out = unique_output_path(it.source, target, reserved)
            reserved.append(out)

            def cb(frac: float, text: str, n=n, it=it) -> None:
                try:
                    f = min(max(float(frac), 0.0), 1.0)
                except (TypeError, ValueError):
                    f = 0.0
                self.msgs.put(("progress", it, (n + f) / total, text))

            try:
                res = convert(it.source, out, overwrite=False, progress=cb)
            except Exception as e:  # 변환 쪽이 약속을 어겨도 창은 계속 동작해야 한다
                res = ConversionResult(source=it.source, output=out, ok=False,
                                       error=f"예상하지 못한 오류가 났습니다: {e}")
            self.msgs.put(("done", it, res, (n + 1) / total))
        self.msgs.put(("finish", total))

    def _poll(self) -> None:
        try:
            while True:
                m = self.msgs.get_nowait()
                self._handle(m)
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _handle(self, m: tuple) -> None:
        kind = m[0]
        if kind == "start":
            _, it, n, total = m
            if it.iid not in self.items:
                return
            self._set_status(it, ST_RUN)
            self.status_var.set(f"({n + 1}/{total}) {it.source.name} 변환 중…")
            self.progress["value"] = 100.0 * n / total
        elif kind == "progress":
            _, it, frac, text = m
            self.progress["value"] = 100.0 * frac
            self.status_var.set(f"{it.source.name}: {text}" if text else f"{it.source.name} 변환 중…")
        elif kind == "done":
            _, it, res, frac = m
            self.progress["value"] = 100.0 * frac
            it.result = res
            it.output = res.output if res.ok else None
            self._set_status(it, status_for(res))
            if it.iid in self.tree.selection():
                self._show_detail(it)
        elif kind == "finish":
            total = m[1]
            self.running = False
            oks = sum(1 for i in self.items.values() if i.status in (ST_OK, ST_PART))
            fails = sum(1 for i in self.items.values() if i.status == ST_FAIL)
            msg = f"변환을 마쳤습니다. 성공 {oks}개"
            if fails:
                msg += f", 실패 {fails}개"
            self.status_var.set(msg)
            self.progress["value"] = 100
            self._refresh_buttons()
            # 마지막으로 처리한 파일의 자세한 내용을 바로 보여 준다
            if not self.tree.selection():
                last = next((self.items[i] for i in reversed(self.order)
                             if self.items[i].result is not None), None)
                if last is not None:
                    self.tree.selection_set(last.iid)
                    self.tree.see(last.iid)

    def _set_status(self, it: Item, status: str) -> None:
        it.status = status
        if it.iid not in self.items:
            return
        where = ""
        if it.output and status in (ST_OK, ST_PART):
            where = it.output.name if it.output.parent == it.source.parent else str(it.output)
        if status == ST_WAIT:
            it.result = None
            it.output = None
            where = ""
        tag = {ST_OK: "ok", ST_PART: "part", ST_FAIL: "fail", ST_RUN: "run"}.get(status, "")
        self.tree.item(it.iid, values=(it.source.name, status, where), tags=(tag,) if tag else ())
        self._refresh_buttons()

    # ---- 자세한 내용, 단추 상태
    def _selected_item(self) -> Item | None:
        sel = self.tree.selection()
        return self.items.get(sel[0]) if sel else None

    def _on_select(self) -> None:
        self._show_detail(self._selected_item())
        self._refresh_buttons()

    def _show_detail(self, it: Item | None) -> None:
        text = "" if it is None else detail_text(it.result, it.status)
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", text)
        self.detail.configure(state="disabled")

    def _refresh_buttons(self) -> None:
        run = self.running
        has_items = bool(self.items)
        self.hint.place_forget() if has_items else self.hint.place(relx=0.5, rely=0.5, anchor="center")
        waiting = any(i.status in (ST_WAIT, ST_FAIL) for i in self.items.values())
        self.btn_convert.state(["!disabled"] if (waiting and not run) else ["disabled"])
        for b in (self.btn_remove, self.btn_clear):
            b.state(["disabled"] if (run or not has_items) else ["!disabled"])
        self.btn_remove.state(["disabled"] if (run or not self.tree.selection()) else ["!disabled"])
        self.btn_browse.state(["disabled"] if run else ["!disabled"])
        it = self._opening_item()
        can = it is not None and it.output is not None and Path(it.output).exists()
        self.btn_open_folder.state(["!disabled"] if can else ["disabled"])
        self.btn_open_hwpx.state(["!disabled"] if can else ["disabled"])

    def _opening_item(self) -> Item | None:
        """열기 단추가 가리킬 파일: 선택한 것, 없으면 마지막으로 변환에 성공한 것."""
        it = self._selected_item()
        if it is not None and it.output is not None:
            return it
        if it is not None:
            return None
        for iid in reversed(self.order):
            cand = self.items[iid]
            if cand.output is not None:
                return cand
        return None

    # ---- 열기
    def open_folder(self) -> None:
        it = self._opening_item()
        if it is None or it.output is None:
            return
        self._startfile(Path(it.output).parent)

    def open_hwpx(self) -> None:
        it = self._opening_item()
        if it is None or it.output is None:
            return
        self._startfile(Path(it.output))

    def _startfile(self, path: Path) -> None:
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except AttributeError:
            messagebox.showinfo(APP_TITLE, f"윈도우에서만 열 수 있습니다.\n{path}", parent=self.root)
        except OSError as e:
            messagebox.showerror(APP_TITLE, f"열지 못했습니다.\n{path}\n\n{e}\n\n"
                                 ".hwpx 를 여는 프로그램(한/글)이 설치되어 있는지 확인해 주세요.", parent=self.root)

    # ---- 정보, 닫기
    def show_about(self) -> None:
        try:
            import hwpx
            engine = getattr(hwpx, "__version__", "")
        except Exception:
            engine = ""
        win = tk.Toplevel(self.root)
        win.title("정보")
        win.transient(self.root)
        win.resizable(False, False)
        f = ttk.Frame(win, padding=22)
        f.pack()
        tk.Label(f, text=APP_TITLE, font=self.font_big).pack(anchor="w")
        ttk.Label(f, text=f"버전 {__version__}").pack(anchor="w", pady=(4, 0))
        ttk.Label(f, text="python-hwpx 기반" + (f" (엔진 {engine})" if engine else "")).pack(anchor="w", pady=(2, 0))
        ttk.Label(f, style="Hint.TLabel", justify="left", wraplength=360,
                  text="워드(.docx) 문서를 한글(.hwpx) 문서로 바꿉니다.\n"
                       "원본 파일은 건드리지 않고, 기존 파일을 덮어쓰지 않습니다.").pack(anchor="w", pady=(12, 0))
        ttk.Label(f, style="Hint.TLabel", justify="left", wraplength=360,
                  text="[사용권 안내 — 추후 기재]").pack(anchor="w", pady=(12, 0))
        ttk.Button(f, text="닫기", command=win.destroy).pack(anchor="e", pady=(16, 0))
        win.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_reqwidth()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - win.winfo_reqheight()) // 3
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        win.bind("<Escape>", lambda e: win.destroy())
        win.grab_set()
        win.focus_set()

    def _on_close(self) -> None:
        if self.running and not messagebox.askokcancel(
                APP_TITLE, "변환하는 중입니다. 지금 끝내면 작업 중인 파일이 저장되지 않을 수 있습니다.\n그래도 끝낼까요?",
                parent=self.root):
            return
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    App().run()


if __name__ == "__main__":
    main()
