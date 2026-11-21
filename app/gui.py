"""
SellerMorning — 화면 프로그램(Windows 10 / 11 · macOS)

실행: 프로젝트 폴더의 start_gui.bat 더블클릭(또는 pythonw app\\gui.py)
  - 왼쪽: 기능 목록(체크하면 여러 개를 차례로 실행), 선택한 기능의 설명·옵션
  - 오른쪽 위: 실행 로그(실시간), 오른쪽 아래: 만들어진 결과 파일(열기 / 폴더에서 보기)
기능은 별도 프로세스(runner.py)로 실행되어 화면이 멈추지 않고, ‘중지’로 언제든 멈출 수 있습니다.
"""

import xlsafe  # noqa: F401  Excel 수식 주입 방지(Amazon에서 받은 글자가 수식으로 실행되지 않게)
import os
import queue
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import tasks

BASE = Path(__file__).resolve().parent            # 프로그램 파일(app 폴더)
ROOT = BASE.parent                                  # 프로젝트 폴더(bat·reports)
REPORTS = ROOT / "reports"
LOG_DIR = REPORTS / "logs"
IS_WIN = sys.platform.startswith("win")
IS_MAC = sys.platform == "darwin"
SKIP_DIRS = {"raw", "debug", "data", "pages", "detail", "snapshots", "brand_analytics", "logs", "history", "rawt"}
FOLDER_NAME = {"account_health": "계정 상태", "inventory": "재고", "orders": "주문", "business": "비즈니스", "ads": "광고",
               "explore": "광고 탐색", "listing": "Listing", "daily": "오늘의 요약", "scan": "메뉴 스캔", "changes": "수정", "funnel": "퍼널", "buybox": "Buy Box", "finance": "재무·보상"}
FONT = "Malgun Gothic" if IS_WIN else ("Apple SD Gothic Neo" if IS_MAC else "Noto Sans CJK KR")
MONO = FONT                          # 로그도 한글이 잘 보이는 일반 글꼴
ICON_ICO = BASE / "assets" / "icon.ico"
ICON_PNG = BASE / "assets" / "icon.png"
# 색(짙은 남색 + 주황)
NAVY, NAVY2, ORANGE, ORANGE2 = "#232F3E", "#37475A", "#FF9900", "#FFAD33"
BG, CARD, LINE, TEXT, MUTED = "#F3F4F6", "#FFFFFF", "#D9DDE3", "#1F2937", "#6B7280"
SEL_BG = "#FFF1DB"


import re
_EMOJI_RE = re.compile("[\U00010000-\U0010FFFF]")
_EMOJI_MAP = {"📁": "▶", "📊": "■", "📋": "■", "📒": "■", "📈": "■", "📉": "■", "🔍": "•"}


def python_exe() -> str:
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe" and (exe.parent / "python.exe").exists():
        return str(exe.parent / "python.exe")
    return str(exe)


def open_path(p) -> None:
    import platform_util as pu
    if not pu.open_path(p):
        messagebox.showerror("열기 실패", f"{p}")


def show_in_folder(p: Path) -> None:
    import platform_util as pu
    pu.show_in_folder(p)


def kill_tree(proc) -> None:
    """실행 중인 기능(브라우저 포함) 멈추기."""
    try:
        if IS_WIN:
            subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True, creationflags=0x08000000)
        else:
            import signal
            try:
                os.killpg(os.getpgid(proc.pid), signal.SIGTERM)     # 같은 그룹(브라우저 등)까지
            except Exception:
                proc.terminate()
    except Exception:
        pass


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.proc = None
        self.q: "queue.Queue[tuple]" = queue.Queue()
        self.run_queue = []
        self.total = 0
        self.current = None
        self.started = 0.0
        self.last_output = None
        self.stopped = False
        self.checks = {}
        self.selected = None
        root.title(f"SellerMorning {__import__('version').text()}")
        root.geometry("1360x860")
        root.minsize(1000, 640)
        self._style()
        self._build()
        self.select(tasks.TASKS[3].key)
        self.refresh_results()
        self._refresh_api_state()
        try:
            import daily
            daily.fix_bat()
        except Exception:
            pass
        root.after(120, self._poll)
        root.after(1000, self._tick)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------ 화면 구성
    def _style(self):
        import tkinter.font as tkfont
        for n in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(n).configure(family=FONT, size=10)
            except Exception:
                pass
        self.root.configure(bg=BG)
        st = ttk.Style()
        try:
            st.theme_use("clam")          # 색을 자유롭게 바꿀 수 있는 테마
        except Exception:
            pass
        st.configure(".", background=BG, foreground=TEXT, font=(FONT, 10), bordercolor=LINE, lightcolor=BG,
                     darkcolor=BG, troughcolor="#E5E7EB", focuscolor=ORANGE)
        st.configure("TFrame", background=BG)
        st.configure("TLabel", background=BG, foreground=TEXT)
        st.configure("TLabelframe", background=BG, bordercolor=LINE, relief="solid")
        st.configure("TLabelframe.Label", background=BG, foreground=NAVY, font=(FONT, 10, "bold"))
        st.configure("TCheckbutton", background=BG)
        st.configure("TRadiobutton", background=BG)
        st.map("TCheckbutton", background=[("active", BG)], indicatorcolor=[("selected", ORANGE), ("!selected", CARD)])
        st.map("TRadiobutton", background=[("active", BG)], indicatorcolor=[("selected", ORANGE), ("!selected", CARD)])
        st.configure("TButton", background=CARD, foreground=TEXT, bordercolor=LINE, padding=(10, 5), relief="flat")
        st.map("TButton", background=[("disabled", "#ECEEF1"), ("pressed", "#E5E7EB"), ("active", "#F7F8FA")],
               foreground=[("disabled", "#A0A6AF")], bordercolor=[("active", ORANGE)])
        st.configure("Run.TButton", background=ORANGE, foreground="#111111", bordercolor=ORANGE, font=(FONT, 10, "bold"),
                     padding=(12, 7))
        st.map("Run.TButton", background=[("disabled", "#F3D9B1"), ("pressed", "#E68A00"), ("active", ORANGE2)],
               bordercolor=[("active", ORANGE2)], foreground=[("disabled", "#8A7A63")])
        st.configure("Stop.TButton", foreground="#B42318", padding=(12, 7))
        st.map("Stop.TButton", foreground=[("disabled", "#C9A9A6")], bordercolor=[("active", "#B42318")])
        st.configure("Head.TButton", background=NAVY2, foreground="#FFFFFF", bordercolor=NAVY2, padding=(10, 5))
        st.map("Head.TButton", background=[("pressed", "#2B3848"), ("active", "#485A70")], bordercolor=[("active", ORANGE)])
        st.configure("Group.TLabel", font=(FONT, 10, "bold"), foreground=ORANGE, background=CARD)
        st.configure("Title.TLabel", font=(FONT, 14, "bold"), foreground=NAVY)
        st.configure("Name.TLabel", font=(FONT, 12, "bold"), foreground=NAVY)
        st.configure("Muted.TLabel", foreground=MUTED)
        st.configure("Card.TFrame", background=CARD)
        st.configure("Card.TCheckbutton", background=CARD)
        st.map("Card.TCheckbutton", background=[("active", CARD)], indicatorcolor=[("selected", ORANGE), ("!selected", CARD)])
        st.configure("Treeview", rowheight=26, background=CARD, fieldbackground=CARD, bordercolor=LINE)
        st.map("Treeview", background=[("selected", "#FFE2B3")], foreground=[("selected", TEXT)])
        st.configure("Treeview.Heading", background="#EEF1F5", foreground=NAVY, font=(FONT, 10, "bold"), relief="flat",
                     bordercolor=LINE)
        st.map("Treeview.Heading", background=[("active", "#E3E8EF")])
        st.configure("TNotebook", background=BG, bordercolor=LINE, tabmargins=(0, 4, 0, 0))
        st.configure("TNotebook.Tab", background="#E5E7EB", foreground=MUTED, padding=(14, 6), bordercolor=LINE)
        st.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", NAVY)],
               lightcolor=[("selected", ORANGE)], expand=[("selected", (0, 2, 0, 0))])
        st.configure("Horizontal.TProgressbar", background=ORANGE, troughcolor="#E5E7EB", bordercolor=LINE,
                     lightcolor=ORANGE, darkcolor=ORANGE)
        st.configure("TEntry", fieldbackground=CARD, bordercolor=LINE, padding=4)
        st.map("TEntry", bordercolor=[("focus", ORANGE)], lightcolor=[("focus", ORANGE)])
        st.configure("TCombobox", fieldbackground=CARD, bordercolor=LINE)
        st.configure("TPanedwindow", background=BG)
        st.configure("Vertical.TScrollbar", background="#E5E7EB", troughcolor=BG, bordercolor=BG, arrowcolor=MUTED)

    def _set_icon(self, win):
        try:
            if IS_WIN and ICON_ICO.exists():
                win.iconbitmap(default=str(ICON_ICO))
            elif ICON_PNG.exists():
                self._icon_img = tk.PhotoImage(file=str(ICON_PNG))
                win.iconphoto(True, self._icon_img)
        except Exception:
            pass

    def _build(self):
        self._set_icon(self.root)
        top = tk.Frame(self.root, bg=NAVY, padx=14, pady=10)
        top.pack(fill="x")
        try:
            self._logo = tk.PhotoImage(file=str(ICON_PNG)).subsample(6, 6)      # 256 → 약 43px
            tk.Label(top, image=self._logo, bg=NAVY).pack(side="left", padx=(0, 10))
        except Exception:
            pass
        tt = tk.Frame(top, bg=NAVY)
        tt.pack(side="left")
        hl = tk.Frame(tt, bg=NAVY)
        hl.pack(anchor="w")
        tk.Label(hl, text="SellerMorning", bg=NAVY, fg="#FFFFFF", font=(FONT, 15, "bold")).pack(side="left")
        tk.Label(hl, text="v" + __import__("version").VERSION, bg=NAVY, fg="#AEB8C4", font=(FONT, 9)).pack(side="left", padx=(8, 0), pady=(5, 0))
        tk.Label(tt, text="조회·제안은 자동 · 수정은 확인한 것만 적용",
                 bg=NAVY, fg="#AEB8C4", font=(FONT, 9)).pack(anchor="w")
        ttk.Button(top, text="바탕화면 바로가기", style="Head.TButton", command=self.make_shortcut).pack(side="right")
        ttk.Button(top, text="⚙ 설정(API·메일)", style="Head.TButton", command=self.open_settings).pack(side="right", padx=6)
        ttk.Button(top, text="reports 폴더", style="Head.TButton", command=lambda: open_path(REPORTS)).pack(side="right")
        self.api_state = tk.StringVar(value="")
        self.api_lbl = tk.Label(top, textvariable=self.api_state, bg=NAVY, fg="#FFC266", font=(FONT, 9))
        self.api_lbl.pack(side="right", padx=12)
        tk.Frame(self.root, bg=ORANGE, height=3).pack(fill="x")
        self.status = tk.StringVar(value="대기 중")
        bar = tk.Frame(self.root, bg=CARD, padx=14, pady=7, highlightthickness=1, highlightbackground=LINE)
        bar.pack(fill="x", padx=10, pady=(8, 8))
        self.status_dot = tk.Label(bar, text="●", bg=CARD, fg="#1E7B34", font=(FONT, 11))
        self.status_dot.pack(side="left")
        self.status_lbl = tk.Label(bar, textvariable=self.status, bg=CARD, fg=TEXT, font=(FONT, 10, "bold"))
        self.status_lbl.pack(side="left", padx=(4, 0))
        self.step = tk.StringVar(value="")
        tk.Label(bar, textvariable=self.step, bg=CARD, fg=MUTED, font=(FONT, 9)).pack(side="left", padx=12)
        self.qinfo = tk.StringVar(value="")
        tk.Label(bar, textvariable=self.qinfo, bg=CARD, fg=MUTED, font=(FONT, 9)).pack(side="right", padx=(8, 0))
        self.pbar = ttk.Progressbar(bar, mode="indeterminate", length=260)
        self.pbar.pack(side="right")
        self.busy_widgets = []       # 실행 중에는 누를 수 없는 버튼·체크 상자

        pw = ttk.PanedWindow(self.root, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        left = ttk.Frame(pw, padding=(0, 0, 8, 0))
        right = ttk.Frame(pw)
        pw.add(left, weight=1)
        pw.add(right, weight=3)

        # ---- 기능 목록
        lf = ttk.LabelFrame(left, text="기능(체크하면 여러 개를 차례로 실행)", padding=6)
        lf.pack(fill="both", expand=True)
        canvas = tk.Canvas(lf, highlightthickness=0, width=360, bg=CARD)
        sb = ttk.Scrollbar(lf, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="Card.TFrame")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120) if IS_WIN else -int(e.delta), "units")
                        if self._over(canvas) else None)
        self.name_lbls = {}
        group = None
        for t in tasks.TASKS:
            if t.group != group:
                group = t.group
                ttk.Label(inner, text=group, style="Group.TLabel").pack(anchor="w", pady=(10, 2), padx=4)
            row = ttk.Frame(inner, style="Card.TFrame")
            row.pack(fill="x", anchor="w")
            v = tk.BooleanVar(value=False)
            self.checks[t.key] = v
            cb = ttk.Checkbutton(row, variable=v, style="Card.TCheckbutton")
            cb.pack(side="left")
            if t.ready:
                self.busy_widgets.append(cb)
            lbl = tk.Label(row, text=t.name + ("" if t.ready else "  (개발 중)"), anchor="w", cursor="hand2",
                           font=(FONT, 10), fg=TEXT if t.ready else "#A0A6AF", bg=CARD, padx=6, pady=3)
            lbl.pack(side="left", fill="x", expand=True)
            lbl.bind("<Button-1>", lambda e, k=t.key: self.select(k))
            lbl.bind("<Double-Button-1>", lambda e, k=t.key: (self.select(k), self.run_selected()))
            if not t.ready:
                cb.state(["disabled"])
            self.name_lbls[t.key] = lbl

        # ---- 선택한 기능
        df = ttk.LabelFrame(left, text="선택한 기능", padding=8)
        df.pack(fill="x", pady=(8, 0))
        self.d_name = ttk.Label(df, text="", style="Name.TLabel")
        self.d_name.pack(anchor="w")
        self.d_desc = ttk.Label(df, text="", wraplength=380, justify="left")
        self.d_desc.pack(anchor="w", pady=(4, 2))
        self.d_status = ttk.Label(df, text="", style="Muted.TLabel", wraplength=380, justify="left")
        self.d_status.pack(anchor="w")
        self.opt = ttk.Frame(df)
        self.opt.pack(fill="x", pady=(6, 0))
        self.period = tk.StringVar(value="2")

        bf = ttk.Frame(left)
        bf.pack(fill="x", pady=8)
        self.btn_run = ttk.Button(bf, text="▶ 이 기능 실행", style="Run.TButton", command=self.run_selected)
        self.btn_run.pack(side="left", fill="x", expand=True)
        self.btn_queue = ttk.Button(bf, text="▶ 체크한 기능 차례로", command=self.run_checked)
        self.btn_queue.pack(side="left", fill="x", expand=True, padx=4)
        self.btn_stop = ttk.Button(bf, text="■ 중지", style="Stop.TButton", command=self.stop)
        self.btn_stop.pack(side="left")
        self.btn_stop.state(["disabled"])
        self.auto_open = tk.BooleanVar(value=True)
        cb_open = ttk.Checkbutton(left, text="끝나면 결과 파일 자동으로 열기", variable=self.auto_open)
        cb_open.pack(anchor="w")
        self.busy_widgets.append(cb_open)
        self.opt_widgets = []        # 선택한 기능의 옵션(기간·예약 등)

        # ---- 로그
        rpw = ttk.PanedWindow(right, orient="vertical")
        rpw.pack(fill="both", expand=True)
        self.nb = ttk.Notebook(rpw)
        lg = ttk.Frame(self.nb, padding=4)
        sm = ttk.Frame(self.nb, padding=4)
        self.nb.add(lg, text="  실행 로그  ")
        self.nb.add(sm, text="  오늘의 요약  ")
        rs = ttk.LabelFrame(rpw, text="결과 파일(더블클릭하면 열림)", padding=4)
        rpw.add(self.nb, weight=3)
        rpw.add(rs, weight=2)
        self._build_summary(sm)
        tb = ttk.Frame(lg)
        tb.pack(fill="x")
        ttk.Button(tb, text="로그 지우기", command=lambda: self.log.delete("1.0", "end")).pack(side="left")
        ttk.Button(tb, text="로그 폴더 열기", command=lambda: (LOG_DIR.mkdir(parents=True, exist_ok=True), open_path(LOG_DIR))).pack(side="left", padx=4)
        self.autoscroll = tk.BooleanVar(value=True)
        ttk.Checkbutton(tb, text="자동 스크롤", variable=self.autoscroll).pack(side="left", padx=8)
        self.btn_last = ttk.Button(tb, text="방금 만든 파일 열기", command=lambda: self.last_output and open_path(self.last_output))
        self.btn_last.pack(side="right")
        self.btn_last_dir = ttk.Button(tb, text="방금 만든 파일 위치", command=lambda: self.last_output and show_in_folder(self.last_output))
        self.btn_last_dir.pack(side="right", padx=4)
        for b in (self.btn_last, self.btn_last_dir):
            b.state(["disabled"])
        self.log = ScrolledText(lg, wrap="word", font=(MONO, 10), bg=CARD, fg=TEXT, undo=False, relief="flat",
                                highlightthickness=1, highlightbackground=LINE, padx=8, pady=6)
        self.log.pack(fill="both", expand=True, pady=(4, 0))
        for tag, color in (("err", "#C01C1C"), ("warn", "#C26A00"), ("ok", "#1E7B34"), ("head", "#1F3864"), ("muted", "#888888")):
            self.log.tag_configure(tag, foreground=color)
        self.log.tag_configure("head", font=(MONO, 10, "bold"))

        # ---- 결과 파일
        tb2 = ttk.Frame(rs)
        tb2.pack(fill="x")
        ttk.Button(tb2, text="파일 열기", command=self.open_selected_file).pack(side="left")
        ttk.Button(tb2, text="폴더에서 보기", command=self.show_selected_file).pack(side="left", padx=4)
        ttk.Button(tb2, text="선택한 기능의 결과 폴더", command=lambda: self._open_task_folder()).pack(side="left")
        ttk.Button(tb2, text="새로 고침", command=self.refresh_results).pack(side="right")
        cols = ("time", "kind", "file")
        self.tree = ttk.Treeview(rs, columns=cols, show="headings", height=8)
        for c, h, w in (("time", "만든 시각", 140), ("kind", "분류", 110), ("file", "파일", 520)):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w")
        ysb = ttk.Scrollbar(rs, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        self.tree.pack(side="left", fill="both", expand=True, pady=(4, 0))
        ysb.pack(side="right", fill="y", pady=(4, 0))
        self.tree.bind("<Double-1>", lambda e: self.open_selected_file())
        self.paths = {}

    def _ui(self, fn):
        """작업 스레드 → 화면 갱신은 반드시 화면 스레드에서(큐로 넘겨 _poll에서 실행)."""
        self.q.put(("call", fn))

    def _over(self, w) -> bool:
        x, y = self.root.winfo_pointerxy()
        return str(self.root.winfo_containing(x, y) or "").startswith(str(w))

    # ------------------------------------------------------------ 기능 선택
    def select(self, key: str):
        t = tasks.BY_KEY[key]
        self.selected = key
        for k, l in self.name_lbls.items():
            l.configure(bg=SEL_BG if k == key else CARD, fg=(NAVY if k == key else (TEXT if tasks.BY_KEY[k].ready else "#A0A6AF")),
                        font=(FONT, 10, "bold") if k == key else (FONT, 10))
        self.d_name.configure(text=t.name)
        self.d_desc.configure(text=t.desc)
        self.d_status.configure(text="")
        for w in self.opt.winfo_children():
            w.destroy()
        self.opt_widgets = []
        if "period" in t.options:
            ttk.Label(self.opt, text="기간:").pack(side="left")
            for v, lab in tasks.PERIODS:
                rb = ttk.Radiobutton(self.opt, text=lab, value=v, variable=self.period)
                rb.pack(side="left", padx=2)
                self.opt_widgets.append(rb)
        if key in ("daily_all", "daily_summary"):
            self._daily_options()
        for w in self.opt_widgets:
            w.state(["disabled"] if self.proc else ["!disabled"])
        if t.status:
            threading.Thread(target=self._load_status, args=(t,), daemon=True).start()
        self.btn_run.state(["!disabled"] if t.ready and not self.proc else ["disabled"])

    def _load_status(self, t):
        try:
            s = t.status()
        except Exception as e:
            s = f"(상태 확인 불가: {type(e).__name__})"
        self._ui(lambda: self.selected == t.key and self.d_status.configure(text=s))

    def _daily_options(self):
        import daily
        f = ttk.Frame(self.opt)
        f.pack(fill="x")
        b1 = ttk.Button(f, text="실행할 기능 선택…", command=self.daily_steps)
        b1.pack(side="left")
        s = daily.load_settings()
        self.sched_time = tk.StringVar(value=s.get("time", "08:10"))
        ttk.Label(f, text="  매일").pack(side="left")
        e = ttk.Entry(f, textvariable=self.sched_time, width=6)
        e.pack(side="left")
        b2 = ttk.Button(f, text="자동 실행 예약", command=self.schedule)
        b2.pack(side="left", padx=2)
        b3 = ttk.Button(f, text="예약 해제", command=self.unschedule)
        b3.pack(side="left")
        self.opt_widgets += [b1, e, b2, b3]

    def daily_steps(self):
        import daily
        s = daily.load_settings()
        win = tk.Toplevel(self.root)
        win.configure(bg=BG)
        win.title("매일 실행할 기능")
        win.transient(self.root)
        vs = {}
        for k, n, _ in daily.STEPS:
            vs[k] = tk.BooleanVar(value=k in s["steps"])
            ttk.Checkbutton(win, text=n, variable=vs[k]).pack(anchor="w", padx=14, pady=2)

        def save():
            s["steps"] = [k for k, _, _ in daily.STEPS if vs[k].get()]
            daily.save_settings(s)
            self.append(f"매일 실행할 기능: {', '.join(n for k, n, _ in daily.STEPS if k in s['steps'])}\n", "muted")
            win.destroy()
        ttk.Button(win, text="저장", command=save).pack(pady=10)

    def schedule(self):
        import daily
        t = self.sched_time.get().strip()
        try:
            datetime.strptime(t, "%H:%M")
        except ValueError:
            messagebox.showwarning("시간 형식", "HH:MM 형식으로 입력하세요(예: 08:10).")
            return
        s = daily.load_settings()
        s["time"] = t
        daily.save_settings(s)
        self._run_inline(lambda: daily.schedule(t), f"매일 {t} 자동 실행 예약")

    def unschedule(self):
        import daily
        self._run_inline(daily.unschedule, "자동 실행 예약 해제")

    def _run_inline(self, fn, title):
        """짧은 작업(예약 등)은 화면 프로세스에서 바로 실행하고 출력만 로그에 표시."""
        import contextlib
        import io

        def work():
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                try:
                    fn()
                except Exception as e:
                    print(f"❌ {type(e).__name__}: {e}")
            self.q.put(("line", f"── {title} ──\n"))
            for line in buf.getvalue().splitlines():
                self.q.put(("line", line + "\n"))
            if self.selected:
                self._ui(lambda: self.select(self.selected))
        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------ 실행
    def run_selected(self):
        if self.selected:
            self.start_queue([self.selected])

    def run_checked(self):
        keys = [t.key for t in tasks.TASKS if self.checks[t.key].get() and t.ready]
        if not keys:
            messagebox.showinfo("기능 선택", "왼쪽 목록에서 실행할 기능을 체크하세요.")
            return
        self.start_queue(keys)

    def open_changes(self):
        if getattr(self, "_chg", None) and self._chg.win.winfo_exists():
            self._chg.win.lift()
            return
        import gui_changes
        self._chg = gui_changes.ChangesWindow(self, FONT, MONO, open_path)
        self._set_icon(self._chg.win)

    def start_queue(self, keys):
        if "changes" in keys:                 # 수정 창은 따로 열림(차례 실행에서는 빠짐)
            keys = [k for k in keys if k != "changes"]
            if not keys:
                self.open_changes()
                return
        if self.proc:
            messagebox.showinfo("실행 중", "다른 기능이 실행 중입니다. 끝난 뒤에 실행하거나 ‘중지’를 누르세요.")
            return
        keys = [k for k in keys if tasks.BY_KEY[k].ready]
        if not keys:
            return
        self.run_queue = list(keys)
        self.total = len(keys)
        self.stopped = False
        if len(keys) > 1:
            self.append(f"\n=== 차례로 실행: {' → '.join(tasks.BY_KEY[k].name for k in keys)} ===\n", "head")
        self._next()

    def _next(self):
        if not self.run_queue or self.stopped:
            self._idle()
            return
        key = self.run_queue.pop(0)
        t = tasks.BY_KEY[key]
        self.current = key
        self.started = time.time()
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        self.logfile = open(LOG_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_{key}.log", "w", encoding="utf-8")
        cmd = [python_exe(), "-u", str(BASE / "runner.py"), key, "--period", self.period.get()]
        # 여러 개를 차례로 실행할 때는 마지막 것만 자동으로 열기
        if self.auto_open.get() and not self.run_queue:
            cmd.append("--open")
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
        flags = 0x08000000 if IS_WIN else 0          # CREATE_NO_WINDOW
        self.append(f"\n━━━━━━━━ {t.name} 시작 {datetime.now():%H:%M:%S} ━━━━━━━━\n", "head")
        try:
            extra = {} if IS_WIN else {"start_new_session": True}       # macOS: 멈출 때 브라우저까지 함께 끄려고 새 그룹
            self.proc = subprocess.Popen(cmd, cwd=str(BASE), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                         stdin=subprocess.DEVNULL, env=env, creationflags=flags, **extra)
        except Exception as e:
            self.append(f"❌ 실행하지 못했습니다: {e}\n", "err")
            self.proc = None
            self._idle()
            return
        self.status.set(f"실행 중: {t.name}")
        self.status_dot.configure(fg=ORANGE)
        self.set_busy(True)
        done = self.total - len(self.run_queue) - 1
        self.qinfo.set(f"{done + 1}/{self.total}번째 기능" if self.total > 1 else "")
        threading.Thread(target=self._reader, args=(self.proc,), daemon=True).start()

    def _reader(self, proc):
        for raw in iter(proc.stdout.readline, b""):
            self.q.put(("line", raw.decode("utf-8", errors="replace").replace("\r\n", "\n")))
        proc.wait()
        self.q.put(("exit", proc.returncode))

    def stop(self):
        if not self.proc:
            return
        if not messagebox.askyesno("중지", "실행 중인 기능을 멈출까요?\n(열려 있는 브라우저도 닫힙니다)"):
            return
        self.stopped = True
        self.run_queue = []
        kill_tree(self.proc)
        self.append("\n■ 중지했습니다.\n", "warn")

    def set_busy(self, busy: bool):
        """실행 중: 진행 표시 + 다른 실행 버튼·체크 상자·옵션은 누를 수 없게. ‘중지’만 누를 수 있음."""
        st = ["disabled"] if busy else ["!disabled"]
        for w in [self.btn_run, self.btn_queue] + self.busy_widgets + self.opt_widgets:
            try:
                w.state(st)
            except Exception:
                pass
        self.btn_stop.state(["!disabled"] if busy else ["disabled"])
        if busy:
            self.pbar.start(15)
        else:
            self.pbar.stop()
            self.pbar["value"] = 0
            self.qinfo.set("")
        if not busy and self.selected and not tasks.BY_KEY[self.selected].ready:
            self.btn_run.state(["disabled"])

    def _idle(self):
        self.current = None
        self.proc = None
        self.status.set("대기 중")
        self.status_dot.configure(fg="#1E7B34")
        self.step.set("")
        self.set_busy(False)
        if self.selected:
            self.select(self.selected)
        self.refresh_results()
        if hasattr(self, "sum_tree"):
            self.load_summary()

    # ------------------------------------------------------------ 로그
    def append(self, text: str, tag: str = None):
        if tag is None:
            s = text.lstrip()
            tag = ("err" if s.startswith(("❌", "Traceback")) or "Error:" in s[:40] else
                   "warn" if s.startswith("⚠") else
                   "ok" if s.startswith(("✅", "📊", "📋", "📒")) else
                   "head" if s.startswith(("【", "▶", "──", "===", "━")) else None)
        # Windows의 Tk는 BMP 밖의 이모지(📁 📊 📋 등)를 그릴 때 글꼴을 찾느라 화면이 멈출 수 있어 바꿔서 표시
        text = _EMOJI_RE.sub(lambda m: _EMOJI_MAP.get(m.group(0), "•"), text)
        if len(text) > 800:                   # 아주 긴 줄은 화면에서만 줄임(로그 파일에는 전체)
            text = text[:800] + " …\n"
        self.log.insert("end", text, tag or ())

    def _after_append(self):
        n = int(self.log.index("end-1c").split(".")[0])
        if n > 8000:
            self.log.delete("1.0", f"{n - 8000}.0")
        if self.autoscroll.get():
            self.log.see("end")

    def _poll(self):
        n = 0
        try:
            while n < 300:                    # 한 번에 너무 많이 처리하지 않음(화면이 멈추지 않게)
                kind, val = self.q.get_nowait()
                n += 1
                if kind == "line":
                    if getattr(self, "logfile", None) and not self.logfile.closed:
                        self.logfile.write(val)
                    if val.startswith("@@OUTPUT@@"):
                        p = Path(val.split(" ", 1)[1].strip())
                        self.last_output = p
                        for b in (self.btn_last, self.btn_last_dir):
                            b.state(["!disabled"])
                        self.append(f"▶ 결과 파일: {p}\n", "ok")
                        continue
                    if val.startswith("@@DONE@@"):
                        continue
                    s = val.strip()
                    if s.startswith(("【", "────────── [", "▶")):
                        self.step.set(s[:90])
                    self.append(val)
                elif kind == "call":
                    try:
                        val()
                    except Exception:
                        pass
                elif kind == "exit":
                    t = tasks.BY_KEY.get(self.current or "", None)
                    mins = (time.time() - self.started) / 60
                    ok = val == 0 and not self.stopped
                    self.append(f"━━━━━━━━ {t.name if t else ''} {'완료' if ok else '실패/중지'} · {mins:.1f}분 ━━━━━━━━\n",
                                "ok" if ok else "err")
                    if getattr(self, "logfile", None):
                        self.logfile.close()
                    self.proc = None
                    if not self.run_queue and t:
                        self._notify_done(t.name if self.total <= 1 else f"{t.name} 등 {self.total}개 기능", ok, mins)
                    self._next()
        except queue.Empty:
            pass
        if n:
            self._after_append()
        self.root.after(50 if n >= 300 else 120, self._poll)

    def _tick(self):
        if self.proc and self.current:
            e = int(time.time() - self.started)
            more = f" · 남은 기능 {len(self.run_queue)}개" if self.run_queue else ""
            self.status.set(f"실행 중: {tasks.BY_KEY[self.current].name} · {e // 3600:02d}:{e % 3600 // 60:02d}:{e % 60:02d}{more}")
        self.root.after(1000, self._tick)

    # ------------------------------------------------------------ 오늘의 요약 탭
    def _build_summary(self, sm):
        tb = ttk.Frame(sm)
        tb.pack(fill="x")
        self.sum_title = tk.StringVar(value="오늘의 요약")
        ttk.Label(tb, textvariable=self.sum_title, font=(FONT, 11, "bold"), foreground="#1F3864").pack(side="left")
        ttk.Button(tb, text="새로 고침", command=self.load_summary).pack(side="right")
        b = ttk.Button(tb, text="요약 다시 만들기", command=lambda: self.start_queue(["daily_summary"]))
        b.pack(side="right", padx=4)
        self.busy_widgets.append(b)
        ttk.Button(tb, text="요약 Excel 열기", command=self.open_summary_xlsx).pack(side="right")
        cnt = ttk.Frame(sm)
        cnt.pack(fill="x", pady=(6, 2))
        self.sum_cnt = {}
        for lv, bg, fg in (("P0", "#F8D7DA", "#9B1C1C"), ("P1", "#FFE8CC", "#9A4A00"), ("P2", "#FFF6D6", "#6B5800"),
                           ("P3", "#EEF2F7", "#44546A")):
            v = tk.StringVar(value=f"{lv}  -")
            tk.Label(cnt, textvariable=v, bg=bg, fg=fg, font=(FONT, 11, "bold"), padx=14, pady=4).pack(side="left", padx=(0, 6))
            self.sum_cnt[lv] = v
        self.sum_filter = tk.StringVar(value="urgent")
        ttk.Radiobutton(cnt, text="전체", value="all", variable=self.sum_filter, command=self._fill_summary).pack(side="right")
        ttk.Radiobutton(cnt, text="P0·P1만", value="urgent", variable=self.sum_filter, command=self._fill_summary).pack(side="right", padx=6)
        self.sum_kpi = tk.StringVar(value="")
        ttk.Label(sm, textvariable=self.sum_kpi, wraplength=900, justify="left").pack(fill="x", pady=(4, 0))
        self.sum_steps = tk.StringVar(value="")
        ttk.Label(sm, textvariable=self.sum_steps, style="Muted.TLabel", wraplength=900, justify="left").pack(fill="x", pady=(2, 4))
        fr = ttk.Frame(sm)
        fr.pack(fill="both", expand=True)
        cols = ("lv", "kind", "what", "todo", "mod")
        self.sum_tree = ttk.Treeview(fr, columns=cols, show="headings")
        for c, h, w in (("lv", "등급", 46), ("kind", "분류", 120), ("what", "내용", 470), ("todo", "조치", 300), ("mod", "기능", 110)):
            self.sum_tree.heading(c, text=h)
            self.sum_tree.column(c, width=w, anchor="w", stretch=c in ("what", "todo"))
        for lv, bg in (("P0", "#F8D7DA"), ("P1", "#FFE8CC")):
            self.sum_tree.tag_configure(lv, background=bg)
        ysb = ttk.Scrollbar(fr, orient="vertical", command=self.sum_tree.yview)
        self.sum_tree.configure(yscrollcommand=ysb.set)
        self.sum_tree.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        self.sum_tree.bind("<Double-1>", lambda e: self._open_summary_row())
        ttk.Label(sm, text="줄을 더블클릭하면 그 기능의 결과 파일이 열립니다. 숫자는 각 기능을 실행할 때마다 쌓입니다(오늘 날짜 기준).",
                  style="Muted.TLabel").pack(anchor="w", pady=(4, 0))
        self.sum_rows, self.sum_reports = [], {}
        self.root.after(300, self.load_summary)

    def load_summary(self):
        threading.Thread(target=self._read_summary, daemon=True).start()

    def _read_summary(self):
        try:
            import daily
            import daily_bus as bus
            data = bus.load()
            alerts, cnt, _ = daily.summary_alerts(data)
            mods = data.get("modules", {})
            home = mods.get("홈 화면", {}).get("kpis", {})
            steps = data.get("steps", {})
            reports = {n: m.get("report") for n, m in mods.items() if m.get("report")}
            res = (alerts, cnt, home, steps, reports, None)
        except Exception as e:
            res = ([], {}, {}, {}, {}, f"{type(e).__name__}: {e}")
        self._ui(lambda: self._show_summary(*res))

    def _show_summary(self, alerts, cnt, home, steps, reports, err):
        self.sum_title.set(f"오늘의 요약 · {datetime.now():%Y-%m-%d} · {datetime.now():%H:%M} 기준" + (f"  (읽기 오류: {err})" if err else ""))
        for lv, v in self.sum_cnt.items():
            v.set(f"{lv}  {cnt.get(lv, 0)}건")
        kp = [f"{k}: {v}" for k, v in list(home.items())[:10] if not isinstance(v, (list, dict))]
        self.sum_kpi.set("홈 화면 — " + " · ".join(kp) if kp else "홈 화면 수치가 아직 없습니다(‘매일 자동 실행’ 또는 ‘오늘 할 일’을 실행하면 채워집니다).")
        st = [f"{v.get('name', k)} {'✔' if v.get('status') == '완료' else '✖' if v.get('status') == '실패' else '–'}"
              f"({v.get('minutes', 0):.0f}분)" for k, v in steps.items()]
        tot = sum(v.get("minutes", 0) or 0 for v in steps.values())
        if steps:
            st.append(f"합계 {tot:.0f}분")
        self.sum_steps.set("오늘 실행: " + "  ".join(st) if st else "오늘 매일 자동 실행 기록 없음")
        self.sum_rows, self.sum_reports = alerts, reports
        self._fill_summary()

    def _fill_summary(self):
        self.sum_tree.delete(*self.sum_tree.get_children())
        urgent = self.sum_filter.get() == "urgent"
        self.sum_iid = {}
        for a, mod in self.sum_rows:
            if urgent and a[0] not in ("P0", "P1"):
                continue
            clean = [_EMOJI_RE.sub("•", str(x)) for x in a[:4]]
            iid = self.sum_tree.insert("", "end", values=(*clean, mod), tags=(a[0],))
            self.sum_iid[iid] = mod
            if len(self.sum_iid) >= 500:
                break
        if urgent and not self.sum_iid and self.sum_rows:
            self.sum_tree.insert("", "end", values=("", "", "P0·P1(급한 일)은 없습니다. ‘전체’를 누르면 P2·P3도 보입니다.", "", ""))

    def _open_summary_row(self):
        s = self.sum_tree.selection()
        mod = self.sum_iid.get(s[0]) if s else None
        p = self.sum_reports.get(mod or "")
        if p and Path(p).exists():
            open_path(Path(p))
        elif mod:
            messagebox.showinfo("결과 파일", f"‘{mod}’의 결과 파일을 찾지 못했습니다.")

    def open_summary_xlsx(self):
        d = REPORTS / "daily"
        files = sorted(d.glob("daily_summary_*.xlsx")) if d.exists() else []
        if files:
            open_path(files[-1])
        else:
            messagebox.showinfo("오늘의 요약", "아직 요약 Excel이 없습니다. ‘요약 다시 만들기’를 누르세요.")

    def _notify_done(self, name: str, ok: bool, mins: float):
        """창이 뒤에 있거나 최소화돼 있을 때만 Windows 알림."""
        try:
            if self.root.focus_displayof() is not None and self.root.state() != "iconic":
                return
        except Exception:
            pass
        out = self.last_output
        def work():
            try:
                import notify
                notify.toast(f"{name} {'완료' if ok else '실패/중지'}", f"{mins:.1f}분 걸림" + (f"\n{out.name}" if out else ""), out)
            except Exception:
                pass
        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------ 결과 파일
    def refresh_results(self):
        """결과 파일 목록은 따로 찾음(폴더에 파일이 많아도 화면이 멈추지 않게)."""
        threading.Thread(target=self._scan_files, daemon=True).start()

    def _scan_files(self):
        files = []
        if REPORTS.exists():
            for root, dirs, names in os.walk(REPORTS):
                rel = Path(root).relative_to(REPORTS).parts
                dirs[:] = [d for d in dirs if d not in SKIP_DIRS and len(rel) < 2]
                for n in names:
                    if n.endswith(".xlsx") and not n.startswith("~$"):
                        p = Path(root) / n
                        try:
                            files.append((p.stat().st_mtime, p))
                        except OSError:
                            pass
        files.sort(reverse=True)
        self._ui(lambda: self._fill_tree(files))

    def _fill_tree(self, files):
        self.tree.delete(*self.tree.get_children())
        self.paths = {}
        for mt, p in files[:80]:
            kind = FOLDER_NAME.get(p.relative_to(REPORTS).parts[0], p.parent.name)
            iid = self.tree.insert("", "end", values=(datetime.fromtimestamp(mt).strftime("%m-%d %H:%M"), kind,
                                                      str(p.relative_to(REPORTS))))
            self.paths[iid] = p

    def _sel_file(self):
        s = self.tree.selection()
        return self.paths.get(s[0]) if s else None

    def open_selected_file(self):
        p = self._sel_file()
        if p:
            open_path(p)

    def show_selected_file(self):
        p = self._sel_file()
        if p:
            show_in_folder(p)

    def _open_task_folder(self):
        t = tasks.BY_KEY.get(self.selected or "")
        p = t.folder_path if t else REPORTS
        p.mkdir(parents=True, exist_ok=True)
        open_path(p)

    # ------------------------------------------------------------ 설정(SP-API·확인 메일)
    def _refresh_api_state(self):
        try:
            import ads_api
            import sp_api
            cfg = sp_api.load_config()
            if not cfg.get("use_api"):
                a = "SP-API: 사용 안 함"
            elif sp_api.ready(cfg):
                a = "SP-API: 설정됨"
            else:
                a = "SP-API: 미설정"
            if not cfg.get("use_ads_api", True):
                b = "Ads API: 사용 안 함"
            elif ads_api.ready(cfg):
                b = "Ads API: 설정됨"
            else:
                b = "Ads API: 미설정"
            warn = " · 토큰 교체 권장" if sp_api.token_reminders(cfg) else ""
            miss = sp_api.profile_missing(cfg)
            if miss:
                warn += " · 설정 필요: " + ", ".join(miss)
            self.api_state.set(f"{a} · {b}{warn}")
        except Exception as e:
            self.api_state.set(f"API 설정 오류({type(e).__name__})")

    def open_settings(self):
        import ads_api
        import sp_api
        cfg = sp_api.load_config()
        win = tk.Toplevel(self.root)
        win.title("설정 — SP-API · Amazon Ads API · 메일·알림")
        win.configure(bg=BG)
        win.transient(self.root)
        win.resizable(True, True)
        outer = ttk.Frame(win, padding=10)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text=("비밀값은 이 컴퓨터의 Windows 계정으로 암호화되어 %LOCALAPPDATA%\\AmazonAIProject 에만 저장됩니다(OneDrive에 올라가지 않음)."
                   if IS_WIN else "비밀값은 이 Mac의 키체인(Keychain)에 저장되고, 설정은 ~/Library/Application Support/AmazonAIProject 에만 저장됩니다(OneDrive에 올라가지 않음)."),
                  style="Muted.TLabel", wraplength=700).pack(anchor="w", pady=(0, 6))
        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True)
        t_sp, t_ads, t_etc = (ttk.Frame(nb, padding=12) for _ in range(3))
        nb.add(t_sp, text="  SP-API(셀러 센트럴)  ")
        nb.add(t_ads, text="  Amazon Ads API(광고)  ")
        nb.add(t_etc, text="  메일·알림·토큰 교체  ")
        vars_, secrets = {}, []

        def field(frm, row, key, label, secret=False, width=70):
            ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", pady=3, padx=(0, 8))
            v = tk.StringVar(value=cfg.get(key, ""))
            e = ttk.Entry(frm, textvariable=v, width=width, show="•" if secret else "")
            e.grid(row=row, column=1, sticky="we", pady=3)
            vars_[key] = v
            if secret:
                secrets.append(e)
            return e

        def saved_label(frm, row, key):
            age = sp_api.token_age(cfg.get(key, ""))
            txt = (f"토큰 저장일: {cfg.get(key)} ({age}일 전)" if age is not None else "토큰 저장일: 기록 없음(교체 권장)")
            if age is not None and age >= sp_api.TOKEN_MAX_DAYS:
                txt += " — 교체 권장"
            ttk.Label(frm, text=txt, style="Muted.TLabel").grid(row=row, column=1, sticky="w")

        # ---- SP-API
        ttk.Label(t_sp, text="Seller Central → Apps and Services → Develop Apps → 앱의 LWA credentials",
                  font=(FONT, 10, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        field(t_sp, 1, "client_id", "LWA Client ID")
        field(t_sp, 2, "client_secret", "LWA Client Secret", True)
        field(t_sp, 3, "refresh_token", "Refresh Token", True)
        saved_label(t_sp, 4, "token_saved")
        field(t_sp, 5, "seller_id", "Seller ID(판매자 토큰)", width=30).grid(sticky="w")
        field(t_sp, 6, "brand", "브랜드 이름", width=30).grid(sticky="w")
        use = tk.BooleanVar(value=bool(cfg.get("use_api", True)))
        ttk.Checkbutton(t_sp, text="SP-API 우선 사용(실패하면 자동으로 원래 브라우저 방식)", variable=use).grid(row=7, column=1, sticky="w", pady=(6, 2))
        t_sp.columnconfigure(1, weight=1)

        # ---- Amazon Ads API
        ttk.Label(t_ads, text="광고 보고서를 API로 받습니다(조회만). 안 되면 광고 콘솔 화면으로 받습니다.",
                  font=(FONT, 10, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        field(t_ads, 1, "ads_client_id", "Ads Client ID")
        field(t_ads, 2, "ads_client_secret", "Ads Client Secret", True)
        field(t_ads, 3, "ads_redirect_uri", "Redirect URI(앱의 Allowed Return URL)")
        ttk.Separator(t_ads).grid(row=4, column=0, columnspan=3, sticky="we", pady=8)
        ttk.Label(t_ads, text="Refresh Token 받기: ① 버튼 → Amazon에서 승인(Allow) → 이동한 페이지의 주소창 전체를 복사해 아래에 붙여 넣고 ②",
                  wraplength=680).grid(row=5, column=0, columnspan=3, sticky="w")
        bf1 = ttk.Frame(t_ads)
        bf1.grid(row=6, column=0, columnspan=3, sticky="we", pady=4)
        code = tk.StringVar()
        ttk.Button(bf1, text="① 인증 페이지 열기", command=lambda: auth_open()).pack(side="left")
        ttk.Entry(bf1, textvariable=code, width=60).pack(side="left", padx=6, fill="x", expand=True)
        ttk.Button(bf1, text="② 토큰 받기", command=lambda: auth_code()).pack(side="left")
        field(t_ads, 7, "ads_refresh_token", "Ads Refresh Token", True)
        saved_label(t_ads, 8, "ads_token_saved")
        field(t_ads, 9, "ads_profile_id", "프로필 ID(미국 판매자)", width=30).grid(sticky="w")
        prof_box = ttk.Combobox(t_ads, state="readonly", width=70)
        prof_box.grid(row=10, column=1, sticky="we", pady=3)
        ttk.Button(t_ads, text="프로필 찾기", command=lambda: find_prof()).grid(row=9, column=2, padx=6)
        prof_box.bind("<<ComboboxSelected>>", lambda e: vars_["ads_profile_id"].set(prof_box.get().split(" ")[0]))
        use_ads = tk.BooleanVar(value=bool(cfg.get("use_ads_api", True)))
        ttk.Checkbutton(t_ads, text="Ads API 우선 사용(실패하면 자동으로 광고 콘솔 화면 방식)", variable=use_ads).grid(row=11, column=1, sticky="w", pady=(6, 2))
        t_ads.columnconfigure(1, weight=1)

        # ---- 메일·알림·토큰 교체
        ttk.Label(t_etc, text="확인 메일(광고 보고서 Send to)").grid(row=0, column=0, sticky="w", padx=(0, 8))
        mail = tk.StringVar(value=cfg.get("report_email", ""))
        ttk.Entry(t_etc, textvariable=mail, width=40).grid(row=0, column=1, sticky="w")
        noti = tk.BooleanVar(value=bool(cfg.get("notify", True)))
        ttk.Checkbutton(t_etc, text="알림 보내기(Windows 알림 · Mac 알림 센터 — 매일 자동 실행이 끝났을 때 · 이 창이 뒤에 있을 때 기능이 끝나면)",
                        variable=noti).grid(row=1, column=1, sticky="w", pady=(8, 2))
        ttk.Button(t_etc, text="알림 시험", command=lambda: threading.Thread(
            target=lambda: __import__("notify").toast("SellerMorning", "알림 시험입니다. 누르면 reports 폴더가 열립니다.", REPORTS, force=True),
            daemon=True).start()).grid(row=1, column=2, padx=6)
        ttk.Separator(t_etc).grid(row=2, column=0, columnspan=3, sticky="we", pady=10)
        rot = ("토큰 교체 방법(대화·메일 등에 토큰이 노출됐거나 오래됐을 때)\n"
               "SP-API: Seller Central → Apps and Services → Develop Apps → 앱의 ‘Authorize’(또는 ‘Edit App’ → 다시 승인)\n"
               "   → 새 Refresh Token을 복사해 SP-API 탭에 붙여 넣고 저장. 그 다음 Manage Your Apps / Authorized apps에서\n"
               "   예전 승인을 지우면(Revoke) 예전 토큰은 더 이상 쓸 수 없습니다. Client Secret도 ‘Rotate’로 바꿀 수 있습니다.\n"
               "Ads API: Amazon Ads API 탭의 ① → ② 로 새 토큰을 받으면 자동으로 바뀝니다.\n"
               "   예전 승인은 amazon.com → 계정 → 로그인 및 보안 → ‘Login with Amazon’ 승인 앱 관리에서 지울 수 있습니다.\n"
               "토큰을 바꿔 저장하면 저장일이 기록되고, 180일이 지나면 오늘의 요약에 교체 알림(P3)이 나옵니다.")
        ttk.Label(t_etc, text=rot, justify="left", wraplength=700).grid(row=3, column=0, columnspan=3, sticky="w")
        t_etc.columnconfigure(1, weight=1)

        show = tk.BooleanVar(value=False)
        result = tk.StringVar(value="")
        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(8, 0))
        ttk.Checkbutton(bottom, text="비밀값 보기", variable=show,
                        command=lambda: [e.configure(show="" if show.get() else "•") for e in secrets]).pack(side="left")
        ttk.Button(bottom, text="닫기", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(bottom, text="저장", command=lambda: save()).pack(side="right", padx=4)
        ttk.Button(bottom, text="연결 테스트", command=lambda: test()).pack(side="right", padx=4)
        ttk.Button(bottom, text="전체 점검", command=lambda: full()).pack(side="right", padx=4)
        res_box = tk.Text(outer, height=9, wrap="word", font=(FONT, 9), bg="#FBFBFB", relief="flat")
        res_box.pack(fill="both", expand=False, pady=(8, 0))

        def say(msg):
            res_box.delete("1.0", "end")
            res_box.insert("end", _EMOJI_RE.sub("•", msg))

        def current():
            c = dict(cfg)
            c.update({k: v.get().strip() for k, v in vars_.items()})
            c["use_api"], c["use_ads_api"], c["notify"] = use.get(), use_ads.get(), noti.get()
            c["report_email"] = mail.get().strip()
            return c

        def bg(fn, wait_msg):
            say(wait_msg)

            def work():
                try:
                    msg = fn()
                except Exception as e:
                    msg = f"❌ {type(e).__name__}: {str(e)[:400]}"
                self._ui(lambda: say(msg))
            threading.Thread(target=work, daemon=True).start()

        def save():
            c = current()
            if c["report_email"] and not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", c["report_email"]):
                messagebox.showwarning("메일 주소", "메일 주소 형식을 확인하세요.", parent=win)
                return
            sp_api.save_config(c)
            sp_api._client = None
            cfg.clear()
            cfg.update(sp_api.load_config())
            self._refresh_api_state()
            say("✅ 저장했습니다.")

        def auth_open():
            c = current()
            if not c.get("ads_client_id"):
                messagebox.showwarning("Ads API", "Ads Client ID를 먼저 입력하세요.", parent=win)
                return
            __import__("webbrowser").open(ads_api.consent_url(c))
            say("브라우저에서 Amazon 로그인 → ‘Allow(허용)’ → 이동한 페이지(열리지 않아도 괜찮음)의 주소창 전체를 복사해 붙여 넣고 ② 토큰 받기.\n"
                "Redirect URI가 Ads 앱(LWA 보안 프로필)의 Allowed Return URLs에 등록돼 있어야 합니다.")

        def auth_code():
            c = current()
            if not (c.get("ads_client_id") and c.get("ads_client_secret") and code.get().strip()):
                messagebox.showwarning("Ads API", "Client ID · Client Secret · 이동한 주소(code)를 모두 넣으세요.", parent=win)
                return

            def fn():
                rt = ads_api.exchange_code(c, code.get())
                self._ui(lambda: (vars_["ads_refresh_token"].set(rt), code.set("")))
                c2 = dict(c, ads_refresh_token=rt)
                prof = ads_api.find_profiles(c2)
                self._ui(lambda: fill_prof(prof))
                return "✅ Refresh Token을 받았습니다. 프로필을 고른 뒤 ‘저장’을 누르세요."
            bg(fn, "토큰 받는 중...")

        def fill_prof(prof):
            us = [p for p in prof if p["country"] == "US"]
            items = [f"{p['profileId']}  {p['name']} · {p['country']} · {p['type']}" for p in (us + [p for p in prof if p not in us])]
            prof_box["values"] = items
            if not vars_["ads_profile_id"].get():
                pick = next((p for p in us if p["type"] == "seller"), us[0] if us else None)
                if pick:
                    vars_["ads_profile_id"].set(pick["profileId"])
            for it in items:
                if it.startswith(vars_["ads_profile_id"].get() + " "):
                    prof_box.set(it)

        def find_prof():
            c = current()

            def fn():
                prof = ads_api.find_profiles(c)
                self._ui(lambda: fill_prof(prof))
                return f"프로필 {len(prof)}개를 찾았습니다(미국 판매자 프로필을 자동으로 골랐습니다. 다르면 목록에서 고르세요)." if prof else \
                    "프로필이 없습니다 — 이 Amazon 계정으로 광고 콘솔에 들어갈 수 있는지 확인하세요."
            bg(fn, "프로필 찾는 중...")

        def test():
            c = current()

            def fn():
                lines = []
                if c.get("client_id") or c.get("refresh_token"):
                    try:
                        lines.append("SP-API: " + sp_api.test_connection(c))
                    except Exception as e:
                        lines.append(f"SP-API: ❌ {type(e).__name__}: {str(e)[:250]}")
                if c.get("ads_client_id") or c.get("ads_refresh_token"):
                    try:
                        lines.append("Ads API: " + ads_api.test(c))
                    except Exception as e:
                        lines.append(f"Ads API: ❌ {type(e).__name__}: {str(e)[:250]}")
                return "\n".join(lines) or "입력된 API 설정이 없습니다."
            bg(fn, "연결 확인 중...")

        def full():
            c = current()
            bg(lambda: "\n".join(f"{r['결과']} {r['항목']}: {r['내용']}" for r in sp_api.self_check(c)), "전체 점검 중...(약 10~30초)")


    # ------------------------------------------------------------ 기타
    def make_shortcut(self):
        if IS_MAC:          # 바탕화면에 더블클릭하면 열리는 .command 파일
            dst = Path.home() / "Desktop" / "SellerMorning.command"
            dst.write_text(f'#!/bin/bash\ncd "{BASE.parent}"\nexec /bin/bash "{BASE.parent / "start_gui.command"}"\n', encoding="utf-8")
            os.chmod(dst, 0o755)
            messagebox.showinfo("바로가기", "바탕화면에 ‘SellerMorning.command’를 만들었습니다(더블클릭하면 열림).")
            return
        if not IS_WIN:
            messagebox.showinfo("바로가기", "Windows·macOS에서만 만들 수 있습니다.")
            return
        exe = Path(python_exe())
        pyw = exe.parent / "pythonw.exe"
        target = pyw if pyw.exists() else exe
        ps = ("$d=[Environment]::GetFolderPath('Desktop');"
              "$s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d 'SellerMorning.lnk'));"
              f"$s.TargetPath='{target}';$s.Arguments='\"{BASE / 'gui.py'}\"';$s.WorkingDirectory='{BASE}';"
              f"$s.IconLocation='{ICON_ICO if ICON_ICO.exists() else 'shell32.dll,43'}';$s.Save()")
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True,
                           creationflags=0x08000000)
        if r.returncode == 0:
            messagebox.showinfo("바로가기", "바탕화면에 ‘SellerMorning’ 바로가기를 만들었습니다.")
        else:
            messagebox.showerror("바로가기", r.stderr[:400] or "만들지 못했습니다.")

    def on_close(self):
        if self.proc:
            if not messagebox.askyesno("종료", "실행 중인 기능이 있습니다. 멈추고 종료할까요?"):
                return
            self.stopped = True
            kill_tree(self.proc)
        self.root.destroy()


def main():
    if IS_WIN:
        try:   # 고해상도 화면에서 글자가 흐리지 않게
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        try:   # 작업 표시줄에 python 대신 이 프로그램 아이콘
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("SellerMorning.App")
        except Exception:
            pass
    root = tk.Tk()
    try:
        App(root)
    except Exception:
        import traceback
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        (LOG_DIR / "gui_error.log").write_text(traceback.format_exc(), encoding="utf-8")
        messagebox.showerror("시작 오류", f"화면을 시작하지 못했습니다.\n{LOG_DIR / 'gui_error.log'}")
        raise
    root.mainloop()


if __name__ == "__main__":
    main()
