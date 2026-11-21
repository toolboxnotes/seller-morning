"""
‘리스팅·가격 수정(SP-API)’ 창 — 요청서 → 검사(실제 변경 없음) → 고른 줄만 적용 → 기록·되돌리기
"""

import threading
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText

NAVY, ORANGE, BG, CARD, MUTED, TEXT = "#232F3E", "#FF9900", "#F3F4F6", "#FFFFFF", "#6B7280", "#1F2937"
ON, OFF = "☑", "☐"


class ChangesWindow:
    def __init__(self, app, font: str, mono: str, open_path):
        import changes
        self.app, self.ch, self.open_path, self.font = app, changes, open_path, font
        self.rows, self.iid = [], {}
        self.busy = False
        w = self.win = tk.Toplevel(app.root)
        w.title("리스팅·가격 수정(SP-API) — 확인한 것만 적용")
        w.geometry("1250x760")
        w.configure(bg=BG)
        f = ttk.Frame(w, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="① 요청서에 바꿀 내용과 적용(Y) 표시 → ② 검사(Amazon이 오류만 확인, 실제 변경 없음) → ③ 적용할 줄을 골라 ‘적용’",
                  font=(font, 11, "bold"), foreground=NAVY).pack(anchor="w")
        ttk.Label(f, text="모든 변경은 원래 값과 함께 기록되어 ‘변경 기록·되돌리기’에서 되돌릴 수 있습니다. 반영까지 몇 분~몇 시간 걸립니다.",
                  style="Muted.TLabel").pack(anchor="w", pady=(2, 8))
        tb = ttk.Frame(f)
        tb.pack(fill="x")
        self.btns = []
        for text, cmd in (("① 빈 요청서 만들기", self.make), ("초안 채운 요청서(제목·하이라이트·가격표)", self.make_drafts),
                          ("요청서 열기", self.open_req), ("② 검사", self.check)):
            b = ttk.Button(tb, text=text, command=cmd)
            b.pack(side="left", padx=(0, 6))
            self.btns.append(b)
        ttk.Button(tb, text="변경 기록·반영 확인", command=self.history).pack(side="right")
        ttk.Button(tb, text="기록 Excel", command=lambda: self.ch.LOG_XLSX.exists() and self.open_path(self.ch.LOG_XLSX)).pack(side="right", padx=6)

        cols = ("sel", "sku", "field", "old", "new", "status", "profit")
        fr = ttk.Frame(f)
        fr.pack(fill="both", expand=True, pady=(8, 0))
        self.tree = ttk.Treeview(fr, columns=cols, show="headings", selectmode="browse")
        for c, h, wd in (("sel", "적용", 50), ("sku", "SKU", 150), ("field", "항목", 60), ("old", "현재 값", 330),
                         ("new", "새 값", 280), ("status", "검사 결과", 260), ("profit", "가격 변경 시 수수료·이익", 330)):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=wd, anchor="center" if c == "sel" else "w", stretch=c in ("old", "new", "status", "profit"))
        self.tree.tag_configure("ok", background="#EAF6EC")
        self.tree.tag_configure("bad", background="#FDECEC")
        self.tree.tag_configure("big", background="#FFF1DB")
        sb = ttk.Scrollbar(fr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self._click)
        self.tree.bind("<Double-1>", self._detail)

        bot = ttk.Frame(f)
        bot.pack(fill="x", pady=8)
        self.allow_big = tk.BooleanVar(value=False)
        ttk.Checkbutton(bot, text="큰 가격 변경(±30% 초과·손해 가격)도 적용", variable=self.allow_big).pack(side="left")
        self.info = tk.StringVar(value="요청서를 만들거나 열어서 Y를 표시한 뒤 ‘② 검사’를 누르세요.")
        ttk.Label(bot, textvariable=self.info, style="Muted.TLabel").pack(side="left", padx=12)
        self.btn_apply = ttk.Button(bot, text="③ 선택한 줄 적용", style="Run.TButton", command=self.apply)
        self.btn_apply.pack(side="right")
        self.btn_apply.state(["disabled"])
        ttk.Button(bot, text="전체 선택/해제", command=self._toggle_all).pack(side="right", padx=6)
        self.log = ScrolledText(f, height=8, font=(mono, 10), bg=CARD, relief="flat", highlightthickness=1,
                                highlightbackground="#D9DDE3")
        self.log.pack(fill="x")

    # ------------------------------------------------------------ 공통
    def say(self, text: str):
        self.app._ui(lambda: (self.log.insert("end", text + "\n"), self.log.see("end")))

    def _bg(self, fn, done=None):
        if self.busy:
            return
        self.busy = True
        for b in self.btns + [self.btn_apply]:
            b.state(["disabled"])

        def work():
            try:
                res = fn()
            except Exception as e:
                res = None
                self.say(f"❌ {type(e).__name__}: {str(e)[:300]}")
            self.app._ui(lambda: self._end(done, res))
        threading.Thread(target=work, daemon=True).start()

    def _end(self, done, res):
        self.busy = False
        for b in self.btns:
            b.state(["!disabled"])
        if done:
            done(res)
        self._count()

    # ------------------------------------------------------------ 요청서
    def _confirm_new(self) -> bool:
        return not self.ch.REQ_FILE.exists() or messagebox.askyesno(
            "요청서 만들기", "요청서가 이미 있습니다. 지금 요청서는 백업해 두고 새로 만들까요?", parent=self.win)

    def make(self):
        """빈 요청서(바로 만들어짐): 머리글·항목 선택 목록·안내만."""
        if not self._confirm_new():
            return
        p = self.ch.make_request()
        self.say(f"✅ 빈 요청서를 만들었습니다: {p}")
        self.open_path(p)

    def make_drafts(self):
        """Listing 점검의 제목·Item Highlights 초안 + 현재 가격표(SP-API, 1~3분)를 채운 요청서."""
        if not self._confirm_new():
            return
        self.say("초안 채운 요청서 만드는 중(현재 가격표를 SP-API로 받는 데 1~3분 걸릴 수 있음)...")

        def fn():
            prices = self.ch.current_prices()
            p = self.ch.make_request(prices, drafts=True)
            self.say(f"✅ 요청서를 만들었습니다: {p} · 초안 {len(self.ch.title_drafts())}개 SKU · 가격표 {len(prices)}개 SKU "
                     "(적용 칸은 비어 있음 — 검토 후 Y)")
            return p
        self._bg(fn, lambda p: p and self.open_path(p))

    def open_req(self):
        if not self.ch.REQ_FILE.exists():
            messagebox.showinfo("요청서", "요청서가 없습니다. ‘① 빈 요청서 만들기’를 먼저 누르세요.", parent=self.win)
            return
        self.open_path(self.ch.REQ_FILE)

    # ------------------------------------------------------------ 검사
    def check(self):
        try:
            reqs = self.ch.load_requests()
        except PermissionError:
            messagebox.showwarning("요청서", "요청서를 읽지 못했습니다. Excel에서 저장했는지 확인하세요.", parent=self.win)
            return
        if not reqs:
            messagebox.showinfo("검사", "요청서에 적용(Y)으로 표시된 줄이 없습니다.\nY를 넣고 저장한 뒤 다시 누르세요.", parent=self.win)
            return
        self.say(f"\n② 검사 시작: {len(reqs)}건(실제 변경 없음)")
        self._bg(lambda: self.ch.check(reqs, log=self.say), self._show)

    def _show(self, rows):
        self.rows = rows or []
        self.tree.delete(*self.tree.get_children())
        self.iid = {}
        for r in self.rows:
            r["sel"] = bool(r.get("ok")) and not r.get("big")
            tag = "bad" if not r.get("ok") else ("big" if r.get("big") else "ok")
            iid = self.tree.insert("", "end", values=(ON if r["sel"] else OFF, r["sku"], r["field"], str(r.get("old", ""))[:300],
                                                     str(r.get("new", ""))[:300], r.get("status", ""), r.get("profit", "")),
                                    tags=(tag,))
            self.iid[iid] = r
        ok = sum(1 for r in self.rows if r.get("ok"))
        self.say(f"검사 끝: 적용 가능 {ok}건 / 전체 {len(self.rows)}건. 줄을 더블클릭하면 전체 내용, ‘적용’ 칸을 누르면 선택/해제.")

    def _click(self, e):
        if self.tree.identify_column(e.x) != "#1":
            return
        iid = self.tree.identify_row(e.y)
        r = self.iid.get(iid)
        if not r or not r.get("ok") or self.busy:
            return
        r["sel"] = not r["sel"]
        self.tree.set(iid, "sel", ON if r["sel"] else OFF)
        self._count()

    def _toggle_all(self):
        okr = [r for r in self.rows if r.get("ok")]
        on = not all(r["sel"] for r in okr) if okr else False
        for iid, r in self.iid.items():
            if r.get("ok"):
                r["sel"] = on
                self.tree.set(iid, "sel", ON if on else OFF)
        self._count()

    def _count(self):
        sel = [r for r in self.rows if r.get("sel") and r.get("ok")]
        if self.rows:
            big = sum(1 for r in sel if r.get("big"))
            self.info.set(f"선택 {len(sel)}건" + (f"(큰 가격 변경 {big}건)" if big else ""))
        self.btn_apply.state(["!disabled"] if sel and not self.busy else ["disabled"])

    def _detail(self, e):
        r = self.iid.get(self.tree.identify_row(e.y))
        if r:
            messagebox.showinfo(f"{r['sku']} · {r['field']}",
                                f"현재 값:\n{r.get('old', '')}\n\n새 값:\n{r.get('new', '')}\n\n검사 결과:\n{r.get('status', '')}"
                                + (f"\n\n수수료·이익:\n{r['profit']}" if r.get("profit") else ""),
                                parent=self.win)

    # ------------------------------------------------------------ 적용
    def apply(self):
        sel = [r for r in self.rows if r.get("sel") and r.get("ok")]
        if not sel:
            return
        big = [r for r in sel if r.get("big")]
        if big and not self.allow_big.get():
            if not messagebox.askyesno("큰 가격 변경", f"현재 가격 ±30%를 넘는 변경 {len(big)}건은 건너뜁니다. 나머지만 적용할까요?",
                                       parent=self.win):
                return
        kinds = {}
        for r in sel:
            kinds[r["field"]] = kinds.get(r["field"], 0) + 1
        msg = (f"Amazon에 실제로 적용합니다: {len(sel)}건\n" + ", ".join(f"{k} {v}건" for k, v in kinds.items()) +
               "\n\n원래 값은 기록되어 되돌릴 수 있습니다. 계속할까요?")
        if not messagebox.askyesno("적용 확인", msg, icon="warning", parent=self.win):
            return
        self.say(f"\n③ 적용 시작: {len(sel)}건")
        allow = self.allow_big.get()
        for r in self.rows:
            if not r.get("sel"):
                r["ok_backup"], r["ok"] = r.get("ok"), False      # 선택하지 않은 줄은 적용 안 함
        self._bg(lambda: self.ch.apply(self.rows, allow_big=allow, log=self.say), self._applied)

    def _applied(self, done):
        for r in self.rows:
            if "ok_backup" in r:
                r["ok"] = r.pop("ok_backup")
        n = sum(1 for r in done or [] if str(r.get("result", "")).startswith("접수"))
        self.say(f"적용 끝: 접수 {n}건. 기록: {self.ch.LOG_XLSX}")
        for iid, r in self.iid.items():
            if r.get("result"):
                self.tree.set(iid, "status", r["result"])
                if str(r["result"]).startswith("접수"):
                    r["ok"], r["sel"] = False, False
                    self.tree.set(iid, "sel", "완료")
        messagebox.showinfo("적용 끝", f"접수 {n}건. Amazon 반영까지 몇 분~몇 시간 걸립니다.\n"
                                       "10~30분 뒤 ‘변경 기록·반영 확인’ → ‘반영 확인’으로 실제 반영 여부를 보세요.\n"
                                       "(매일 자동 실행 때도 자동으로 확인합니다)\n요청서의 해당 줄은 ‘완료’로 바뀝니다.",
                            parent=self.win)

    # ------------------------------------------------------------ 기록·되돌리기
    def history(self):
        h = tk.Toplevel(self.win)
        h.title("변경 기록 · 반영 확인 · 되돌리기")
        h.geometry("1320x600")
        h.configure(bg=BG)
        f = ttk.Frame(h, padding=10)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="‘반영 확인’: 현재 값을 Amazon에서 다시 읽어 새 값과 비교(선택한 기록, 선택 안 하면 아직 반영 확인 안 된 것 전부). "
                          "되돌리기: 기록을 고르고(Ctrl·Shift로 여러 개) ‘선택한 기록 되돌리기’.",
                  style="Muted.TLabel", wraplength=1280).pack(anchor="w")
        self.h_sum = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.h_sum, font=(self.font, 10, "bold"), foreground=NAVY).pack(anchor="w", pady=(4, 0))
        cols = ("status", "time", "conf", "checked", "kind", "sku", "field", "old", "new", "msg")
        fr = ttk.Frame(f)
        fr.pack(fill="both", expand=True, pady=6)
        tv = ttk.Treeview(fr, columns=cols, show="headings", selectmode="extended")
        for c, t, wd in (("status", "상태", 90), ("time", "적용 시각", 120), ("conf", "반영 확인 시각", 120),
                         ("checked", "마지막 확인", 120), ("kind", "작업", 70), ("sku", "SKU", 130), ("field", "항목", 55),
                         ("old", "원래 값", 200), ("new", "새 값", 200), ("msg", "Amazon 메시지", 260)):
            tv.heading(c, text=t)
            tv.column(c, width=wd, anchor="w", stretch=c in ("old", "new", "msg"))
        for tag, bg in (("반영됨", "#EAF6EC"), ("제출값 저장됨", "#E8F1FB"), ("처리 중", "#FFF8E1"), ("접수됨", "#FFF8E1"), ("반영 안 됨", "#FDECEC"),
                        ("되돌림", "#EEF1F5"), ("이후 변경으로 대체", "#EEF1F5")):
            tv.tag_configure(tag, background=bg)
        sb = ttk.Scrollbar(fr, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        ids = {}
        fmt = self.ch._fmt

        def load():
            tv.delete(*tv.get_children())
            ids.clear()
            cnt = {}
            for e in self.ch.history():
                st = e.get("status", "접수됨")
                cnt[st] = cnt.get(st, 0) + 1
                iid = tv.insert("", "end", tags=(st,), values=(
                    st, fmt(e["time"]), fmt(e.get("confirmed_at", "")), fmt(e.get("checked_at", "")), e.get("kind", ""),
                    e["sku"], e["field"], str(e.get("old", ""))[:200], str(e.get("new", ""))[:200], e.get("issues", "")))
                ids[iid] = e
            order = [k for k in ("반영됨", "제출값 저장됨", "처리 중", "접수됨", "반영 안 됨", "되돌림", "이후 변경으로 대체") if cnt.get(k)]
            self.h_sum.set("  ·  ".join(f"{k} {cnt[k]}건" for k in order) or "기록이 없습니다.")
        load()

        def do_verify():
            sel = [ids[i]["id"] for i in tv.selection()]
            self.say("\n반영 확인 시작" + (f": 선택한 {len(sel)}건" if sel else ": 아직 반영 확인 안 된 기록 전부"))
            self._bg(lambda: self.ch.verify(sel or None, log=self.say),
                     lambda cnt: (load(), self.say("반영 확인 끝: " + ", ".join(f"{k} {v}건" for k, v in (cnt or {}).items()))))

        def do_revert():
            sel = [ids[i] for i in tv.selection() if ids[i].get("kind") == "적용" and not ids[i].get("reverted")]
            if not sel:
                messagebox.showinfo("되돌리기", "되돌릴 수 있는 ‘적용’ 기록을 고르세요.", parent=h)
                return
            if not messagebox.askyesno("되돌리기 확인", f"{len(sel)}건을 적용 전 값으로 되돌립니다(Amazon에 실제 변경). 계속할까요?",
                                       icon="warning", parent=h):
                return
            self.say(f"\n되돌리기 시작: {len(sel)}건")
            self._bg(lambda: self.ch.revert([e["id"] for e in sel], log=self.say), lambda _: load())

        def detail(e_):
            e = ids.get(tv.identify_row(e_.y))
            if e:
                messagebox.showinfo(f"{e['sku']} · {e['field']} · {e.get('status', '')}",
                                    f"확인 기준: {e.get('verified_by', '') or '-'}\n적용 시각: {fmt(e['time'])}\n"
                                    f"제출값 저장 확인: {fmt(e.get('submitted_at', '')) or '-'}\n반영 확인 시각: {fmt(e.get('confirmed_at', '')) or '-'}\n"
                                    f"마지막 확인: {fmt(e.get('checked_at', '')) or '-'}\nAmazon 마지막 업데이트: {fmt(e.get('amazon_updated', '')) or '-'}\n\n"
                                    f"원래 값:\n{e.get('old', '')}\n\n새 값:\n{e.get('new', '')}\n\n"
                                    f"Amazon 메시지:\n{e.get('issues', '') or '-'}\n\nAPI 응답: {e.get('result', '')}", parent=h)
        tv.bind("<Double-1>", detail)
        b = ttk.Frame(f)
        b.pack(fill="x")
        ttk.Button(b, text="반영 확인", style="Run.TButton", command=do_verify).pack(side="left")
        ttk.Button(b, text="새로 고침", command=load).pack(side="left", padx=6)
        ttk.Button(b, text="닫기", command=h.destroy).pack(side="right")
        ttk.Button(b, text="선택한 기록 되돌리기", style="Stop.TButton", command=do_revert).pack(side="right", padx=6)
