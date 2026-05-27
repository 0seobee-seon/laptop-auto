"""
야근수당내역 자동 생성기 (GUI v2.2)
  - 여러 파일 배치 처리
  - 드래그앤드롭 (tkinterdnd2)
  - append 모드: 하나의 기존 파일에 여러 1번 파일을 부서명 시트로 일괄 추가

빌드:
  pip install tkinterdnd2
  pyinstaller --onefile --windowed --name "야근수당_생성기" ^
    --add-data "야근수당_자동입력.py;." ^
    --hidden-import tkinterdnd2 ^
    --hidden-import openpyxl --hidden-import xlrd --hidden-import xlwt ^
    --hidden-import xlutils.copy --hidden-import pandas ^
    --hidden-import win32com.client ^
    야근수당_생성기_GUI.py
"""

import os
import sys
import threading
import traceback
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# core 모듈 로드 (PyInstaller frozen 대응)
if getattr(sys, 'frozen', False):
    BUNDLE_DIR = sys._MEIPASS                              # noqa
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    SCRIPT_DIR = BUNDLE_DIR
sys.path.insert(0, BUNDLE_DIR)

import importlib.util
_spec = importlib.util.spec_from_file_location(
    "core", os.path.join(BUNDLE_DIR, "야근수당_자동입력.py")
)
core = importlib.util.module_from_spec(_spec)
core.__name__ = "core"
_spec.loader.exec_module(core)


# ══════════════════════════════════════════════════
#  데이터 클래스
# ══════════════════════════════════════════════════

class FileEntry:
    def __init__(self, path, dept=""):
        self.path   = path
        self.dept   = dept
        self.status = "대기"


# ══════════════════════════════════════════════════
#  GUI 메인 클래스
# ══════════════════════════════════════════════════

class OvertimeApp:
    def __init__(self, root):
        self.root = root
        root.title("야근수당내역 자동 생성기  v2.2")
        root.geometry("840x820")
        root.resizable(True, True)
        root.minsize(720, 650)

        self.fn_label = ("맑은 고딕", 10)
        self.fn_entry = ("맑은 고딕", 10)
        self.fn_btn   = ("맑은 고딕", 10, "bold")

        self.entries = []   # list[FileEntry]

        self.var_year   = tk.StringVar(value=str(datetime.date.today().year))
        self.var_month  = tk.StringVar(value=f"{datetime.date.today().month:02d}")
        self.var_mode   = tk.StringVar(value="new")
        self.var_target = tk.StringVar()   # 기존 누적 파일 (append 모드 공통)
        self.var_outdir = tk.StringVar()

        self._build_ui()
        self._setup_dnd()

    # ── UI 구성 ───────────────────────────────────
    def _build_ui(self):
        p = {"padx": 12, "pady": 5}

        # 제목
        tk.Label(self.root, text="야근수당내역 자동 생성기",
                 font=("맑은 고딕", 16, "bold"), fg="#1f4788").pack(pady=(14, 2))
        if HAS_DND:
            hint = "✔ 드래그앤드롭 지원  |  파일 목록: 창 어디서나  |  기존 파일: 해당 입력칸에 드래그"
            hint_fg = "#1a7a1a"
        else:
            hint = "1번 야근현황 파일을 추가하세요. (드래그앤드롭: tkinterdnd2 설치 필요)"
            hint_fg = "#888"
        tk.Label(self.root, text=hint, font=("맑은 고딕", 9), fg=hint_fg).pack()

        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=12, pady=7)

        # ── ① 파일 목록 ──────────────────────────
        frm1 = tk.LabelFrame(
            self.root,
            text="  ① 처리할 1번 야근현황 파일 목록  (부서명·파일명 더블클릭으로 수정)",
            font=self.fn_btn, padx=8, pady=6)
        frm1.pack(fill="both", expand=True, **p)

        cols = ("no", "file", "dept", "status")
        self.tree = ttk.Treeview(frm1, columns=cols, show="headings", height=8)
        self.tree.heading("no",     text="#",      anchor="center")
        self.tree.heading("file",   text="파일명",  anchor="w")
        self.tree.heading("dept",   text="부서명",  anchor="center")
        self.tree.heading("status", text="상태",   anchor="center")
        self.tree.column("no",     width=32,  stretch=False, anchor="center")
        self.tree.column("file",   width=430, stretch=True,  anchor="w")
        self.tree.column("dept",   width=120, stretch=False, anchor="center")
        self.tree.column("status", width=65,  stretch=False, anchor="center")

        vsb = ttk.Scrollbar(frm1, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._tree_dbl)

        # 파일 관리 버튼
        bf = tk.Frame(self.root)
        bf.pack(fill="x", padx=12, pady=(2, 4))
        tk.Button(bf, text="+ 파일 추가", font=self.fn_btn, width=12,
                  command=self._add_files, bg="#e8e8e8").pack(side="left", padx=(0, 6))
        tk.Button(bf, text="선택 삭제",   font=self.fn_btn, width=10,
                  command=self._remove_sel, bg="#e8e8e8").pack(side="left", padx=(0, 6))
        tk.Button(bf, text="전체 지우기", font=self.fn_btn, width=10,
                  command=self._clear_all, bg="#e8e8e8").pack(side="left")

        # ── ② 공통 설정 ──────────────────────────
        frm2 = tk.LabelFrame(self.root, text="  ② 공통 설정 (연도 / 월)  ",
                              font=self.fn_btn, padx=10, pady=7)
        frm2.pack(fill="x", **p)

        tk.Label(frm2, text="연도:", font=self.fn_label).grid(row=0, column=0, sticky="w")
        ttk.Combobox(frm2, textvariable=self.var_year,
                     values=[str(y) for y in range(2020, 2031)],
                     width=7, state="readonly").grid(row=0, column=1, padx=(4, 20))
        tk.Label(frm2, text="월:", font=self.fn_label).grid(row=0, column=2, sticky="w")
        ttk.Combobox(frm2, textvariable=self.var_month,
                     values=[f"{m:02d}" for m in range(1, 13)],
                     width=5, state="readonly").grid(row=0, column=3, padx=(4, 0))

        # ── ③ 출력 방식 ──────────────────────────
        frm3 = tk.LabelFrame(self.root, text="  ③ 출력 방식 / 저장 위치  ",
                              font=self.fn_btn, padx=10, pady=7)
        frm3.pack(fill="x", **p)
        frm3.columnconfigure(1, weight=1)

        tk.Radiobutton(frm3, text="각 파일별 새 xlsx 생성",
                       variable=self.var_mode, value="new",
                       font=self.fn_label, command=self._toggle_mode)\
            .grid(row=0, column=0, sticky="w", padx=(0, 20))
        tk.Radiobutton(frm3,
                       text="기존 누적 파일에 시트 추가  (부서명으로 시트 생성, 여러 파일 동시 처리 가능)",
                       variable=self.var_mode, value="append",
                       font=self.fn_label, command=self._toggle_mode)\
            .grid(row=0, column=1, sticky="w", columnspan=2)

        # 저장 위치
        tk.Label(frm3, text="저장 위치:", font=self.fn_label)\
            .grid(row=1, column=0, sticky="w", pady=(6, 0))
        tk.Entry(frm3, textvariable=self.var_outdir, font=self.fn_entry)\
            .grid(row=1, column=1, sticky="ew", pady=(6, 0), padx=(0, 6))
        tk.Button(frm3, text="저장 위치...", font=self.fn_btn,
                  command=self._pick_outdir, width=10)\
            .grid(row=1, column=2, sticky="w", pady=(6, 0))

        # 기존 누적 파일 (append 전용)
        self.lbl_target = tk.Label(frm3,
                                    text="기존 파일: (드래그 가능)",
                                    font=self.fn_label)
        self.lbl_target.grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.ent_target = tk.Entry(frm3, textvariable=self.var_target, font=self.fn_entry)
        self.ent_target.grid(row=2, column=1, sticky="ew", pady=(6, 0), padx=(0, 6))
        self.btn_target = tk.Button(frm3, text="찾아보기", font=self.fn_btn,
                                     command=self._pick_target, width=10)
        self.btn_target.grid(row=2, column=2, sticky="w", pady=(6, 0))

        # append 모드 설명 라벨
        self.lbl_append_info = tk.Label(
            frm3,
            text="  ※ 위 목록의 각 파일이 '부서명' 시트로 기존 파일에 추가됩니다.",
            font=("맑은 고딕", 9), fg="#1f4788")
        self.lbl_append_info.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))

        self._toggle_mode()

        # ── ④ 진행 상황 ──────────────────────────
        frm4 = tk.LabelFrame(self.root, text="  ④ 진행 상황  ",
                              font=self.fn_btn, padx=8, pady=4)
        frm4.pack(fill="x", padx=12, pady=5)
        self.log = scrolledtext.ScrolledText(frm4, height=9, font=("Consolas", 9),
                                              bg="#f8f8f8", fg="#222")
        self.log.pack(fill="both", expand=True)

        # ── ⑤ 실행 버튼 ──────────────────────────
        frm5 = tk.Frame(self.root)
        frm5.pack(fill="x", padx=12, pady=8)
        self.btn_run = tk.Button(frm5, text="▶  전체 실행",
                                  font=("맑은 고딕", 11, "bold"),
                                  bg="#1f4788", fg="white",
                                  command=self._run, height=2, width=16)
        self.btn_run.pack(side="left", padx=(0, 10))
        tk.Button(frm5, text="로그 지우기", font=self.fn_btn,
                  command=lambda: self.log.delete("1.0", "end"),
                  width=12, height=2).pack(side="left")
        tk.Button(frm5, text="종료", font=self.fn_btn,
                  command=self.root.quit, width=10, height=2).pack(side="right")

    # ── 드래그앤드롭 ──────────────────────────────
    def _setup_dnd(self):
        if not HAS_DND:
            return
        # 창 전체 → 1번 파일 목록에 추가
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self._on_drop)
        # 기존 파일 입력란 → var_target 설정 + append 모드 자동 전환
        for w in (self.ent_target, self.lbl_target):
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", self._on_drop_target)

    def _on_drop(self, event):
        for p in self._parse_dnd(event.data):
            if p.lower().endswith((".xlsx", ".xls")):
                self._add_file(p)

    def _on_drop_target(self, event):
        paths = self._parse_dnd(event.data)
        if not paths:
            return
        p = paths[0]
        if not p.lower().endswith((".xlsx", ".xls")):
            return
        self.var_target.set(p)
        self.var_mode.set("append")
        self._toggle_mode()
        if not self.var_outdir.get():
            self.var_outdir.set(os.path.dirname(p))

    @staticmethod
    def _parse_dnd(raw):
        raw = raw.strip()
        paths, i = [], 0
        while i < len(raw):
            if raw[i] == '{':
                j = raw.index('}', i)
                paths.append(raw[i + 1:j])
                i = j + 2
            elif raw[i] == ' ':
                i += 1
            else:
                j = raw.find(' ', i)
                if j == -1:
                    paths.append(raw[i:])
                    break
                paths.append(raw[i:j])
                i = j
        return [p for p in paths if p]

    # ── 파일 목록 관리 ────────────────────────────
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="1번 야근현황 파일 선택 (여러 개 가능)",
            filetypes=[("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")])
        for p in paths:
            self._add_file(p)

    def _add_file(self, path):
        if any(e.path == path for e in self.entries):
            return
        entry = FileEntry(path, self._guess_dept(path))
        self.entries.append(entry)
        self._refresh()
        if len(self.entries) == 1 and not self.var_outdir.get():
            self.var_outdir.set(os.path.dirname(path))

    @staticmethod
    def _guess_dept(path):
        name = os.path.splitext(os.path.basename(path))[0]
        for kw in ["안전진단팀", "서울설계팀", "설계1본부", "설계2본부",
                   "설계본부", "구조팀", "조경팀", "토목팀", "건축팀"]:
            if kw in name:
                return kw
        return ""

    def _remove_sel(self):
        selected = self.tree.selection()
        if not selected:
            return
        for idx in sorted([int(s) for s in selected], reverse=True):
            self.entries.pop(idx)
        self._refresh()

    def _clear_all(self):
        self.entries.clear()
        self._refresh()

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for i, e in enumerate(self.entries):
            tag = {"완료": "ok", "오류": "err"}.get(e.status, "")
            self.tree.insert("", "end", iid=str(i),
                             values=(i + 1, os.path.basename(e.path), e.dept, e.status),
                             tags=(tag,))
        self.tree.tag_configure("ok",  foreground="#1a7a1a")
        self.tree.tag_configure("err", foreground="#cc0000")

    # ── Treeview 더블클릭 편집 ────────────────────
    def _tree_dbl(self, event):
        item = self.tree.identify_row(event.y)
        col  = self.tree.identify_column(event.x)
        if not item or self.tree.identify_region(event.x, event.y) != "cell":
            return
        idx = int(item)
        if col == "#3":
            self._edit_dept(idx)
        elif col == "#2":
            self._replace_file(idx)

    def _edit_dept(self, idx):
        entry = self.entries[idx]
        dlg = tk.Toplevel(self.root)
        dlg.title("부서명 수정")
        dlg.geometry("310x115")
        dlg.resizable(False, False)
        dlg.grab_set()
        tk.Label(dlg, text=os.path.basename(entry.path),
                 font=("맑은 고딕", 9), fg="#666").pack(pady=(10, 4))
        var = tk.StringVar(value=entry.dept)
        ent = tk.Entry(dlg, textvariable=var, font=("맑은 고딕", 10), width=27)
        ent.pack()
        ent.select_range(0, "end")
        ent.focus()

        def ok():
            entry.dept = var.get().strip()
            self._refresh()
            dlg.destroy()

        ent.bind("<Return>", lambda _: ok())
        tk.Button(dlg, text="확인", command=ok,
                  font=("맑은 고딕", 10, "bold"), width=10).pack(pady=8)

    def _replace_file(self, idx):
        p = filedialog.askopenfilename(
            title="파일 교체",
            filetypes=[("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")])
        if p:
            self.entries[idx].path = p
            self._refresh()

    # ── 모드 토글 ─────────────────────────────────
    def _toggle_mode(self):
        if self.var_mode.get() == "append":
            self.ent_target.config(state="normal")
            self.btn_target.config(state="normal")
            self.lbl_append_info.grid()
        else:
            self.ent_target.config(state="disabled")
            self.btn_target.config(state="disabled")
            self.lbl_append_info.grid_remove()

    def _pick_outdir(self):
        d = filedialog.askdirectory(title="저장 위치 선택")
        if d:
            self.var_outdir.set(d)

    def _pick_target(self):
        p = filedialog.askopenfilename(
            title="기존 누적 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")])
        if p:
            self.var_target.set(p)
            if not self.var_outdir.get():
                self.var_outdir.set(os.path.dirname(p))

    # ── 로그 ──────────────────────────────────────
    def _log(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    # ── 실행 ──────────────────────────────────────
    def _run(self):
        if not self.entries:
            messagebox.showwarning("입력 오류", "처리할 파일을 먼저 추가해주세요.")
            return

        year   = self.var_year.get()
        month  = self.var_month.get()
        mode   = self.var_mode.get()
        outdir = self.var_outdir.get().strip()
        target = self.var_target.get().strip()

        if not outdir:
            messagebox.showwarning("입력 오류", "저장 위치를 지정해주세요.")
            return

        if mode == "append" and (not target or not os.path.exists(target)):
            messagebox.showwarning("입력 오류", "기존 누적 파일을 선택해주세요.")
            return

        # append 모드: 부서명 필수 확인
        if mode == "append":
            missing = [e for e in self.entries if not e.dept]
            if missing:
                names = "\n".join(f"  · {os.path.basename(e.path)}" for e in missing)
                messagebox.showwarning("부서명 미입력",
                    f"시트 이름으로 사용할 부서명이 비어 있습니다:\n{names}\n\n"
                    "목록에서 부서명을 더블클릭해 입력해주세요.")
                return

        missing_dept = [e for e in self.entries if not e.dept] if mode == "new" else []
        if missing_dept:
            names = "\n".join(f"  · {os.path.basename(e.path)}" for e in missing_dept)
            if not messagebox.askyesno("부서명 미입력",
                    f"부서명이 비어 있는 파일이 있습니다:\n{names}\n\n그대로 계속하시겠습니까?"):
                return

        self.btn_run.config(state="disabled", text="실행 중...")
        for e in self.entries:
            e.status = "대기"
        self._refresh()

        threading.Thread(
            target=self._do_batch,
            args=(year, month, mode, outdir, target),
            daemon=True
        ).start()

    def _do_batch(self, year, month, mode, outdir, target):
        total = len(self.entries)
        success = 0

        self._log("=" * 55)
        if mode == "append":
            self._log(f"배치 처리 시작  총 {total}개 파일  →  {os.path.basename(target)}")
            self._log(f"시트명: 각 파일의 부서명으로 생성")
        else:
            self._log(f"배치 처리 시작  총 {total}개 파일  ({year}.{month}월)")
        self._log("=" * 55)

        # append 모드: 기존 파일을 저장 위치로 복사 (원본 보존)
        if mode == "append":
            import shutil
            ext = os.path.splitext(target)[1].lower()
            xlsx_target = target
            if ext == ".xls":
                self._log("기존 파일 xls → xlsx 변환 중...")
                xlsx_target = core.convert_xls_to_xlsx(target)
                self._log(f"변환 완료: {os.path.basename(xlsx_target)}")
            out_path = os.path.join(outdir, os.path.basename(xlsx_target))
            if not out_path.endswith(".xlsx"):
                out_path = os.path.splitext(out_path)[0] + ".xlsx"
            if os.path.abspath(xlsx_target) != os.path.abspath(out_path):
                shutil.copy(xlsx_target, out_path)
                self._log(f"복사 완료: {os.path.basename(out_path)}")
        else:
            out_path = None  # new 모드에서는 파일별로 생성

        for i, entry in enumerate(self.entries):
            self._log(f"\n[{i + 1}/{total}] {os.path.basename(entry.path)}")
            self._log(f"    부서: {entry.dept or '(미입력)'}")
            entry.status = "처리 중"
            self._refresh()
            try:
                self._process_one(entry, year, month, mode, outdir, out_path)
                entry.status = "완료"
                success += 1
            except Exception as exc:
                entry.status = "오류"
                self._log(f"  [오류] {type(exc).__name__}: {exc}")
                self._log(traceback.format_exc())
            self._refresh()

        self._log("\n" + "=" * 55)
        self._log(f"  ✓ 배치 완료  {success}/{total} 성공")
        if mode == "append":
            self._log(f"  저장 파일: {out_path}")
        else:
            self._log(f"  저장 폴더: {outdir}")
        self._log("=" * 55)

        save_loc = out_path if mode == "append" else outdir
        msg = f"배치 처리 완료!\n\n성공: {success}/{total}"
        if success < total:
            msg += f"\n실패: {total - success}개 (로그 참조)"
        msg += f"\n\n저장 위치:\n{save_loc}"
        messagebox.showinfo("완료", msg)

        if messagebox.askyesno("파일/폴더 열기", "저장 위치를 여시겠습니까?"):
            os.startfile(save_loc if mode == "append" else outdir)

        self.btn_run.config(state="normal", text="▶  전체 실행")

    def _process_one(self, entry, year, month, mode, outdir, out_path):
        import pandas as pd
        xf = pd.ExcelFile(entry.path)
        data_sheets = [s for s in xf.sheet_names if s not in core.SKIP_SHEETS]
        core.DEPT_MAP     = {s: entry.dept for s in data_sheets}
        core.SPECIAL_DEPT = {}

        emp = core.extract_overtime(entry.path)
        self._log(f"    추출: {len(emp)}명")

        if mode == "new":
            sheet_name = f"{year}.{month}월"
            fname  = f"{entry.dept or '야근수당내역'}_{sheet_name.replace('.', '_')}.xlsx"
            output = os.path.join(outdir, fname)
            core.create_new(emp, sheet_name, output)
            self._log(f"    → {fname}")

        else:  # append: 시트명 = 부서명, E2 = 기간 라벨, 위치 = 목록 순서대로 맨 앞
            sheet_name   = entry.dept
            batch_idx    = self.entries.index(entry)   # 0, 1, 2, ...
            period_label = core._make_period_label(f"{year}.{month}월")
            updated, skipped = core.update_xlsx_workbook(
                out_path, emp, sheet_name,
                output_path=None, template_sheet=None, auto_create=True,
                move_to_index=batch_idx, period_label=period_label)
            self._log(f"    시트 '{sheet_name}' 추가 (기간: {period_label})")
            self._log(f"    위치: 맨 앞에서 {batch_idx + 1}번째 / 업데이트: {len(updated)}명")
            if skipped:
                self._log(f"    없는 직원(0 유지): {', '.join(skipped)}")


# ══════════════════════════════════════════════════
#  진입점
# ══════════════════════════════════════════════════

if __name__ == "__main__":
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    OvertimeApp(root)
    root.mainloop()
