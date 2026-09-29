"""
条形码生成工具（独立版）：读取 Excel 指定列，把每个值生成条形码图片导出到指定目录。
依赖：openpyxl, python-barcode, Pillow, tkinter
"""
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import barcode_tools

BARCODE_TYPES = ["code128", "code39", "ean13", "ean8", "upca"]


def _selftest(excel_path: str, out_dir: str = None, header_row: int = 2):
    """用真实的打包代码跑一遍，验证客户端能否生成条形码，并把过程写入日志。"""
    import datetime
    log_lines = []
    try:
        if not os.path.exists(excel_path):
            log_lines.append(f"[ERROR] 文件不存在: {excel_path}")
            raise SystemExit("\n".join(log_lines))
        headers = barcode_tools.get_headers(excel_path, header_row)
        log_lines.append(f"[INFO] 列名(第{header_row}行): {headers}")
        if not headers:
            log_lines.append("[ERROR] 未读取到列名，请检查列名行设置。")
            raise SystemExit("\n".join(log_lines))

        # 优先选“物品编码”，否则选第一个全为ASCII(可编码)且非空的列
        target = None
        if "物品编码" in headers:
            target = "物品编码"
        else:
            for h in headers:
                idx = headers.index(h)
                info = barcode_tools.analyze_column(excel_path, idx, header_row + 1)
                if info["non_empty"] > 0 and info["non_ascii"] == 0:
                    target = h
                    break
        if not target:
            log_lines.append("[ERROR] 未找到可编码列（一维条码需ASCII/数字，中文列会被跳过）。")
            raise SystemExit("\n".join(log_lines))

        idx = headers.index(target)
        info = barcode_tools.analyze_column(excel_path, idx, header_row + 1)
        log_lines.append(
            f"[INFO] 选用列【{target}】(索引{idx})，非空 {info['non_empty']} 个，其中中文/非ASCII {info['non_ascii']} 个"
        )

        if out_dir is None:
            out_dir = os.path.join(os.path.dirname(excel_path) or ".", "selftest_条形码输出")
        res = barcode_tools.generate_barcodes(
            excel_path, idx, out_dir, "code128", True, header_row + 1
        )
        log_lines.append(
            f"[RESULT] 成功 {res['count']} 张，跳过 {len(res['skipped'])} 张"
        )
        if res["skipped"]:
            for s in res["skipped"][:10]:
                log_lines.append(f"   跳过 行{s[0]}: {s[1]!r} -> {s[2]}")
        log_lines.append(f"[DONE] 输出目录: {out_dir}")
    except SystemExit:
        pass
    except Exception as e:  # noqa: BLE001
        log_lines.append(f"[EXCEPTION] {type(e).__name__}: {e}")

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(out_dir or os.getcwd(), f"selftest_log_{ts}.txt")
    try:
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log_lines) + "\n")
    except Exception:  # noqa: BLE001
        print("\n".join(log_lines))
    return log_lines, log_path


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 指定列生成条形码工具")
        self.geometry("560x420")
        self.resizable(True, True)
        ttk.Style().configure("TButton", padding=4)

        ttk.Label(
            self,
            text="选择 Excel 文件 → 读取列名并选择条形码列 → 选择导出目录 → 生成。\n"
                 "（列名行/数据起始行可调：默认列名在第 2 行、数据从第 3 行开始）",
            wraplength=520, justify="left",
        ).pack(padx=12, pady=10, anchor="w")

        lf = ttk.LabelFrame(self, text="条形码生成设置")
        lf.pack(fill="x", padx=10, pady=6)

        # 表格文件
        f = ttk.Frame(lf); f.pack(fill="x", padx=10, pady=4)
        ttk.Label(f, text="Excel 文件", width=12).pack(side="left")
        self.c_path = tk.StringVar()
        ttk.Entry(f, textvariable=self.c_path).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(f, text="浏览…", command=self._browse_file).pack(side="left", padx=2)

        # 列选择
        col_f = ttk.Frame(lf); col_f.pack(fill="x", padx=10, pady=4)
        ttk.Label(col_f, text="条形码列", width=12).pack(side="left")
        self.bc_col = ttk.Combobox(col_f, state="readonly", width=24)
        self.bc_col.pack(side="left", padx=4)
        ttk.Button(col_f, text="读取列名", command=self._load_cols).pack(side="left", padx=2)

        # 选项
        opt_f = ttk.Frame(lf); opt_f.pack(fill="x", padx=10, pady=4)
        ttk.Label(opt_f, text="条码类型", width=12).pack(side="left")
        self.bc_type = ttk.Combobox(opt_f, state="readonly", width=10, values=BARCODE_TYPES)
        self.bc_type.current(0); self.bc_type.pack(side="left", padx=4)
        self.bc_text = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_f, text="显示文字", variable=self.bc_text).pack(side="left", padx=6)

        # 列名行 / 数据起始行
        row_f = ttk.Frame(lf); row_f.pack(fill="x", padx=10, pady=4)
        ttk.Label(row_f, text="列名行", width=12).pack(side="left")
        self.bc_header_row = tk.IntVar(value=2)
        hspin = ttk.Spinbox(
            row_f, from_=1, to=100, width=5, textvariable=self.bc_header_row, increment=1,
        )
        hspin.pack(side="left", padx=2)
        ttk.Label(row_f, text="数据起始行").pack(side="left", padx=(8, 2))
        self.bc_start_row = tk.IntVar(value=3)
        ttk.Spinbox(
            row_f, from_=1, to=100, width=5, textvariable=self.bc_start_row, increment=1,
        ).pack(side="left", padx=2)

        # 列名行变化时，数据起始行自动跟随（列名行的下一行）
        def _sync_start(*_):
            try:
                hv = self.bc_header_row.get()
                if hv >= 1:
                    self.bc_start_row.set(hv + 1)
            except Exception:  # noqa: BLE001
                pass

        hspin.configure(command=_sync_start)
        self.bc_header_row.trace_add("write", _sync_start)

        # 导出目录
        d = ttk.Frame(lf); d.pack(fill="x", padx=10, pady=4)
        ttk.Label(d, text="导出目录", width=12).pack(side="left")
        self.bc_dir = tk.StringVar()
        ttk.Entry(d, textvariable=self.bc_dir).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(d, text="选择目录…", command=self._browse_dir).pack(side="left", padx=2)

        # 操作行
        act = ttk.Frame(lf); act.pack(fill="x", padx=10, pady=8)
        ttk.Button(act, text="▶ 生成条形码", command=self._do_barcode).pack(side="right", padx=4)
        self.open_btn = ttk.Button(act, text="打开导出目录", command=self._open_dir, state="disabled")
        self.open_btn.pack(side="right", padx=4)

        self.status = tk.StringVar(value="待操作")
        ttk.Label(self, textvariable=self.status, foreground="#1a7f37").pack(anchor="w", padx=14, pady=6)

    # ---------------- 辅助 ----------------
    def _browse_file(self):
        p = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if p:
            self.c_path.set(p)
            # 选完文件自动读取列名，避免忘记点“读取列名”导致无法生成
            self._load_cols()

    def _browse_dir(self):
        p = filedialog.askdirectory(title="选择条形码导出目录")
        if p:
            self.bc_dir.set(p)

    def _load_cols(self):
        p = self.c_path.get()
        if not p or not os.path.exists(p):
            messagebox.showwarning("提示", "请先选择 Excel 文件。")
            return
        try:
            headers = barcode_tools.get_headers(p, self.bc_header_row.get())
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("错误", f"读取表头失败：{e}")
            return
        self.bc_col["values"] = headers
        if headers:
            self.bc_col.current(0)
        self.status.set(f"已读取 {len(headers)} 个列，请选择条形码列。")

    def _open_dir(self):
        p = self.bc_dir.get()
        if p and os.path.isdir(p):
            try:
                os.startfile(p)
            except Exception:  # noqa: BLE001
                subprocess.Popen(["explorer", p])

    def _run_threaded(self, func, *args):
        self.status.set("处理中，请稍候…")
        self.open_btn.configure(state="disabled")

        def worker():
            try:
                result = func(*args)
                self.after(0, lambda: self._on_done(result, None))
            except Exception as e:  # noqa: BLE001
                self.after(0, lambda: self._on_done(None, str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_done(self, result, error):
        if error:
            self.status.set("❌ 出错：" + error)
            messagebox.showerror("错误", error)
            return
        msg = f"✅ 生成完成：成功 {result['count']} 张"
        skipped = result.get("skipped") or []
        if skipped:
            msg += f"，跳过 {len(skipped)} 张"
            # 若跳过是因中文/非ASCII，给出明确指引
            if any("not valid" in (s[2] if len(s) > 2 else "") or "Code" in (s[2] if len(s) > 2 else "") for s in skipped):
                msg += (
                    "\n\n⚠️ 被跳过的值含中文/非ASCII字符，一维条形码无法编码中文。"
                    "\n请改用数字编码列（如“物品编码”），或改用二维码(QR)。"
                )
        msg += f"\n导出目录：{result['output_dir']}"
        self.status.set(msg)
        messagebox.showinfo("完成", msg)
        print("[RESULT]", result)
        self.open_btn.configure(state="normal")

    # ---------------- 生成 ----------------
    def _do_barcode(self):
        p = self.c_path.get()
        bdir = self.bc_dir.get()
        idx = self.bc_col.current()
        if not p or not os.path.exists(p):
            messagebox.showwarning("提示", "请选择 Excel 文件。")
            return
        if idx < 0:
            messagebox.showwarning("提示", "请先读取并选择条形码列。")
            return
        if not bdir:
            messagebox.showwarning("提示", "请选择条形码导出目录。")
            return
        btype = self.bc_type.get() or "code128"
        start_row = self.bc_start_row.get()

        # 预检：所选列是否含中文/非ASCII（一维条码无法编码）
        try:
            info = barcode_tools.analyze_column(p, idx, start_row)
        except Exception:  # noqa: BLE001
            info = {"non_empty": 0, "non_ascii": 0}
        if info["non_empty"] == 0:
            messagebox.showwarning(
                "提示",
                f"所选列从第 {start_row} 行起没有可读数据。\n"
                "请检查“数据起始行”是否设置正确（列名在第 2 行时，数据起始行应为 3）。",
            )
            return
        if info["non_ascii"] == info["non_empty"]:
            ans = messagebox.askyesno(
                "可能无法生成条形码",
                f"所选列“{self.bc_col.get()}”的全部 {info['non_empty']} 个值都含中文/非ASCII字符，\n"
                "一维条形码（Code128 等）无法编码中文，将会全部跳过、生成 0 张。\n\n"
                "建议改用数字编码列（如“物品编码”）。\n\n"
                "仍要尝试生成吗？",
            )
            if not ans:
                return
        elif info["non_ascii"] > 0:
            messagebox.showwarning(
                "注意",
                f"所选列有 {info['non_ascii']}/{info['non_empty']} 个值含中文/非ASCII字符，"
                "这些会被跳过（一维条码无法编码中文）。",
            )

        self._run_threaded(
            barcode_tools.generate_barcodes,
            p, idx, bdir, btype, self.bc_text.get(), start_row,
        )


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        # 用法: BarcodeTool.exe --selftest <excel路径> [输出目录] [列名行]
        args = sys.argv[1:]
        excel = None
        out = None
        hrow = 2
        rest = [a for a in args if a != "--selftest"]
        if rest:
            excel = rest[0]
        if len(rest) > 1:
            out = rest[1]
        if len(rest) > 2:
            try:
                hrow = int(rest[2])
            except ValueError:
                pass
        if not excel:
            print("用法: BarcodeTool.exe --selftest <excel路径> [输出目录] [列名行]")
            sys.exit(1)
        lines, log_path = _selftest(excel, out, hrow)
        print("\n".join(lines))
        print(f"日志已写入: {log_path}")
        sys.exit(0)
    App().mainloop()
