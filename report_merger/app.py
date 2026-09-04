from __future__ import annotations

import ctypes
import json
import os
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText

from .core import ReportModel
from .exporters import export_docx, export_xlsx


SUPPORTED = {".docx", ".pdf"}


def resource_path(*parts: str) -> Path:
    """Return a bundled resource path in both source and PyInstaller builds."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base.joinpath(*parts)


class WindowsDropTarget:
    WM_DROPFILES = 0x0233
    GWLP_WNDPROC = -4

    def __init__(self, root: tk.Tk, callback):
        self.root = root
        self.callback = callback
        self.enabled = False
        self._new_proc = None
        self._old_proc = None
        if sys.platform != "win32":
            return
        try:
            root.update_idletasks()
            hwnd = root.winfo_id()
            user32 = ctypes.windll.user32
            shell32 = ctypes.windll.shell32
            lresult = ctypes.c_ssize_t
            wndproc_type = ctypes.WINFUNCTYPE(
                lresult, ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t
            )
            set_long = user32.SetWindowLongPtrW
            set_long.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
            set_long.restype = ctypes.c_void_p
            call_proc = user32.CallWindowProcW
            call_proc.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint,
                                  ctypes.c_size_t, ctypes.c_ssize_t]
            call_proc.restype = lresult

            def window_proc(window, message, wparam, lparam):
                if message == self.WM_DROPFILES:
                    count = shell32.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)
                    paths = []
                    for index in range(count):
                        length = shell32.DragQueryFileW(wparam, index, None, 0)
                        buffer = ctypes.create_unicode_buffer(length + 1)
                        shell32.DragQueryFileW(wparam, index, buffer, length + 1)
                        paths.append(buffer.value)
                    shell32.DragFinish(wparam)
                    root.after(0, lambda: callback(paths))
                    return 0
                return call_proc(self._old_proc, window, message, wparam, lparam)

            self._new_proc = wndproc_type(window_proc)
            self._old_proc = set_long(hwnd, self.GWLP_WNDPROC, self._new_proc)
            shell32.DragAcceptFiles(hwnd, True)
            self.enabled = True
        except Exception:
            self.enabled = False


class ReportMergerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("BoctachvaGopBaoCao")
        self.root.geometry("1420x820")
        self.root.minsize(1120, 680)
        self.model = ReportModel()
        self.files: list[dict[str, str]] = []
        self.tree_to_group: dict[str, str] = {}
        self.group_to_tree: dict[str, str] = {}
        self.support_window: tk.Toplevel | None = None
        self.support_qr: tk.PhotoImage | None = None
        self.guide_window: tk.Toplevel | None = None
        self.status = tk.StringVar(value="Sẵn sàng. Hãy thêm các báo cáo Word hoặc PDF dạng text.")
        self.output_title = tk.StringVar(value="BÁO CÁO TỔNG HỢP")
        self._configure_style()
        self._build_ui()
        self.drop_target = WindowsDropTarget(self.root, self.add_paths)
        if self.drop_target.enabled:
            self.drop_hint.configure(text="Kéo thả .docx/.pdf vào cửa sổ hoặc bấm “Thêm file”")

    def _configure_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#17324D")
        style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#5B6573")
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Support.TButton", font=("Segoe UI", 10, "bold"), foreground="#0B6B61")
        style.configure("SupportTitle.TLabel", font=("Segoe UI", 20, "bold"), foreground="#0B6B61")
        style.configure("SupportText.TLabel", font=("Segoe UI", 10), foreground="#344054")
        style.configure("SupportLink.TLabel", font=("Segoe UI", 10, "underline"), foreground="#1267A5")
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=(16, 12, 16, 8))
        top.pack(fill="x")
        ttk.Label(top, text="BOCTACHVAGOPBAOCAO", style="Title.TLabel").pack(side="left")
        self.drop_hint = ttk.Label(top, text="Bấm “Thêm file” để chọn .docx/.pdf", style="Sub.TLabel")
        self.drop_hint.pack(side="right", pady=(8, 0))

        toolbar = ttk.Frame(self.root, padding=(16, 0, 16, 10))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="+ Thêm file", command=self.choose_files).pack(side="left")
        ttk.Button(toolbar, text="Xóa file", command=self.remove_file).pack(side="left", padx=5)
        ttk.Button(toolbar, text="Đổi tên nguồn", command=self.rename_source).pack(side="left")
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Button(toolbar, text="Phân tích cấu trúc", style="Accent.TButton",
                   command=self.analyze).pack(side="left")
        ttk.Button(toolbar, text="Nạp quy đổi", command=self.load_aliases).pack(side="left", padx=(10, 5))
        ttk.Button(toolbar, text="Lưu quy đổi", command=self.save_aliases).pack(side="left")
        tk.Button(
            toolbar,
            text="HƯỚNG DẪN SỬ DỤNG",
            command=self.show_guide,
            foreground="#C81E1E",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            padx=8,
            pady=3,
        ).pack(side="right", padx=(0, 6))
        ttk.Button(toolbar, text="Ủng hộ tác giả", style="Support.TButton",
                   command=self.show_support).pack(side="right")

        paned = ttk.Panedwindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=16)

        left = ttk.Labelframe(paned, text="1. Danh sách báo cáo", padding=8)
        paned.add(left, weight=1)
        self.file_list = tk.Listbox(left, selectmode="browse", font=("Segoe UI", 9),
                                    activestyle="none", borderwidth=0, highlightthickness=0)
        self.file_list.pack(fill="both", expand=True)
        self.file_list.bind("<Double-1>", lambda _event: self.rename_source())
        ttk.Label(left, text="Tên nguồn sẽ xuất hiện trước nội dung của từng báo cáo.",
                  style="Sub.TLabel", wraplength=260).pack(fill="x", pady=(8, 0))

        middle = ttk.Labelframe(paned, text="2. Cây phần, mục", padding=8)
        paned.add(middle, weight=3)
        tree_frame = ttk.Frame(middle)
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=("level", "sources"),
                                 selectmode="extended", show="tree headings")
        self.tree.heading("#0", text="Tên phần / mục")
        self.tree.heading("level", text="Cấp")
        self.tree.heading("sources", text="Số nguồn")
        self.tree.column("#0", width=500, minwidth=280)
        self.tree.column("level", width=55, anchor="center", stretch=False)
        self.tree.column("sources", width=75, anchor="center", stretch=False)
        scrollbar = ttk.Scrollbar(tree_frame, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_preview)
        self.tree.bind("<Double-1>", lambda _event: self.rename_group())

        editbar = ttk.Frame(middle)
        editbar.pack(fill="x", pady=(8, 0))
        for text, command in (
            ("Đổi tên", self.rename_group), ("Gộp mục", self.merge_groups),
            ("Tách theo nguồn", self.split_group), ("Bỏ/Chọn", self.toggle_group),
            ("↑", lambda: self.move_group(-1)), ("↓", lambda: self.move_group(1)),
        ):
            ttk.Button(editbar, text=text, command=command).pack(side="left", padx=(0, 5))

        right = ttk.Labelframe(paned, text="3. Xem trước nội dung nguyên văn", padding=8)
        paned.add(right, weight=2)
        self.preview = tk.Text(right, wrap="word", font=("Times New Roman", 12),
                               background="#FAFAF8", borderwidth=0, padx=10, pady=10)
        preview_scroll = ttk.Scrollbar(right, command=self.preview.yview)
        self.preview.configure(yscrollcommand=preview_scroll.set)
        self.preview.pack(side="left", fill="both", expand=True)
        preview_scroll.pack(side="right", fill="y")
        self.preview.configure(state="disabled")

        export = ttk.Frame(self.root, padding=(16, 10, 16, 8))
        export.pack(fill="x")
        ttk.Label(export, text="Tiêu đề file tổng hợp:").pack(side="left")
        ttk.Entry(export, textvariable=self.output_title, width=44).pack(side="left", padx=8)
        ttk.Button(export, text="Xuất Word", style="Accent.TButton",
                   command=self.export_word).pack(side="right")
        ttk.Button(export, text="Xuất Excel", command=self.export_excel).pack(side="right", padx=7)

        statusbar = ttk.Label(self.root, textvariable=self.status, relief="sunken",
                              anchor="w", padding=(8, 4))
        statusbar.pack(fill="x")

    def show_guide(self):
        if self.guide_window and self.guide_window.winfo_exists():
            self.guide_window.deiconify()
            self.guide_window.lift()
            self.guide_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        self.guide_window = window
        window.title("Hướng dẫn sử dụng")
        window.geometry("760x650")
        window.minsize(560, 460)
        window.transient(self.root)

        def close_window():
            self.guide_window = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", close_window)
        container = ttk.Frame(window, padding=16)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="HƯỚNG DẪN SỬ DỤNG", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="Nội dung được lấy từ tài liệu README của phần mềm.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 10))

        guide = ScrolledText(
            container,
            wrap="word",
            font=("Segoe UI", 11),
            background="#FAFAF8",
            borderwidth=1,
            relief="solid",
            padx=14,
            pady=12,
        )
        try:
            guide.insert("1.0", resource_path("README.md").read_text(encoding="utf-8"))
        except OSError:
            guide.insert("1.0", "Không thể tải nội dung hướng dẫn.")
        guide.configure(state="disabled")
        guide.pack(fill="both", expand=True)
        ttk.Button(container, text="Đóng", command=close_window).pack(anchor="e", pady=(10, 0))

    def show_support(self):
        if self.support_window and self.support_window.winfo_exists():
            self.support_window.deiconify()
            self.support_window.lift()
            self.support_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        self.support_window = window
        window.title("Ủng hộ tác giả")
        window.geometry("560x760")
        window.minsize(500, 680)
        window.configure(background="#F6FAF9")
        window.transient(self.root)

        def close_window():
            self.support_window = None
            self.support_qr = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", close_window)

        container = ttk.Frame(window, padding=(34, 24))
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="ỦNG HỘ TÁC GIẢ", style="SupportTitle.TLabel").pack()
        ttk.Label(
            container,
            text=("Nếu phần mềm hữu ích với bạn, một lời góp ý, chia sẻ hoặc khoản "
                  "ủng hộ nhỏ đều là động lực để mình tiếp tục hoàn thiện sản phẩm."),
            style="SupportText.TLabel",
            justify="center",
            wraplength=470,
        ).pack(pady=(8, 16))

        qr_path = resource_path("assets", "ung-ho-tac-gia-qr.png")
        try:
            qr = tk.PhotoImage(file=str(qr_path))
            factor = max(1, (qr.width() + 359) // 360)
            if factor > 1:
                qr = qr.subsample(factor, factor)
            self.support_qr = qr
            ttk.Label(container, image=qr).pack(pady=(0, 16))
        except tk.TclError:
            ttk.Label(
                container,
                text="Không thể tải ảnh mã QR ngân hàng.",
                style="SupportText.TLabel",
            ).pack(pady=(24, 32))

        ttk.Label(container, text="Liên hệ và góp ý", style="SupportText.TLabel").pack()
        email = ttk.Label(container, text="manh.hoangquang@gmail.com", style="SupportLink.TLabel",
                          cursor="hand2")
        email.pack(pady=(5, 2))
        email.bind("<Button-1>", lambda _event: webbrowser.open_new_tab(
            "mailto:manh.hoangquang@gmail.com"))
        facebook = ttk.Label(container, text="facebook.com/manhdada", style="SupportLink.TLabel",
                             cursor="hand2")
        facebook.pack()
        facebook.bind("<Button-1>", lambda _event: webbrowser.open_new_tab(
            "https://facebook.com/manhdada"))

        actions = ttk.Frame(container)
        actions.pack(pady=(18, 0))
        ttk.Button(actions, text="Gửi email", command=lambda: webbrowser.open_new_tab(
            "mailto:manh.hoangquang@gmail.com")).pack(side="left", padx=4)
        ttk.Button(actions, text="Mở Facebook", command=lambda: webbrowser.open_new_tab(
            "https://facebook.com/manhdada")).pack(side="left", padx=4)
        ttk.Button(actions, text="Đóng", command=close_window).pack(side="left", padx=4)

    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Chọn các báo cáo", filetypes=[("Word và PDF", "*.docx *.pdf"),
                                                  ("Word", "*.docx"), ("PDF", "*.pdf")]
        )
        self.add_paths(paths)

    def add_paths(self, paths):
        existing = {str(Path(item["path"]).resolve()).lower() for item in self.files}
        added = 0
        for raw in paths:
            path = Path(raw)
            if path.suffix.lower() not in SUPPORTED or not path.is_file():
                continue
            key = str(path.resolve()).lower()
            if key in existing:
                continue
            self.files.append({"path": str(path), "name": path.stem})
            existing.add(key)
            added += 1
        self._refresh_file_list()
        if added:
            self.status.set(f"Đã thêm {added} file. Tổng cộng {len(self.files)} báo cáo.")

    def _refresh_file_list(self):
        self.file_list.delete(0, "end")
        for index, item in enumerate(self.files, 1):
            self.file_list.insert("end", f"{index:02d}. {item['name']}")

    def remove_file(self):
        selection = self.file_list.curselection()
        if not selection:
            return
        del self.files[selection[0]]
        self._refresh_file_list()
        self.status.set("Đã xóa file khỏi danh sách (file gốc không bị thay đổi).")

    def rename_source(self):
        selection = self.file_list.curselection()
        if not selection:
            messagebox.showinfo("Đổi tên nguồn", "Hãy chọn một báo cáo trong danh sách.")
            return
        index = selection[0]
        current = self.files[index]["name"]
        name = simpledialog.askstring("Đổi tên nguồn", "Tên hiển thị trong file tổng hợp:",
                                      initialvalue=current, parent=self.root)
        if not name or not name.strip():
            return
        self.files[index]["name"] = name.strip()
        self._refresh_file_list()
        if index < len(self.model.sources):
            self.model.update_source_name(self.model.sources[index].id, name.strip())
            self._refresh_tree()

    def analyze(self):
        if not self.files:
            messagebox.showinfo("Chưa có file", "Hãy thêm ít nhất một báo cáo Word hoặc PDF.")
            return
        self.status.set("Đang phân tích cấu trúc báo cáo...")
        self.root.config(cursor="wait")
        payload = [(item["path"], item["name"]) for item in self.files]

        def work():
            try:
                self.model.analyze(payload)
                self.root.after(0, self._analysis_done)
            except Exception as exc:
                self.root.after(0, lambda: self._operation_error("Không thể phân tích", exc))

        threading.Thread(target=work, daemon=True).start()

    def _analysis_done(self):
        self.root.config(cursor="")
        self._refresh_tree()
        section_count = len(self.model.groups)
        warning_count = sum(len(source.warnings) for source in self.model.sources)
        message = f"Đã phân tích {len(self.model.sources)} báo cáo, nhận diện {section_count} phần/mục."
        if warning_count:
            message += f" Có {warning_count} cảnh báo."
        self.status.set(message)

    def _refresh_tree(self, select_group: str | None = None):
        self.tree.delete(*self.tree.get_children())
        self.tree_to_group.clear()
        self.group_to_tree.clear()

        def add(parent_tree="", parent_group=None):
            for group in self.model.children(parent_group):
                mark = "✓" if group.enabled else "–"
                variants = len({occ.norm_title for occ in group.occurrences})
                suffix = f"  ({variants} tên biến thể)" if variants > 1 else ""
                item = self.tree.insert(parent_tree, "end", text=f"[{mark}] {group.title}{suffix}",
                                        values=(group.level, len(group.occurrences)), open=group.level <= 2)
                self.tree_to_group[item] = group.id
                self.group_to_tree[group.id] = item
                add(item, group.id)
        add()
        if select_group and select_group in self.group_to_tree:
            item = self.group_to_tree[select_group]
            self.tree.selection_set(item)
            self.tree.see(item)

    def _selected_group_ids(self):
        return [self.tree_to_group[item] for item in self.tree.selection() if item in self.tree_to_group]

    def rename_group(self):
        selected = self._selected_group_ids()
        if len(selected) != 1:
            messagebox.showinfo("Đổi tên", "Hãy chọn đúng một phần hoặc mục.")
            return
        group = self.model.group(selected[0])
        title = simpledialog.askstring("Đổi tên mục", "Tên chuẩn dùng trong file tổng hợp:",
                                       initialvalue=group.title, parent=self.root)
        if title and title.strip():
            self.model.rename_group(group.id, title)
            self._refresh_tree(group.id)
            self.status.set("Đã đổi tên mục. Nội dung nguyên văn không bị thay đổi.")

    def merge_groups(self):
        selected = self._selected_group_ids()
        try:
            target = self.model.merge_groups(selected)
            self._refresh_tree(target)
            self.status.set("Đã gộp các mục đã chọn và lưu quan hệ quy đổi trong phiên làm việc.")
        except Exception as exc:
            messagebox.showwarning("Không thể gộp", str(exc))

    def split_group(self):
        selected = self._selected_group_ids()
        if len(selected) != 1:
            messagebox.showinfo("Tách mục", "Hãy chọn đúng một mục đang chứa nhiều nguồn.")
            return
        try:
            created = self.model.split_group(selected[0])
            self._refresh_tree(created[0])
            self.status.set("Đã tách mục thành các mục riêng theo từng báo cáo nguồn.")
        except Exception as exc:
            messagebox.showwarning("Không thể tách", str(exc))

    def toggle_group(self):
        selected = self._selected_group_ids()
        if not selected:
            return
        for group_id in selected:
            self.model.toggle_group(group_id)
        self._refresh_tree(selected[0])

    def move_group(self, delta):
        selected = self._selected_group_ids()
        if len(selected) != 1:
            return
        self.model.move_group(selected[0], delta)
        self._refresh_tree(selected[0])

    def show_preview(self, _event=None):
        selected = self._selected_group_ids()
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        if len(selected) == 1:
            group = self.model.group(selected[0])
            self.preview.insert("end", f"{group.title}\n")
            self.preview.insert("end", "=" * min(64, len(group.title) + 8) + "\n\n")
            for occ in sorted(group.occurrences, key=lambda item: item.source_order):
                self.preview.insert("end", f"NGUỒN: {occ.source_name}\n")
                self.preview.insert("end", f"Tiêu đề gốc: {occ.title}\n\n")
                self.preview.insert("end", (occ.content_text or "[Mục không có nội dung trực tiếp]") + "\n\n")
        self.preview.configure(state="disabled")

    def save_aliases(self):
        if not self.model.groups:
            messagebox.showinfo("Chưa phân tích", "Hãy phân tích báo cáo trước khi lưu quy đổi.")
            return
        path = filedialog.asksaveasfilename(title="Lưu bảng quy đổi tiêu đề",
                                            defaultextension=".json",
                                            filetypes=[("JSON", "*.json")],
                                            initialfile="quy_doi_tieu_de.json")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(self.model.alias_payload(), handle, ensure_ascii=False, indent=2)
        self.status.set(f"Đã lưu bảng quy đổi: {path}")

    def load_aliases(self):
        path = filedialog.askopenfilename(title="Nạp bảng quy đổi tiêu đề",
                                          filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
            if not isinstance(payload, dict):
                raise ValueError("File quy đổi không đúng định dạng.")
            self.model.aliases = {str(key): str(value) for key, value in payload.items()}
            self.status.set("Đã nạp bảng quy đổi. Bấm “Phân tích cấu trúc” để áp dụng.")
        except Exception as exc:
            self._operation_error("Không thể nạp quy đổi", exc)

    def _ensure_analyzed(self):
        if not self.model.groups:
            messagebox.showinfo("Chưa phân tích", "Hãy bấm “Phân tích cấu trúc” trước khi xuất file.")
            return False
        return True

    def export_word(self):
        if not self._ensure_analyzed():
            return
        path = filedialog.asksaveasfilename(title="Xuất báo cáo tổng hợp Word",
                                            defaultextension=".docx",
                                            filetypes=[("Word", "*.docx")],
                                            initialfile="Bao cao tong hop.docx")
        if path:
            self._run_export(export_docx, path, self.output_title.get())

    def export_excel(self):
        if not self._ensure_analyzed():
            return
        path = filedialog.asksaveasfilename(title="Xuất bảng tổng hợp Excel",
                                            defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="Bang tong hop.xlsx")
        if path:
            self._run_export(export_xlsx, path)

    def _run_export(self, function, *args):
        self.root.config(cursor="wait")
        self.status.set("Đang tạo file đầu ra...")

        def work():
            try:
                result = function(self.model, *args)
                self.root.after(0, lambda: self._export_done(result))
            except Exception as exc:
                self.root.after(0, lambda: self._operation_error("Không thể xuất file", exc))
        threading.Thread(target=work, daemon=True).start()

    def _export_done(self, path):
        self.root.config(cursor="")
        self.status.set(f"Đã tạo: {path}")
        if messagebox.askyesno("Xuất file thành công", f"Đã tạo file:\n{path}\n\nBạn có muốn mở file không?"):
            try:
                os.startfile(path)
            except OSError:
                pass

    def _operation_error(self, title, exc):
        self.root.config(cursor="")
        self.status.set(str(exc))
        messagebox.showerror(title, str(exc))

    def run(self):
        self.root.mainloop()


def main():
    ReportMergerApp().run()


if __name__ == "__main__":
    main()
