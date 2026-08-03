"""
Document Search GUI Application (tkinter)
- Compatible with Windows 7/8/10/11 (Python 3.8 + tkinter)
- No browser required, standalone desktop application
- Blue color scheme, optimized layout
"""
import ctypes
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# === Enable DPI awareness for sharp rendering on high-DPI displays ===
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-Monitor DPI Aware V2 (Win 8.1+)
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Per-Monitor DPI Aware (Win 8.1+)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # Fallback for Windows 7
        except Exception:
            pass

# Determine base directory (supports both dev and PyInstaller frozen modes)
if getattr(sys, 'frozen', False):
    _APP_DIR = os.path.dirname(sys.executable)
    _BUNDLE_DIR = sys._MEIPASS
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))
    _BUNDLE_DIR = _APP_DIR

sys.path.insert(0, _BUNDLE_DIR)
from search_docs import search_docs, get_match_excerpts

# Allowed file extensions
ALLOWED_EXTENSIONS = {'.doc', '.docx'}

# === Blue color scheme ===
COLOR_PRIMARY = '#2563eb'        # blue-600
COLOR_PRIMARY_DARK = '#1d4ed8'   # blue-700
COLOR_PRIMARY_LIGHT = '#3b82f6'  # blue-500
COLOR_BG = '#eff6ff'             # blue-50
COLOR_CARD = '#ffffff'
COLOR_RESULT_BG = '#f8fafc'      # slate-50
COLOR_RESULT_HOVER = '#e0e7ff'   # indigo-100
COLOR_MATCH_BG = '#eff6ff'       # blue-50
COLOR_MATCH_BORDER = '#3b82f6'   # blue-500
COLOR_MATCH_HIGHLIGHT = '#bfdbfe' # blue-200
COLOR_MATCH_TEXT = '#1e40af'     # blue-800
COLOR_TITLE_TEXT = '#1e293b'     # slate-800
COLOR_SUBTITLE = '#64748b'       # slate-500
COLOR_BORDER = '#cbd5e1'         # slate-300
COLOR_FOCUS_BORDER = '#2563eb'   # blue-600

# === Font sizes (1.5x from original) ===
FONT_TITLE = ('Microsoft YaHei', 18, 'bold')
FONT_SUBTITLE = ('Microsoft YaHei', 11)
FONT_FORM_LABEL = ('Microsoft YaHei', 11, 'bold')
FONT_ENTRY = ('Microsoft YaHei', 14)
FONT_BUTTON = ('Microsoft YaHei', 14, 'bold')
FONT_CHECKBOX = ('Microsoft YaHei', 12)
FONT_RESULT_TITLE = ('Microsoft YaHei', 14, 'bold')
FONT_RESULT_COUNT = ('Microsoft YaHei', 12, 'bold')
FONT_FILENAME = ('Microsoft YaHei', 14)
FONT_ARROW = ('Microsoft YaHei', 11)
FONT_BANNER = ('Microsoft YaHei', 13)
FONT_PLACEHOLDER = ('Microsoft YaHei', 16)
FONT_MATCH_TITLE = ('Microsoft YaHei', 13, 'bold')
FONT_MATCH_TEXT = ('Microsoft YaHei', 14)
FONT_MATCH_HIGHLIGHT = ('Microsoft YaHei', 14, 'bold')
FONT_PREVIEW_HEADER = ('Microsoft YaHei', 12)
FONT_LOADING = ('Microsoft YaHei', 13)
FONT_ERROR = ('Microsoft YaHei', 14)
FONT_NO_RESULT = ('Microsoft YaHei', 16)


class DocumentSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文档搜索工具")
        self.root.geometry("1280x900")
        self.root.minsize(1000, 700)
        self.root.configure(bg=COLOR_BG)

        # State
        self.search_results = []
        self.preview_cache = {}
        self.expanded_items = set()
        self.searching = False

        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        # === Dropdown list font (14pt, clean and balanced) ===
        self.root.option_add('*TCombobox*Listbox.font',
                             ('Microsoft YaHei', 14))
        self.root.option_add('*TCombobox*Listbox.foreground', COLOR_TITLE_TEXT)
        self.root.option_add('*TCombobox*Listbox.background', 'white')
        self.root.option_add('*TCombobox*Listbox.selectForeground', 'white')
        self.root.option_add('*TCombobox*Listbox.selectBackground', COLOR_PRIMARY)
        self.root.option_add('*TCombobox*Listbox.justify', 'left')

        # Search button
        style.configure('Search.TButton', font=FONT_BUTTON,
                        foreground='white', background=COLOR_PRIMARY,
                        borderwidth=0, focusthickness=0)
        style.map('Search.TButton',
                  background=[('active', COLOR_PRIMARY_DARK),
                              ('disabled', '#94a3b8')],
                  foreground=[('disabled', '#e2e8f0')])

        # Refresh button
        style.configure('Refresh.TButton', font=('Microsoft YaHei', 16),
                        foreground=COLOR_PRIMARY, background=COLOR_CARD,
                        borderwidth=0, focusthickness=0)
        style.map('Refresh.TButton',
                  background=[('active', COLOR_RESULT_HOVER)])

        # Combobox
        style.configure('Custom.TCombobox', font=FONT_ENTRY,
                        fieldbackground='white', background='white',
                        bordercolor=COLOR_BORDER, lightcolor=COLOR_BORDER,
                        darkcolor=COLOR_BORDER, arrowcolor=COLOR_PRIMARY)
        style.map('Custom.TCombobox',
                  bordercolor=[('focus', COLOR_FOCUS_BORDER)],
                  lightcolor=[('focus', COLOR_FOCUS_BORDER)],
                  darkcolor=[('focus', COLOR_FOCUS_BORDER)])

        # Scrollbar
        style.configure('Vertical.TScrollbar',
                        background=COLOR_BORDER,
                        troughcolor=COLOR_BG,
                        arrowcolor=COLOR_PRIMARY)

    def _build_ui(self):
        # === Header ===
        header_frame = tk.Frame(self.root, bg=COLOR_PRIMARY, height=90)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        title_label = tk.Label(header_frame, text="📄 文档搜索工具",
                               font=FONT_TITLE,
                               fg='white', bg=COLOR_PRIMARY)
        title_label.pack(pady=(18, 0))

        subtitle_label = tk.Label(header_frame,
                                  text="搜索文件夹中的 .doc 和 .docx 文档内容",
                                  font=FONT_SUBTITLE,
                                  fg='#dbeafe', bg=COLOR_PRIMARY)
        subtitle_label.pack()

        # === Search Form ===
        form_frame = tk.Frame(self.root, bg=COLOR_CARD)
        form_frame.pack(fill=tk.X)

        form_inner = tk.Frame(form_frame, bg=COLOR_CARD)
        form_inner.pack(padx=40, pady=22, fill=tk.X)

        # Search string
        tk.Label(form_inner, text="搜索字符串",
                 font=FONT_FORM_LABEL,
                 fg=COLOR_TITLE_TEXT, bg=COLOR_CARD).pack(
            anchor=tk.W, pady=(0, 6))
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(form_inner, textvariable=self.search_var,
                                     font=FONT_ENTRY,
                                     relief=tk.SOLID, bd=2,
                                     highlightthickness=2,
                                     highlightcolor=COLOR_FOCUS_BORDER,
                                     highlightbackground=COLOR_BORDER,
                                     insertbackground='#333333')
        self.search_entry.pack(fill=tk.X, ipady=8, pady=(0, 18))
        self.search_entry.bind('<Return>', lambda e: self._on_search())

        # Folder selection + browse button
        folder_top = tk.Frame(form_inner, bg=COLOR_CARD)
        folder_top.pack(fill=tk.X, pady=(0, 6))
        tk.Label(folder_top, text="选择文件夹",
                 font=FONT_FORM_LABEL,
                 fg=COLOR_TITLE_TEXT, bg=COLOR_CARD).pack(
            side=tk.LEFT)

        folder_input_frame = tk.Frame(form_inner, bg=COLOR_CARD)
        folder_input_frame.pack(fill=tk.X, pady=(0, 18))

        self.folder_var = tk.StringVar()
        self.folder_entry = tk.Entry(folder_input_frame,
                                     textvariable=self.folder_var,
                                     font=FONT_ENTRY,
                                     relief=tk.SOLID, bd=2,
                                     highlightthickness=2,
                                     highlightcolor=COLOR_FOCUS_BORDER,
                                     highlightbackground=COLOR_BORDER,
                                     insertbackground='#333333',
                                     state='readonly')
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True,
                               ipady=8, padx=(0, 10))

        self.browse_btn = ttk.Button(folder_input_frame, text="浏览...",
                                     style='Search.TButton',
                                     command=self._browse_folder)
        self.browse_btn.pack(side=tk.RIGHT, ipady=10)

        # Case sensitive checkbox
        case_frame = tk.Frame(form_inner, bg=COLOR_CARD)
        case_frame.pack(fill=tk.X, pady=(0, 18))
        self.case_var = tk.BooleanVar(value=False)
        case_cb = tk.Checkbutton(case_frame, text="区分大小写",
                                 variable=self.case_var,
                                 font=FONT_CHECKBOX,
                                 bg=COLOR_CARD, fg=COLOR_TITLE_TEXT,
                                 activebackground=COLOR_CARD,
                                 activeforeground=COLOR_TITLE_TEXT,
                                 selectcolor=COLOR_CARD)
        case_cb.pack(anchor=tk.W)

        # Search button
        self.search_btn = ttk.Button(form_inner, text="开始搜索",
                                     style='Search.TButton',
                                     command=self._on_search)
        self.search_btn.pack(fill=tk.X, ipady=10)

        # === Results Area ===
        results_container = tk.Frame(self.root, bg=COLOR_BG)
        results_container.pack(fill=tk.BOTH, expand=True, padx=40, pady=(0, 22))

        # Results header
        self.results_header = tk.Frame(results_container, bg=COLOR_BG)
        self.results_header.pack(fill=tk.X, pady=(0, 12))

        self.results_title = tk.Label(self.results_header, text="",
                                      font=FONT_RESULT_TITLE,
                                      fg=COLOR_TITLE_TEXT, bg=COLOR_BG)
        self.results_title.pack(side=tk.LEFT)

        self.results_count = tk.Label(self.results_header, text="",
                                      font=FONT_RESULT_COUNT,
                                      fg='white', bg=COLOR_PRIMARY)

        # Scrollable results
        self.canvas = tk.Canvas(results_container, bg=COLOR_BG,
                                highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(results_container,
                                       orient=tk.VERTICAL,
                                       command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=COLOR_BG)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Initial placeholder
        self._show_placeholder()

    def _on_canvas_configure(self, event):
        width = event.width
        self.canvas.itemconfig(self.canvas_window, width=width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _show_placeholder(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        placeholder = tk.Label(self.scrollable_frame,
                               text="🔍 输入关键词后点击「开始搜索」",
                               font=FONT_PLACEHOLDER,
                               fg='#94a3b8', bg=COLOR_BG)
        placeholder.pack(pady=100)

    def _browse_folder(self):
        """Open system folder picker dialog."""
        selected = filedialog.askdirectory(title="选择要搜索的文件夹")
        if selected:
            self.folder_var.set(selected)

    def _on_search(self):
        if self.searching:
            return

        search_str = self.search_var.get().strip()
        if not search_str:
            messagebox.showwarning("提示", "请输入搜索字符串")
            return

        folder_path = self.folder_var.get().strip()
        if not folder_path:
            messagebox.showwarning("提示", "请先选择文件夹")
            return

        if not os.path.isdir(folder_path):
            messagebox.showerror("错误", f'文件夹路径 "{folder_path}" 不存在')
            return

        # Start search in background thread (includes pre-validation)
        self.searching = True
        self.search_btn.configure(state=tk.DISABLED, text="搜索中...")
        self._show_searching()

        thread = threading.Thread(target=self._search_worker,
                                  args=(folder_path, search_str,
                                        self.case_var.get()),
                                  daemon=True)
        thread.start()

    def _show_searching(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        loading_frame = tk.Frame(self.scrollable_frame, bg=COLOR_BG)
        loading_frame.pack(pady=100)
        tk.Label(loading_frame, text="⏳ 搜索中，请稍候...",
                 font=FONT_LOADING,
                 fg=COLOR_PRIMARY, bg=COLOR_BG).pack()
        self.progress_label = tk.Label(loading_frame, text="",
                                       font=FONT_SUBTITLE,
                                       fg=COLOR_SUBTITLE, bg=COLOR_BG)
        self.progress_label.pack(pady=(6, 0))

    def _on_search_progress(self, processed, total):
        """Thread-safe progress callback (called from search thread)."""
        self.root.after(0, self._update_progress, processed, total)

    def _update_progress(self, processed, total):
        if hasattr(self, 'progress_label') and self.progress_label.winfo_exists():
            self.progress_label.config(
                text=f"已搜索 {processed} / {total} 个文件...")

    def _validate_folder(self, folder_path):
        """Pre-scan folder for non-doc/docx files.
        Returns list of bad file paths (empty if all valid).
        """
        bad_files = []
        for root, dirs, files in os.walk(folder_path):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext not in ALLOWED_EXTENSIONS:
                    bad_files.append(os.path.join(root, f))
        return bad_files

    def _search_worker(self, folder_path, search_str, case_sensitive):
        try:
            # Pre-validate: check for non-doc/docx files
            bad_files = self._validate_folder(folder_path)
            if bad_files:
                self.root.after(0, self._validation_error,
                                bad_files, search_str)
                return

            results = search_docs(folder_path, search_str, case_sensitive,
                                  progress_callback=self._on_search_progress)
            self.root.after(0, self._display_results, results, search_str)
        except Exception as e:
            self.root.after(0, self._search_error, str(e))

    def _validation_error(self, bad_files, search_str):
        self.searching = False
        self.search_btn.configure(state=tk.NORMAL, text="开始搜索")

        # Build error message with specific file names
        display_files = bad_files[:30]
        bad_list = '\n'.join(display_files)
        if len(bad_files) > 30:
            bad_list += f'\n... 等共 {len(bad_files)} 个文件'

        messagebox.showerror(
            "文件类型错误",
            f'文件夹中包含 {len(bad_files)} 个非 .doc/.docx 文件，已终止搜索：\n\n{bad_list}')

        # Show error in results area
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        error_frame = tk.Frame(self.scrollable_frame, bg=COLOR_BG)
        error_frame.pack(pady=80)

        tk.Label(error_frame, text="❌", font=('Microsoft YaHei', 48),
                 fg='#dc2626', bg=COLOR_BG).pack()
        tk.Label(error_frame,
                 text=f"文件夹包含 {len(bad_files)} 个非 .doc/.docx 文件",
                 font=FONT_ERROR, fg='#dc2626', bg=COLOR_BG).pack(pady=(12, 0))
        tk.Label(error_frame,
                 text="已终止搜索，请检查文件夹内容后重试",
                 font=FONT_SUBTITLE, fg=COLOR_SUBTITLE,
                 bg=COLOR_BG).pack(pady=(6, 0))

    def _search_error(self, error_msg):
        self.searching = False
        self.search_btn.configure(state=tk.NORMAL, text="开始搜索")
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        tk.Label(self.scrollable_frame,
                 text=f"❌ 错误: {error_msg}",
                 font=FONT_ERROR, fg='#dc2626',
                 bg=COLOR_BG, wraplength=1000, justify=tk.LEFT).pack(pady=100)

    def _display_results(self, results, search_str):
        self.searching = False
        self.search_btn.configure(state=tk.NORMAL, text="开始搜索")
        self.search_results = results
        self.preview_cache = {}
        self.expanded_items = set()

        count = len(results)
        self.results_title.config(text="搜索结果")
        if count > 0:
            self.results_count.config(text=f" {count} 个文件 ")
            self.results_count.pack(side=tk.RIGHT, padx=(10, 0))
        else:
            self.results_count.pack_forget()

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if count == 0:
            no_result = tk.Frame(self.scrollable_frame, bg=COLOR_BG)
            no_result.pack(pady=100)
            tk.Label(no_result, text="🔍", font=('Microsoft YaHei', 48),
                     fg='#cbd5e1', bg=COLOR_BG).pack()
            tk.Label(no_result,
                     text=f'未找到包含 "{search_str}" 的文档',
                     font=FONT_NO_RESULT,
                     fg='#94a3b8', bg=COLOR_BG).pack(pady=(12, 0))
            return

        # Success banner
        banner = tk.Frame(self.scrollable_frame, bg='#dbeafe', relief=tk.SOLID,
                          bd=0, highlightbackground=COLOR_PRIMARY_LIGHT,
                          highlightthickness=1)
        banner.pack(fill=tk.X, pady=(0, 12), ipady=10)
        tk.Label(banner,
                 text=f'✅ 成功找到 {count} 个包含 "{search_str}" 的文档，点击文档查看匹配内容',
                 font=FONT_BANNER, fg=COLOR_PRIMARY_DARK, bg='#dbeafe').pack(padx=12)

        # Result items
        for index, file_path in enumerate(results):
            self._create_result_item(index, file_path, search_str)

        self.canvas.yview_moveto(0)

    def _create_result_item(self, index, file_path, search_str):
        filename = os.path.basename(file_path)

        item_frame = tk.Frame(self.scrollable_frame, bg=COLOR_RESULT_BG,
                              relief=tk.SOLID, bd=0,
                              highlightbackground=COLOR_BORDER,
                              highlightthickness=1)
        item_frame.pack(fill=tk.X, pady=4)

        # Header (clickable)
        header = tk.Frame(item_frame, bg=COLOR_RESULT_BG, cursor='hand2')
        header.pack(fill=tk.X, padx=12, pady=12)
        header.bind('<Button-1>', lambda e, idx=index: self._toggle_item(idx))
        header.bind('<Enter>', lambda e, f=header: f.config(bg=COLOR_RESULT_HOVER))
        header.bind('<Leave>', lambda e, f=header: f.config(bg=COLOR_RESULT_BG))

        tk.Label(header, text="📝", font=('Microsoft YaHei', 16),
                 bg=COLOR_RESULT_BG).pack(side=tk.LEFT, padx=(0, 10))
        fname_label = tk.Label(header, text=filename,
                               font=FONT_FILENAME,
                               fg=COLOR_TITLE_TEXT, bg=COLOR_RESULT_BG,
                               anchor=tk.W)
        fname_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        fname_label.bind('<Button-1>', lambda e, idx=index: self._toggle_item(idx))

        arrow_label = tk.Label(header, text="▼", font=FONT_ARROW,
                               fg='#94a3b8', bg=COLOR_RESULT_BG)
        arrow_label.pack(side=tk.RIGHT)
        arrow_label.bind('<Button-1>', lambda e, idx=index: self._toggle_item(idx))

        # Preview container (hidden by default)
        preview_frame = tk.Frame(item_frame, bg=COLOR_CARD)
        item_frame.preview_frame = preview_frame
        item_frame.arrow_label = arrow_label
        item_frame.header = header

        self.scrollable_frame.result_items = getattr(
            self.scrollable_frame, 'result_items', {})
        self.scrollable_frame.result_items[index] = item_frame

    def _toggle_item(self, index):
        items = getattr(self.scrollable_frame, 'result_items', {})
        if index not in items:
            return

        item_frame = items[index]
        preview_frame = item_frame.preview_frame
        arrow_label = item_frame.arrow_label

        if index in self.expanded_items:
            # Collapse
            self.expanded_items.discard(index)
            preview_frame.pack_forget()
            arrow_label.config(text="▼")
        else:
            # Expand
            self.expanded_items.add(index)
            arrow_label.config(text="▲")
            preview_frame.pack(fill=tk.X, padx=12, pady=(0, 12))

            if index not in self.preview_cache:
                # Show loading
                tk.Label(preview_frame, text="⏳ 加载中...",
                         font=FONT_LOADING,
                         fg='#94a3b8', bg=COLOR_CARD).pack(pady=12)

                file_path = self.search_results[index]
                search_str = self.search_var.get().strip()
                case_sensitive = self.case_var.get()

                thread = threading.Thread(
                    target=self._preview_worker,
                    args=(index, file_path, search_str, case_sensitive),
                    daemon=True)
                thread.start()
            else:
                self._render_preview(index, preview_frame)

    def _preview_worker(self, index, file_path, search_str, case_sensitive):
        try:
            excerpts = get_match_excerpts(file_path, search_str, case_sensitive)
            self.root.after(0, self._preview_ready, index, excerpts)
        except Exception as e:
            self.root.after(0, self._preview_ready, index, [], str(e))

    def _preview_ready(self, index, excerpts, error=None):
        self.preview_cache[index] = (excerpts, error)
        if index in self.expanded_items:
            items = getattr(self.scrollable_frame, 'result_items', {})
            if index in items:
                preview_frame = items[index].preview_frame
                self._render_preview(index, preview_frame)

    @staticmethod
    def _is_chinese_text(text):
        """Detect if text contains Chinese characters (CJK Unified Ideographs)."""
        for ch in text:
            if '\u4e00' <= ch <= '\u9fff':
                return True
        return False

    def _render_preview(self, index, preview_frame):
        # Clear loading
        for widget in preview_frame.winfo_children():
            widget.destroy()

        excerpts, error = self.preview_cache.get(index, ([], None))

        if error:
            tk.Label(preview_frame, text=f"❌ 加载失败: {error}",
                     font=FONT_PREVIEW_HEADER, fg='#dc2626',
                     bg='#fef2f2', anchor=tk.W, justify=tk.LEFT).pack(
                fill=tk.X, padx=6, pady=6)
            return

        # Auto-detect search language: Chinese search → show Chinese results,
        # English search → show English results
        search_str = self.search_var.get().strip()
        searching_chinese = self._is_chinese_text(search_str)

        if searching_chinese:
            filtered = [e for e in excerpts if not e.get('is_english', False)]
        else:
            filtered = [e for e in excerpts if e.get('is_english', False)]

        count = len(filtered)
        lang_label = "中文" if searching_chinese else "英文"
        header_text = f"找到 {count} 处{lang_label}匹配" if count > 0 else "未找到匹配内容"
        tk.Label(preview_frame, text=header_text,
                 font=FONT_PREVIEW_HEADER, fg=COLOR_SUBTITLE,
                 bg=COLOR_CARD, anchor=tk.W).pack(fill=tk.X, pady=(6, 10))

        if count == 0:
            return

        for i, excerpt in enumerate(filtered):
            self._render_excerpt(preview_frame, excerpt, i)

    def _render_excerpt(self, parent, excerpt, idx):
        text = excerpt.get('text', '')
        match_start = excerpt.get('match_start', 0)
        match_end = excerpt.get('match_end', 0)

        # Separator
        if idx > 0:
            sep = tk.Frame(parent, bg=COLOR_BORDER, height=1)
            sep.pack(fill=tk.X, pady=12)

        # Match title
        tk.Label(parent, text="🎯 匹配内容",
                 font=FONT_MATCH_TITLE,
                 fg=COLOR_MATCH_TEXT, bg=COLOR_CARD,
                 anchor=tk.W).pack(fill=tk.X, pady=(0, 6))

        # Match paragraph with highlighting
        match_frame = tk.Frame(parent, bg=COLOR_MATCH_BG, relief=tk.SOLID, bd=0)
        match_frame.pack(fill=tk.X)

        # Left border color bar
        bar = tk.Frame(match_frame, bg=COLOR_MATCH_BORDER, width=5)
        bar.pack(side=tk.LEFT, fill=tk.Y)

        text_widget = tk.Text(match_frame, wrap=tk.WORD,
                              font=FONT_MATCH_TEXT,
                              bg=COLOR_MATCH_BG, fg='#475569',
                              relief=tk.FLAT, bd=0,
                              padx=12, pady=10,
                              height=self._calc_text_height(text),
                              cursor='arrow')
        text_widget.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Configure tags for highlighting
        text_widget.tag_configure('highlight',
                                  background=COLOR_MATCH_HIGHLIGHT,
                                  foreground=COLOR_MATCH_TEXT,
                                  font=FONT_MATCH_HIGHLIGHT)

        # Insert text with highlight
        before = text[:match_start]
        match = text[match_start:match_end]
        after = text[match_end:]

        text_widget.insert(tk.END, before)
        if match:
            text_widget.insert(tk.END, match, 'highlight')
        text_widget.insert(tk.END, after)
        text_widget.config(state=tk.DISABLED)

    def _calc_text_height(self, text):
        """Estimate Text widget height based on content length."""
        # Roughly 40 chars per line at larger font, min 2 lines
        lines = max(2, len(text) // 40 + 1)
        return min(lines, 20)


def main():
    try:
        root = tk.Tk()
        # Adjust tkinter scaling to match system DPI for crisp rendering
        try:
            import ctypes
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, hdc)
            if dpi > 0:
                root.tk.call('tk', 'scaling', dpi / 72)
        except Exception:
            pass
        app = DocumentSearchApp(root)
        root.mainloop()
    except Exception as e:
        try:
            from tkinter import messagebox
            messagebox.showerror("程序错误", f"程序发生错误:\n{e}\n\n请联系技术支持。")
        except Exception:
            pass


if __name__ == '__main__':
    main()
