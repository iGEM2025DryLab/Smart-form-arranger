import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, font, simpledialog
import pandas as pd
import numpy as np
from fuzzywuzzy import process

class FindHeaderDialog(tk.Toplevel):
    """
    一個用於查找行頭或列頭的獨立對話框。
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        self.title("查找行頭/列頭")
        self.geometry("400x150")
        self.resizable(False, False)

        self.last_find_index = -1
        self.last_search_term = ""
        self.last_search_type = ""

        self._create_widgets()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 查找內容 ---
        ttk.Label(main_frame, text="查找內容:").grid(row=0, column=0, sticky="w", pady=2)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(main_frame, textvariable=self.search_var)
        self.search_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=2)
        self.search_entry.focus_set()

        # --- 選項 ---
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=1, column=1, columnspan=2, sticky="w", pady=5)
        
        self.search_type_var = tk.StringVar(value="column")
        ttk.Radiobutton(options_frame, text="列頭", variable=self.search_type_var, value="column").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(options_frame, text="行頭", variable=self.search_type_var, value="row").pack(side=tk.LEFT)

        self.fuzzy_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="模糊搜索", variable=self.fuzzy_var).pack(side=tk.LEFT, padx=(20, 0))

        # --- 按鈕 ---
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=1, columnspan=2, sticky="e", pady=(10, 0))

        self.find_next_btn = ttk.Button(button_frame, text="查找下一個", command=self.find_next)
        self.find_next_btn.pack()

    def find_next(self):
        """查找下一個匹配的行頭或列頭。"""
        term = self.search_var.get()
        search_type = self.search_type_var.get()
        if not term: return

        # 如果搜索條件改變，重置搜索
        if term != self.last_search_term or search_type != self.last_search_type:
            self.last_find_index = -1
        
        self.last_search_term = term
        self.last_search_type = search_type

        start_index = self.last_find_index + 1
        
        if search_type == 'column':
            headers = list(self.parent.dataframe.columns)
        else: # row
            headers = list(self.parent.dataframe.index.map(str))

        if start_index >= len(headers):
            if messagebox.askyesno("查找完畢", "已搜索到結尾。是否從頭開始搜索？", parent=self):
                start_index = 0
            else:
                return

        # 遍歷查找
        for i in range(start_index, len(headers)):
            header = headers[i]
            match = False
            if self.fuzzy_var.get():
                # 模糊搜索，設定一個相似度閾值，例如 70
                if process.fuzz.ratio(term.lower(), header.lower()) > 70:
                    match = True
            else:
                # 精確搜索 (包含即可)
                if term.lower() in header.lower():
                    match = True
            
            if match:
                self.last_find_index = i
                if search_type == 'column':
                    self.parent.highlight_column(i)
                else: # row
                    self.parent.highlight_row(i)
                return
        
        # 如果從中間找到結尾都沒找到
        self.last_find_index = -1
        messagebox.showinfo("未找到", f"找不到匹配 '{term}' 的標題。", parent=self)

    def _on_close(self):
        self.parent.find_header_dialog = None
        self.destroy()

class PySheetApp(tk.Tk):
    """
    一個使用 Python 和 Tkinter 開發的專業版表格編輯器應用程式。
    將查找與替換功能整合至主 UI。
    """
    def __init__(self):
        super().__init__()
        self.title("Python 表格編輯器 (PySheet)")
        self.geometry("1200x750")

        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self._setup_styles()

        self.dataframe = pd.DataFrame()
        self.file_path = None
        self.clipboard_data = None
        self.undo_stack = []
        self.redo_stack = []
        
        # --- 查找相關狀態 ---
        self.find_replace_frame_visible = False
        self.last_find_coords = (-1, -1) # (row, col)
        self.find_header_dialog = None

        self._create_widgets()
        self._create_menu()
        self._create_context_menu()

    def _setup_styles(self):
        self.bg_color = '#e0e0e0'
        self.fg_color = '#212121'
        self.entry_bg = '#ffffff'
        self.tree_bg_even = '#ffffff'
        self.tree_bg_odd = '#f5f5f5'
        self.select_bg = '#0078d7'
        self.configure(bg=self.bg_color)
        self.style.configure('.', background=self.bg_color, foreground=self.fg_color, font=('Microsoft YaHei UI', 10))
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground=self.fg_color, padding=5)
        self.style.configure('TEntry', fieldbackground=self.entry_bg)
        self.style.configure('Treeview.Heading', font=('Microsoft YaHei UI', 10, 'bold'), padding=5)
        self.style.configure('Treeview', rowheight=25, fieldbackground=self.tree_bg_even)
        self.style.map('Treeview', background=[('selected', self.select_bg)], foreground=[('selected', '#ffffff')])

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- 頂部框架：儲存格輸入框 ---
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 5))
        self.cell_pos_label = ttk.Label(top_frame, text="儲存格:", font=('Microsoft YaHei UI', 10))
        self.cell_pos_label.pack(side=tk.LEFT, padx=(0, 5), anchor='w')
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(top_frame, textvariable=self.input_var, font=('Microsoft YaHei UI', 12))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.input_entry.bind("<Return>", self._update_cell_from_input)
        self.input_entry.bind("<FocusOut>", self._update_cell_from_input)

        # --- 新增：查找與替換框架 (預設隱藏) ---
        self._create_find_replace_frame(main_frame)

        # --- 表格框架 ---
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tree_frame, show='tree headings')
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.tag_configure('oddrow', background=self.tree_bg_odd)
        self.tree.tag_configure('evenrow', background=self.tree_bg_even)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_cell_select)
        self.tree.bind("<Button-3>", self._show_context_menu)

    def _create_find_replace_frame(self, parent):
        """創建可收合的查找與替換 UI 面板。"""
        self.find_replace_frame = ttk.Frame(parent, padding=5)
        
        ttk.Label(self.find_replace_frame, text="查找:").pack(side=tk.LEFT, padx=(0, 2))
        self.find_var = tk.StringVar()
        find_entry = ttk.Entry(self.find_replace_frame, textvariable=self.find_var, width=20)
        find_entry.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(self.find_replace_frame, text="替換為:").pack(side=tk.LEFT, padx=(5, 2))
        self.replace_var = tk.StringVar()
        replace_entry = ttk.Entry(self.find_replace_frame, textvariable=self.replace_var, width=20)
        replace_entry.pack(side=tk.LEFT, padx=(0, 5))

        find_next_btn = ttk.Button(self.find_replace_frame, text="下一個", command=self.find_next)
        find_next_btn.pack(side=tk.LEFT, padx=2)

        replace_btn = ttk.Button(self.find_replace_frame, text="替換", command=self.replace_cell)
        replace_btn.pack(side=tk.LEFT, padx=2)

        replace_all_btn = ttk.Button(self.find_replace_frame, text="全部替換", command=self.replace_all)
        replace_all_btn.pack(side=tk.LEFT, padx=2)
        
        # 使用一個更美觀的關閉按鈕
        close_btn = ttk.Button(self.find_replace_frame, text="×", command=self.toggle_find_replace_frame, width=3)
        close_btn.pack(side=tk.RIGHT, padx=(10, 0))

    def _create_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檔案", menu=file_menu)
        file_menu.add_command(label="新建", command=self.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="開啟...", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="保存", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="另存為...", command=self.save_file_as, accelerator="Ctrl+Shift+S")
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.quit)
        
        self.edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="編輯", menu=self.edit_menu)
        self.edit_menu.add_command(label="回滾", command=self.undo_action, accelerator="Ctrl+Z", state="disabled")
        self.edit_menu.add_command(label="前滾", command=self.redo_action, accelerator="Ctrl+Y", state="disabled")
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="剪下", command=self._cut_cell, accelerator="Ctrl+X")
        self.edit_menu.add_command(label="複製", command=self._copy_cell, accelerator="Ctrl+C")
        self.edit_menu.add_command(label="貼上", command=self._paste_cell, accelerator="Ctrl+V")
        
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="查看", menu=view_menu)
        view_menu.add_command(label="查找與替換...", command=self.toggle_find_replace_frame, accelerator="Ctrl+F")
        view_menu.add_command(label="查找行頭/列頭...", command=self.open_find_header_dialog)

        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="設置", menu=settings_menu)
        settings_menu.add_command(label="更換背景顏色...", command=self._change_background_color)
        resolution_menu = tk.Menu(settings_menu, tearoff=0)
        settings_menu.add_cascade(label="更改解析度", menu=resolution_menu)
        resolutions = ["1024x768", "1280x720", "1600x900", "1920x1080 (FHD)", "3840x2160 (4K)", "7680x4320 (8K)"]
        for res in resolutions:
            resolution_menu.add_command(label=res, command=lambda r=res: self._change_resolution(r))
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="幫助", menu=help_menu)
        help_menu.add_command(label="關於...", command=self.show_about)
        
        self.bind_all("<Control-n>", lambda e: self.new_file())
        self.bind_all("<Control-o>", lambda e: self.open_file())
        self.bind_all("<Control-s>", lambda e: self.save_file())
        self.bind_all("<Control-Shift-s>", lambda e: self.save_file_as())
        self.bind_all("<Control-z>", lambda e: self.undo_action())
        self.bind_all("<Control-y>", lambda e: self.redo_action())
        self.bind_all("<Control-x>", lambda e: self._cut_cell())
        self.bind_all("<Control-c>", lambda e: self._copy_cell())
        self.bind_all("<Control-v>", lambda e: self._paste_cell())
        self.bind_all("<Control-f>", lambda e: self.toggle_find_replace_frame())

    def _create_context_menu(self):
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="剪下", command=self._cut_cell)
        self.context_menu.add_command(label="複製", command=self._copy_cell)
        self.context_menu.add_command(label="貼上", command=self._paste_cell)
        self.context_menu.add_separator()
        insert_menu = tk.Menu(self.context_menu, tearoff=0)
        self.context_menu.add_cascade(label="插入", menu=insert_menu)
        insert_menu.add_command(label="插入行 (上方)", command=lambda: self._insert_row(above=True))
        insert_menu.add_command(label="插入行 (下方)", command=lambda: self._insert_row(above=False))
        insert_menu.add_command(label="插入列 (左側)", command=lambda: self._insert_column(left=True))
        insert_menu.add_command(label="插入列 (右側)", command=lambda: self._insert_column(left=False))
        delete_menu = tk.Menu(self.context_menu, tearoff=0)
        self.context_menu.add_cascade(label="刪除", menu=delete_menu)
        delete_menu.add_command(label="刪除行", command=self._delete_row)
        delete_menu.add_command(label="刪除列", command=self._delete_column)

    # --- 查找與替換邏輯 ---
    def open_find_header_dialog(self):
        """打開查找行頭/列頭對話框。"""
        if self.find_header_dialog is None or not self.find_header_dialog.winfo_exists():
            self.find_header_dialog = FindHeaderDialog(self)
        self.find_header_dialog.deiconify()
        self.find_header_dialog.lift()
        self.find_header_dialog.focus_set()

    def toggle_find_replace_frame(self):
        """顯示或隱藏查找與替換面板。"""
        if self.find_replace_frame_visible:
            self.find_replace_frame.pack_forget()
        else:
            # 將其放置在頂部輸入框和表格之間
            self.find_replace_frame.pack(fill=tk.X, pady=(0, 5), before=self.tree.master)
            # 將焦點設置到查找輸入框
            self.find_replace_frame.winfo_children()[1].focus_set()
        self.find_replace_frame_visible = not self.find_replace_frame_visible

    def find_next(self):
        """從當前位置開始查找下一個匹配項。"""
        find_term = self.find_var.get()
        if not find_term: return

        df = self.dataframe
        start_row, start_col = self.last_find_coords
        start_col += 1

        for i in range(start_row if start_row != -1 else 0, len(df)):
            for j in range(start_col if i == start_row else 0, len(df.columns)):
                cell_value = str(df.iat[i, j])
                if find_term in cell_value:
                    self.highlight_cell(i, j)
                    return
        
        if self.last_find_coords != (-1, -1):
            if messagebox.askyesno("查找完畢", "已搜索到文件結尾。是否從頭開始搜索？", parent=self):
                self.last_find_coords = (-1, -1)
                self.find_next()
        else:
            messagebox.showinfo("未找到", f"找不到 '{find_term}'。", parent=self)

    def replace_cell(self):
        """替換當前選中的匹配項並查找下一個。"""
        selected_items = self.tree.selection()
        if not selected_items: self.find_next(); return
        
        row_id_str, row_idx, col_idx = selected_items[0], *self.last_find_coords
        
        if self.dataframe.index.get_loc(int(row_id_str)) != row_idx or col_idx == -1:
            self.find_next(); return

        find_term, replace_term = self.find_var.get(), self.replace_var.get()
        col_name = self.dataframe.columns[col_idx]
        old_value = str(self.dataframe.iat[row_idx, col_idx])
        
        new_value = old_value.replace(find_term, replace_term, 1)
        
        self._add_to_undo(('edit', row_id_str, col_name, old_value))
        self._apply_change(row_id_str, col_name, new_value)
        self.find_next()

    def replace_all(self):
        """替換所有匹配項。"""
        find_term, replace_term = self.find_var.get(), self.replace_var.get()
        if not find_term: return

        original_df = self.dataframe.copy()
        new_df = self.dataframe.astype(str).applymap(lambda x: x.replace(find_term, replace_term))
        found_count = (self.dataframe.astype(str) != new_df.astype(str)).sum().sum()

        if found_count > 0:
            self._add_to_undo(('replace_all', original_df))
            self.dataframe = new_df
            self._load_data_to_treeview()
            messagebox.showinfo("完成", f"已完成 {found_count} 處替換。", parent=self)
        else:
            messagebox.showinfo("未找到", f"找不到 '{find_term}'。", parent=self)

    def highlight_cell(self, row_idx, col_idx):
        """在 Treeview 中高亮顯示指定的單元格。"""
        self.last_find_coords = (row_idx, col_idx)
        row_id_str = str(self.dataframe.index[row_idx])
        self.tree.selection_set(row_id_str)
        self.tree.focus(row_id_str)
        self.tree.see(row_id_str)

    def highlight_column(self, col_idx):
        """高亮顯示一列（通過滾動並選擇第一個單元格）。"""
        col_id = f"#{col_idx + 1}"
        self.tree.see(col_id)
        all_rows = self.tree.get_children()
        if all_rows:
            first_row_id = all_rows[0]
            self.tree.selection_set(first_row_id)
            self.tree.focus(first_row_id)
            col_name = self.dataframe.columns[col_idx]
            self.cell_pos_label.config(text=f"欄: {col_name}")

    def highlight_row(self, row_idx):
        """高亮顯示一行。"""
        row_id_str = str(self.dataframe.index[row_idx])
        self.tree.selection_set(row_id_str)
        self.tree.focus(row_id_str)
        self.tree.see(row_id_str)

    # --- 回滾/前滾與其他核心方法 ---
    def _update_edit_menu_state(self):
        self.edit_menu.entryconfig("回滾", state="normal" if self.undo_stack else "disabled")
        self.edit_menu.entryconfig("前滾", state="normal" if self.redo_stack else "disabled")
        paste_state = "normal" if self.clipboard_data is not None else "disabled"
        self.edit_menu.entryconfig("貼上", state=paste_state)

    def _add_to_undo(self, action):
        self.undo_stack.append(action)
        self.redo_stack.clear()
        self._update_edit_menu_state()

    def undo_action(self, event=None):
        if not self.undo_stack: return
        action = self.undo_stack.pop()
        action_type = action[0]
        if action_type == 'edit':
            _, row_id, col_name, old_value = action
            current_value = self.dataframe.loc[int(row_id), col_name]
            self.redo_stack.append(('edit', row_id, col_name, current_value))
            self._apply_change(row_id, col_name, old_value)
        elif action_type == 'delete_row':
            _, index, deleted_data = action
            part1 = self.dataframe.iloc[:index]
            part2 = self.dataframe.iloc[index:]
            self.dataframe = pd.concat([part1, deleted_data.to_frame().T, part2]).reset_index(drop=True)
            self.redo_stack.append(('insert_row', index))
            self._load_data_to_treeview()
        elif action_type == 'insert_row':
            _, index = action
            deleted_data = self.dataframe.loc[index].copy()
            self.dataframe = self.dataframe.drop(index).reset_index(drop=True)
            self.redo_stack.append(('delete_row', index, deleted_data))
            self._load_data_to_treeview()
        elif action_type == 'replace_all':
            _, old_dataframe = action
            self.redo_stack.append(('replace_all', self.dataframe.copy()))
            self.dataframe = old_dataframe
            self._load_data_to_treeview()
        self._update_edit_menu_state()

    def redo_action(self, event=None):
        if not self.redo_stack: return
        action = self.redo_stack.pop()
        action_type = action[0]
        if action_type == 'edit':
            _, row_id, col_name, new_value = action
            old_value = self.dataframe.loc[int(row_id), col_name]
            self.undo_stack.append(('edit', row_id, col_name, old_value))
            self._apply_change(row_id, col_name, new_value)
        elif action_type == 'delete_row':
            _, index, _ = action
            deleted_data = self.dataframe.loc[index].copy()
            self.dataframe = self.dataframe.drop(index).reset_index(drop=True)
            self.undo_stack.append(('insert_row', index, deleted_data))
            self._load_data_to_treeview()
        elif action_type == 'insert_row':
            _, index = action
            new_row = pd.DataFrame([[""] * len(self.dataframe.columns)], columns=self.dataframe.columns)
            part1 = self.dataframe.iloc[:index]
            part2 = self.dataframe.iloc[index:]
            self.dataframe = pd.concat([part1, new_row, part2]).reset_index(drop=True)
            self.undo_stack.append(('delete_row', index))
            self._load_data_to_treeview()
        elif action_type == 'replace_all':
            _, new_dataframe = action
            self.undo_stack.append(('replace_all', self.dataframe.copy()))
            self.dataframe = new_dataframe
            self._load_data_to_treeview()
        self._update_edit_menu_state()

    def _show_context_menu(self, event):
        self.context_menu_row_id = self.tree.identify_row(event.y)
        self.context_menu_col_id = self.tree.identify_column(event.x)
        if not self.context_menu_row_id: return
        if self.tree.selection() != (self.context_menu_row_id,): self.tree.selection_set(self.context_menu_row_id)
        paste_state = "normal" if self.clipboard_data is not None else "disabled"
        self.context_menu.entryconfig("貼上", state=paste_state)
        self.context_menu.post(event.x_root, event.y_root)
    def _get_context_cell(self):
        if not self.context_menu_row_id or not self.context_menu_col_id: return None, None
        row_id = self.context_menu_row_id
        try:
            col_index = int(self.context_menu_col_id.replace('#', '')) - 1
            if col_index < 0: return None, None
            col_name = self.dataframe.columns[col_index]
            return row_id, col_name
        except (ValueError, IndexError): return None, None
    def _copy_cell(self, event=None):
        row_id, col_name = self._get_context_cell()
        if row_id is None: return
        self.clipboard_data = self.dataframe.loc[int(row_id), col_name]
        self._update_edit_menu_state()
    def _cut_cell(self, event=None):
        row_id, col_name = self._get_context_cell()
        if row_id is None: return
        self._copy_cell()
        self._add_to_undo(('edit', row_id, col_name, self.clipboard_data))
        self._apply_change(row_id, col_name, "")
    def _paste_cell(self, event=None):
        if self.clipboard_data is None: return
        row_id, col_name = self._get_context_cell()
        if row_id is None: return
        old_value = self.dataframe.loc[int(row_id), col_name]
        self._add_to_undo(('edit', row_id, col_name, old_value))
        self._apply_change(row_id, col_name, self.clipboard_data)
    def _delete_row(self):
        if not self.context_menu_row_id: return
        row_index = self.dataframe.index.get_loc(int(self.context_menu_row_id))
        deleted_data = self.dataframe.loc[row_index].copy()
        self.dataframe = self.dataframe.drop(row_index).reset_index(drop=True)
        self._add_to_undo(('delete_row', row_index, deleted_data))
        self._load_data_to_treeview()
    def _insert_row(self, above=True):
        if not self.context_menu_row_id: return
        row_index = self.dataframe.index.get_loc(int(self.context_menu_row_id))
        insert_pos = row_index if above else row_index + 1
        new_row = pd.DataFrame([[""] * len(self.dataframe.columns)], columns=self.dataframe.columns)
        part1 = self.dataframe.iloc[:insert_pos]
        part2 = self.dataframe.iloc[insert_pos:]
        self.dataframe = pd.concat([part1, new_row, part2]).reset_index(drop=True)
        self._add_to_undo(('insert_row', insert_pos))
        self._load_data_to_treeview()
    def _delete_column(self):
        _, col_name = self._get_context_cell()
        if col_name is None: return
        if messagebox.askyesno("確認刪除", f"您確定要刪除 '{col_name}' 這一整列嗎？此操作目前無法復原。"):
            self.dataframe.drop(columns=[col_name], inplace=True)
            self._load_data_to_treeview()
    def _insert_column(self, left=True):
        _, col_name = self._get_context_cell()
        if col_name is None: return
        new_col_name = simpledialog.askstring("新列名稱", "請輸入新列的名稱：")
        if not new_col_name or new_col_name in self.dataframe.columns:
            messagebox.showwarning("無效名稱", "列名稱不可為空或與現有列重複。")
            return
        col_index = self.dataframe.columns.get_loc(col_name)
        insert_pos = col_index if left else col_index + 1
        self.dataframe.insert(insert_pos, new_col_name, "")
        self._load_data_to_treeview()
    def _apply_change(self, row_id, col_name, value):
        self.dataframe.loc[int(row_id), col_name] = value
        updated_row_values = [str(v) for v in self.dataframe.loc[int(row_id)].values]
        self.tree.item(str(row_id), values=updated_row_values)
    def _update_cell_from_input(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items: return
        item_id = selected_items[0]
        try:
            label_text = self.cell_pos_label.cget("text")
            if "儲存格:" in label_text: column_name = label_text.split(":")[1].strip().split("[")[0]
            else: return
        except IndexError: return
        new_value = self.input_var.get()
        old_value = self.dataframe.loc[int(item_id), column_name]
        if str(old_value) == str(new_value): return
        self._add_to_undo(('edit', item_id, column_name, old_value))
        self._apply_change(item_id, column_name, new_value)
        self.tree.focus_set()
    def _load_data_to_treeview(self):
        self._clear_treeview()
        if self.dataframe.empty: return
        self.tree["columns"] = list(self.dataframe.columns)
        default_font = font.Font(font=('Microsoft YaHei UI', 10))
        for col in self.dataframe.columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            header_width = default_font.measure(col)
            content_series = self.dataframe[col].dropna()
            max_width = max([default_font.measure(str(x)) for x in content_series]) if not content_series.empty else 0
            col_width = max(header_width, max_width) + 30
            self.tree.column(col, width=col_width, minwidth=60, anchor=tk.W)
        for index, row in self.dataframe.iterrows():
            values = [str(v) for v in row.values]
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", values=values, iid=str(index), tags=(tag,))
    def new_file(self, event=None):
        self.dataframe = pd.DataFrame([['', ''], ['', '']], columns=['欄位A', '欄位B'])
        self.file_path = None
        self.title("Python 表格編輯器 (PySheet) - 未命名")
        self._load_data_to_treeview()
        self.undo_stack.clear(); self.redo_stack.clear(); self._update_edit_menu_state()
    def open_file(self, event=None):
        path = filedialog.askopenfilename(filetypes=[("支援的檔案", "*.xlsx *.csv *.json"), ("Excel 檔案", "*.xlsx"), ("CSV 檔案", "*.csv"), ("JSON 檔案", "*.json"), ("所有檔案", "*.*")])
        if not path: return
        self.file_path = path
        try:
            if path.endswith('.xlsx'): self.dataframe = pd.read_excel(path)
            elif path.endswith('.csv'): self.dataframe = pd.read_csv(path)
            elif path.endswith('.json'):
                try: self.dataframe = pd.read_json(path, orient='records')
                except ValueError:
                    with open(path, 'r', encoding='utf-8') as f: data = pd.read_json(f); self.dataframe = pd.json_normalize(data)
            for col in self.dataframe.columns: self.dataframe[col] = self.dataframe[col].astype(str)
            self.title(f"Python 表格編輯器 - {path.split('/')[-1]}")
            self._load_data_to_treeview()
            self.undo_stack.clear(); self.redo_stack.clear(); self._update_edit_menu_state()
        except Exception as e: messagebox.showerror("開啟檔案錯誤", f"無法讀取檔案：\n{e}"); self.dataframe = pd.DataFrame(); self._clear_treeview()
    def save_file(self, event=None):
        if self.file_path: self._execute_save(self.file_path)
        else: self.save_file_as()
    def save_file_as(self, event=None):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 檔案", "*.xlsx"), ("CSV 檔案", "*.csv"), ("JSON 檔案", "*.json")])
        if path: self._execute_save(path)
    def _execute_save(self, path):
        if self.dataframe.empty: messagebox.showwarning("無資料", "表格中沒有資料可以儲存。"); return
        try:
            df_to_save = self.dataframe.copy()
            for col in df_to_save.columns: df_to_save[col] = pd.to_numeric(df_to_save[col], errors='ignore')
            if path.endswith('.xlsx'): df_to_save.to_excel(path, index=False)
            elif path.endswith('.csv'): df_to_save.to_csv(path, index=False, encoding='utf-8-sig')
            elif path.endswith('.json'): df_to_save.to_json(path, orient='records', indent=4, force_ascii=False)
            self.file_path = path; self.title(f"Python 表格編輯器 - {path.split('/')[-1]}"); messagebox.showinfo("儲存成功", f"檔案已成功儲存至：\n{path}")
        except Exception as e: messagebox.showerror("儲存檔案錯誤", f"無法儲存檔案：\n{e}")
    def _change_background_color(self):
        color_data = colorchooser.askcolor(title="選擇背景顏色", initialcolor=self.bg_color)
        if not color_data: return
        new_color = color_data[1]; self.bg_color = new_color; self.configure(bg=self.bg_color)
        self.style.configure('.', background=self.bg_color); self.style.configure('TFrame', background=self.bg_color); self.style.configure('TLabel', background=self.bg_color)
    def _change_resolution(self, resolution): self.geometry(resolution.split(" ")[0])
    def _clear_treeview(self): self.tree.delete(*self.tree.get_children()); self.tree["columns"] = ()
    def _on_cell_select(self, event):
        selected_items = self.tree.selection()
        if not selected_items: return
        item_id = selected_items[0]
        x = self.winfo_pointerx() - self.tree.winfo_rootx()
        column_id = self.tree.identify_column(x)
        if not column_id: return
        try:
            column_index = int(column_id.replace('#', '')) - 1
            if column_index < 0 or column_index >= len(self.tree["columns"]): return
            column_name = self.tree["columns"][column_index]
            value = self.dataframe.loc[int(item_id), column_name]
            self.input_var.set(str(value)); self.cell_pos_label.config(text=f"儲存格: {column_name}[{item_id}]")
        except (KeyError, IndexError, ValueError):
             self.input_var.set(""); self.cell_pos_label.config(text=f"儲存格:")
    def show_about(self): messagebox.showinfo("關於 PySheet", "Python 表格編輯器 (PySheet) v5.3\n\n一個使用 Tkinter 和 Pandas 製作的全功能表格應用程式。\n新增查找行/列頭功能，支援模糊搜索。\n開發者：Gemini")

if __name__ == "__main__":
    # 為了讓模糊搜索能運作，需要安裝 fuzzywuzzy 和 python-Levenshtein
    # pip install fuzzywuzzy python-Levenshtein
    app = PySheetApp()
    app.mainloop()
