import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, font
import pandas as pd
import json

class PySheetApp(tk.Tk):
    """
    一個使用 Python 和 Tkinter 開發的增強型表格編輯器應用程式。
    具有更精緻的 UI，並支援主題更換與解析度調整。
    支援讀取和儲存 XLSX, CSV, 和 JSON 檔案。
    """
    def __init__(self):
        super().__init__()
        self.title("Python 表格編輯器 (PySheet)")
        self.geometry("1200x750")  # 預設較大的解析度

        # --- 顏色與風格設定 ---
        self.style = ttk.Style(self)
        self.style.theme_use('clam') # 使用一個更現代的主題
        self._setup_styles()

        # --- 內部資料存儲 ---
        self.dataframe = pd.DataFrame()
        self.file_path = None

        # --- UI 元件 ---
        self._create_widgets()
        self._create_menu()

    def _setup_styles(self):
        """配置應用程式的視覺風格。"""
        self.bg_color = '#e0e0e0' # 預設背景色
        self.fg_color = '#212121' # 前景色
        self.entry_bg = '#ffffff'
        self.tree_bg = '#ffffff'
        self.select_bg = '#0078d7' # 選中項目的背景色

        self.configure(bg=self.bg_color)
        
        # 通用風格
        self.style.configure('.', background=self.bg_color, foreground=self.fg_color, font=('Microsoft YaHei UI', 10))
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground=self.fg_color, padding=5)
        self.style.configure('TEntry', fieldbackground=self.entry_bg)
        
        # Treeview 風格
        self.style.configure('Treeview.Heading', font=('Microsoft YaHei UI', 10, 'bold'), padding=5)
        self.style.configure('Treeview', rowheight=25, fieldbackground=self.tree_bg)
        self.style.map('Treeview', background=[('selected', self.select_bg)], foreground=[('selected', '#ffffff')])

    def _create_widgets(self):
        """創建應用程式的主要 UI 元件。"""
        main_frame = ttk.Frame(self, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 頂部框架：輸入框和按鈕 ---
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # 儲存格位置標籤
        self.cell_pos_label = ttk.Label(top_frame, text="儲存格:", font=('Microsoft YaHei UI', 10))
        self.cell_pos_label.pack(side=tk.LEFT, padx=(0, 5), anchor='w')

        # 獨立輸入框 (類似 Excel 的公式欄)
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(top_frame, textvariable=self.input_var, font=('Microsoft YaHei UI', 12))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.input_entry.bind("<Return>", self._update_cell_from_input)
        self.input_entry.bind("<FocusOut>", self._update_cell_from_input)

        # --- 表格框架：Treeview 和捲軸 ---
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Treeview 作為表格
        self.tree = ttk.Treeview(tree_frame, show='headings')
        
        # 捲軸
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)

        # 綁定事件
        self.tree.bind("<<TreeviewSelect>>", self._on_cell_select)
        self.tree.bind("<Button-1>", self._on_cell_select, add='+') # 確保點擊時能觸發

    def _create_menu(self):
        """創建應用程式的頂部選單欄。"""
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # 檔案選單
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檔案", menu=file_menu)
        file_menu.add_command(label="開啟...", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="另存為...", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.quit)

        # 設置選單
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="設置", menu=settings_menu)
        settings_menu.add_command(label="更換背景顏色...", command=self._change_background_color)
        
        resolution_menu = tk.Menu(settings_menu, tearoff=0)
        settings_menu.add_cascade(label="更改解析度", menu=resolution_menu)
        resolutions = ["1024x768", "1280x720", "1600x900", "1920x1080"]
        for res in resolutions:
            resolution_menu.add_command(label=res, command=lambda r=res: self._change_resolution(r))
        
        # 幫助選單
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="幫助", menu=help_menu)
        help_menu.add_command(label="關於...", command=self.show_about)

        # 快捷鍵綁定
        self.bind_all("<Control-o>", lambda event: self.open_file())
        self.bind_all("<Control-s>", lambda event: self.save_file())

    def _change_background_color(self):
        """開啟顏色選擇器並更新應用程式背景色。"""
        color_data = colorchooser.askcolor(title="選擇背景顏色", initialcolor=self.bg_color)
        if not color_data:
            return
        
        new_color = color_data[1] # 獲取十六進位顏色碼
        self.bg_color = new_color
        
        # 更新所有相關元件的顏色
        self.configure(bg=self.bg_color)
        self.style.configure('.', background=self.bg_color)
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color)

    def _change_resolution(self, resolution):
        """更改應用程式視窗的解析度。"""
        self.geometry(resolution)

    def _clear_treeview(self):
        """清除 Treeview 中的所有資料和欄位。"""
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()

    def _load_data_to_treeview(self):
        """將 self.dataframe 的資料載入到 Treeview 中並自動調整欄寬。"""
        self._clear_treeview()
        if self.dataframe.empty:
            return

        self.tree["columns"] = list(self.dataframe.columns)
        default_font = font.Font(font=('Microsoft YaHei UI', 10))

        for col in self.dataframe.columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            header_width = default_font.measure(col)
            # 確保 self.dataframe[col].dropna() 不為空
            content_series = self.dataframe[col].dropna()
            if not content_series.empty:
                max_width = max([default_font.measure(str(x)) for x in content_series])
            else:
                max_width = 0
            col_width = max(header_width, max_width) + 30 # 增加更多邊距
            self.tree.column(col, width=col_width, minwidth=60, anchor=tk.W)

        for index, row in self.dataframe.iterrows():
            values = [str(v) for v in row.values]
            self.tree.insert("", "end", values=values, iid=str(index))

    def _on_cell_select(self, event):
        """當使用者在 Treeview 中選擇一個儲存格時觸發。"""
        selected_items = self.tree.selection()
        if not selected_items:
            return

        item_id = selected_items[0]
        column_id = self.tree.identify_column(event.x)
        if not column_id: return
        
        column_index = int(column_id.replace('#', '')) - 1
        if column_index < 0 or column_index >= len(self.tree["columns"]): return
        
        column_name = self.tree["columns"][column_index]

        try:
            value = self.dataframe.loc[int(item_id), column_name]
        except (KeyError, IndexError):
            value = ""

        self.input_var.set(str(value))
        self.cell_pos_label.config(text=f"儲存格: {column_name}[{item_id}]")

    def _update_cell_from_input(self, event=None):
        """從頂部輸入框獲取值並更新選中的儲存格。"""
        selected_items = self.tree.selection()
        if not selected_items:
            return

        item_id = selected_items[0]
        
        try:
            label_text = self.cell_pos_label.cget("text")
            if "儲存格:" in label_text:
                column_name = label_text.split(":")[1].strip().split("[")[0]
            else:
                return
        except IndexError:
            return

        new_value = self.input_var.get()

        try:
            original_dtype = self.dataframe[column_name].dtype
            try:
                converted_value = pd.Series([new_value]).astype(original_dtype).iloc[0]
            except (ValueError, TypeError):
                converted_value = new_value

            self.dataframe.loc[int(item_id), column_name] = converted_value
            
            updated_row_values = [str(v) for v in self.dataframe.loc[int(item_id)].values]
            self.tree.item(item_id, values=updated_row_values)

        except (KeyError, IndexError, ValueError) as e:
            messagebox.showerror("更新錯誤", f"無法更新儲存格：\n{e}")
        
        self.tree.focus_set()

    def open_file(self, event=None):
        """開啟檔案對話框，讀取支援的檔案格式。"""
        path = filedialog.askopenfilename(
            filetypes=[
                ("支援的檔案", "*.xlsx *.csv *.json"),
                ("Excel 檔案", "*.xlsx"),
                ("CSV 檔案", "*.csv"),
                ("JSON 檔案", "*.json"),
                ("所有檔案", "*.*")
            ]
        )
        if not path:
            return

        self.file_path = path
        try:
            if path.endswith('.xlsx'):
                self.dataframe = pd.read_excel(path)
            elif path.endswith('.csv'):
                self.dataframe = pd.read_csv(path)
            elif path.endswith('.json'):
                try:
                    self.dataframe = pd.read_json(path, orient='records')
                except ValueError:
                    with open(path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    self.dataframe = pd.json_normalize(data)

            for col in self.dataframe.columns:
                self.dataframe[col] = self.dataframe[col].astype(str)

            self.title(f"Python 表格編輯器 - {path.split('/')[-1]}")
            self._load_data_to_treeview()

        except Exception as e:
            messagebox.showerror("開啟檔案錯誤", f"無法讀取檔案：\n{e}")
            self.dataframe = pd.DataFrame()
            self._clear_treeview()

    def save_file(self, event=None):
        """開啟另存為對話框，將資料儲存為支援的格式。"""
        if self.dataframe.empty:
            messagebox.showwarning("無資料", "表格中沒有資料可以儲存。")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel 檔案", "*.xlsx"),
                ("CSV 檔案", "*.csv"),
                ("JSON 檔案", "*.json")
            ]
        )
        if not path:
            return

        try:
            # 在儲存前，嘗試將資料轉換回數值類型
            df_to_save = self.dataframe.copy()
            for col in df_to_save.columns:
                df_to_save[col] = pd.to_numeric(df_to_save[col], errors='ignore')

            if path.endswith('.xlsx'):
                df_to_save.to_excel(path, index=False)
            elif path.endswith('.csv'):
                df_to_save.to_csv(path, index=False, encoding='utf-8-sig')
            elif path.endswith('.json'):
                df_to_save.to_json(path, orient='records', indent=4, force_ascii=False)
            
            messagebox.showinfo("儲存成功", f"檔案已成功儲存至：\n{path}")
            self.file_path = path
            self.title(f"Python 表格編輯器 - {path.split('/')[-1]}")

        except Exception as e:
            messagebox.showerror("儲存檔案錯誤", f"無法儲存檔案：\n{e}")

    def show_about(self):
        """顯示關於對話框。"""
        messagebox.showinfo(
            "關於 PySheet",
            "Python 表格編輯器 (PySheet) v2.0\n\n"
            "一個使用 Tkinter 和 Pandas 製作的簡易表格應用程式。\n"
            "支援 XLSX, CSV, 和 JSON 格式，並可自訂界面。\n"
            "開發者：Gemini"
        )


if __name__ == "__main__":
    # 為了讓 pandas 能處理 excel，需要安裝 openpyxl
    # pip install pandas openpyxl
    app = PySheetApp()
    app.mainloop()
