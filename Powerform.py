import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json

class PySheetApp(tk.Tk):
    """
    一個使用 Python 和 Tkinter 開發的簡易表格編輯器應用程式。
    支援讀取和儲存 XLSX, CSV, 和 JSON 檔案。
    """
    def __init__(self):
        super().__init__()
        self.title("Python 表格編輯器 (PySheet)")
        self.geometry("1000x700")
        self.configure(bg='#f0f0f0')

        # 內部資料存儲
        self.dataframe = pd.DataFrame()
        self.file_path = None

        # --- UI 元件 ---
        self._create_widgets()
        self._create_menu()

    def _create_widgets(self):
        """創建應用程式的主要 UI 元件。"""
        main_frame = tk.Frame(self, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- 頂部框架：輸入框和按鈕 ---
        top_frame = tk.Frame(main_frame, bg='#f0f0f0')
        top_frame.pack(fill=tk.X, pady=(0, 5))

        # 儲存格位置標籤
        self.cell_pos_label = tk.Label(top_frame, text="儲存格:", bg='#f0f0f0', fg='#333', font=('Arial', 10))
        self.cell_pos_label.pack(side=tk.LEFT, padx=(0, 5))

        # 獨立輸入框 (類似 Excel 的公式欄)
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(top_frame, textvariable=self.input_var, font=('Arial', 12))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.input_entry.bind("<Return>", self._update_cell_from_input)
        self.input_entry.bind("<FocusOut>", self._update_cell_from_input)

        # --- 表格框架：Treeview 和捲軸 ---
        tree_frame = tk.Frame(main_frame)
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
        
        # 幫助選單
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="幫助", menu=help_menu)
        help_menu.add_command(label="關於...", command=self.show_about)

        # 快捷鍵綁定
        self.bind_all("<Control-o>", lambda event: self.open_file())
        self.bind_all("<Control-s>", lambda event: self.save_file())

    def _clear_treeview(self):
        """清除 Treeview 中的所有資料和欄位。"""
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()

    def _load_data_to_treeview(self):
        """將 self.dataframe 的資料載入到 Treeview 中並自動調整欄寬。"""
        self._clear_treeview()
        if self.dataframe.empty:
            return

        # 設定欄位
        self.tree["columns"] = list(self.dataframe.columns)
        for col in self.dataframe.columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            # 根據欄位名稱和內容計算初始欄寬
            font = tk.font.Font()
            header_width = font.measure(col)
            max_width = max([font.measure(str(x)) for x in self.dataframe[col].dropna()])
            # 設定一個合理的最小和最大寬度
            col_width = max(header_width, max_width) + 20 # 增加一些邊距
            self.tree.column(col, width=col_width, minwidth=50, anchor=tk.W)

        # 插入資料行，並處理換行以自適應高度
        # ttk.Treeview 的行高是固定的，但可以透過插入換行符 '\n' 來模擬多行效果
        for index, row in self.dataframe.iterrows():
            # 將所有值轉換為字串
            values = [str(v) for v in row.values]
            self.tree.insert("", "end", values=values, iid=str(index))

    def _on_cell_select(self, event=None):
        """當使用者在 Treeview 中選擇一個儲存格時觸發。"""
        selected_items = self.tree.selection()
        if not selected_items:
            return

        item_id = selected_items[0]  # 獲取選中的行ID (我們設定為 dataframe 的索引)
        
        # 獲取點擊的具體欄位
        column_id = self.tree.identify_column(event.x)
        if not column_id: return # 如果點擊在行號區域外
        
        column_index = int(column_id.replace('#', '')) - 1
        column_name = self.tree["columns"][column_index]

        # 獲取儲存格的值
        try:
            value = self.dataframe.loc[int(item_id), column_name]
        except (KeyError, IndexError):
            value = "" # 如果索引出錯，顯示空值

        # 更新輸入框和位置標籤
        self.input_var.set(str(value))
        self.cell_pos_label.config(text=f"儲存格: {column_name}[{item_id}]")


    def _update_cell_from_input(self, event=None):
        """從頂部輸入框獲取值並更新選中的儲存格。"""
        selected_items = self.tree.selection()
        if not selected_items:
            return

        item_id = selected_items[0]
        
        # 確定是哪一欄被選中
        # 這部分比較棘手，因為 Treeview 的選擇是整行
        # 我們需要一個變數來追蹤最後點擊的欄位
        try:
            # 從標籤中解析出欄位名稱
            label_text = self.cell_pos_label.cget("text")
            if "儲存格:" in label_text:
                column_name = label_text.split(":")[1].strip().split("[")[0]
            else:
                return # 如果標籤格式不對，則不更新
        except IndexError:
            return

        new_value = self.input_var.get()

        try:
            # 更新 DataFrame
            # 嘗試轉換回原始資料類型，但為了簡單起見，這裡我們先全部存為字串
            original_dtype = self.dataframe[column_name].dtype
            try:
                # 嘗試轉換回原始類型
                converted_value = pd.Series([new_value]).astype(original_dtype).iloc[0]
            except (ValueError, TypeError):
                converted_value = new_value # 轉換失敗則保持字串

            self.dataframe.loc[int(item_id), column_name] = converted_value
            
            # 更新 Treeview 顯示
            # 重新整理整行資料以確保一致性
            updated_row_values = [str(v) for v in self.dataframe.loc[int(item_id)].values]
            self.tree.item(item_id, values=updated_row_values)

        except (KeyError, IndexError, ValueError) as e:
            messagebox.showerror("更新錯誤", f"無法更新儲存格：\n{e}")
        
        # 更新後，讓表格重新獲得焦點
        self.tree.focus_set()

    def open_file(self, event=None):
        """開啟檔案對話框，讀取支援的檔案格式。"""
        path = filedialog.askopenfilename(
            filetypes=[
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
                # JSON 有多種格式 (records, split, etc.)，這裡嘗試最常見的
                try:
                    self.dataframe = pd.read_json(path, orient='records')
                except ValueError:
                    # 如果 records 格式失敗，嘗試其他格式或直接載入
                    with open(path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    self.dataframe = pd.json_normalize(data)

            # 將所有資料轉換為字串類型以便顯示，避免類型問題
            for col in self.dataframe.columns:
                self.dataframe[col] = self.dataframe[col].astype(str)

            self.title(f"Python 表格編輯器 - {path.split('/')[-1]}")
            self._load_data_to_treeview()

        except Exception as e:
            messagebox.showerror("開啟檔案錯誤", f"無法讀取檔案：\n{e}")
            self.dataframe = pd.DataFrame() # 出錯時清空
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
            if path.endswith('.xlsx'):
                self.dataframe.to_excel(path, index=False)
            elif path.endswith('.csv'):
                self.dataframe.to_csv(path, index=False, encoding='utf-8-sig')
            elif path.endswith('.json'):
                self.dataframe.to_json(path, orient='records', indent=4, force_ascii=False)
            
            messagebox.showinfo("儲存成功", f"檔案已成功儲存至：\n{path}")
            self.file_path = path
            self.title(f"Python 表格編輯器 - {path.split('/')[-1]}")

        except Exception as e:
            messagebox.showerror("儲存檔案錯誤", f"無法儲存檔案：\n{e}")

    def show_about(self):
        """顯示關於對話框。"""
        messagebox.showinfo(
            "關於 PySheet",
            "Python 表格編輯器 (PySheet)\n\n"
            "一個使用 Tkinter 和 Pandas 製作的簡易表格應用程式。\n"
            "支援 XLSX, CSV, 和 JSON 格式。\n"
            "開發者：Gemini"
        )


if __name__ == "__main__":
    # 為了讓 pandas 能處理 excel，需要安裝 openpyxl
    # pip install pandas openpyxl
    app = PySheetApp()
    app.mainloop()
