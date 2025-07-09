import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, font, simpledialog
import pandas as pd
import numpy as np
from fuzzywuzzy import process
import matplotlib
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import re
from scipy.optimize import curve_fit
from scipy import stats
from sklearn.metrics import r2_score

# --- Matplotlib 中文顯示設定 ---
# 設置一個支援中文的字體，例如 'Microsoft YaHei' (Windows) 或 'PingFang TC' (macOS)
# 並處理負號顯示問題
try:
    matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'PingFang TC', 'SimHei']
    matplotlib.rcParams['axes.unicode_minus'] = False
except Exception as e:
    print(f"無法設置中文字體，圖表中的中文可能無法正常顯示: {e}")


# Helper function for column name generation
def _to_excel_col(n):
    """Converts a 1-based integer to an Excel-style column name (A, B, ..., Z, AA, ...)."""
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

class AnalysisDialog(tk.Toplevel):
    """
    一個用於分析兩列數據之間關係的對話框。
    """
    def __init__(self, parent, df):
        super().__init__(parent)
        self.parent = parent
        self.df = df.copy()
        self.title("關係分析")
        self.geometry("800x700")

        self._create_widgets()
        # 確保此子視窗關閉時能釋放資源
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 控制面板 ---
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, pady=5)

        ttk.Label(controls_frame, text="X 軸:").grid(row=0, column=0, padx=5, pady=5)
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(controls_frame, textvariable=self.x_var, values=list(self.df.columns), state='readonly')
        self.x_combo.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(controls_frame, text="Y 軸:").grid(row=0, column=2, padx=5, pady=5)
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(controls_frame, textvariable=self.y_var, values=list(self.df.columns), state='readonly')
        self.y_combo.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(controls_frame, text="模型:").grid(row=1, column=0, padx=5, pady=5)
        models = ["線性", "多項式", "指數", "對數"]
        self.model_var = tk.StringVar(value="線性")
        self.model_combo = ttk.Combobox(controls_frame, textvariable=self.model_var, values=models, state='readonly')
        self.model_combo.grid(row=1, column=1, padx=5, pady=5)
        self.model_combo.bind("<<ComboboxSelected>>", self._on_model_select)

        self.poly_degree_label = ttk.Label(controls_frame, text="階數:")
        self.poly_degree_var = tk.StringVar(value="2")
        self.poly_degree_entry = ttk.Entry(controls_frame, textvariable=self.poly_degree_var, width=5)

        ttk.Button(controls_frame, text="開始分析", command=self.run_analysis).grid(row=1, column=4, padx=20, pady=5)

        # --- 結果顯示區 ---
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.fig, self.ax = plt.subplots(figsize=(7, 4), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, master=result_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill=tk.BOTH, expand=True)

        self.conclusion_label = ttk.Label(result_frame, text="結論:", font=('Microsoft YaHei UI', 12, 'bold'))
        self.conclusion_label.pack(fill=tk.X, pady=(10, 0))
        self.conclusion_text = tk.Text(result_frame, height=6, wrap=tk.WORD, state='disabled', font=('Microsoft YaHei UI', 10))
        self.conclusion_text.pack(fill=tk.X)

    def _on_model_select(self, event=None):
        if self.model_var.get() == "多項式":
            self.poly_degree_label.grid(row=1, column=2, padx=(10, 2), pady=5)
            self.poly_degree_entry.grid(row=1, column=3, padx=(0, 5), pady=5)
        else:
            self.poly_degree_label.grid_forget()
            self.poly_degree_entry.grid_forget()

    def run_analysis(self):
        x_col = self.x_var.get()
        y_col = self.y_var.get()
        model_type = self.model_var.get()

        if not x_col or not y_col:
            messagebox.showerror("錯誤", "請選擇 X 軸和 Y 軸。", parent=self)
            return

        try:
            x_data = pd.to_numeric(self.df[x_col], errors='coerce').dropna()
            y_data = pd.to_numeric(self.df[y_col], errors='coerce').dropna()
            
            # 對齊數據
            common_index = x_data.index.intersection(y_data.index)
            if len(common_index) < 3:
                messagebox.showerror("數據不足", "沒有足夠的重疊有效數據進行分析 (至少需要 3 個點)。", parent=self)
                return
            x_data = x_data[common_index]
            y_data = y_data[common_index]

        except Exception as e:
            messagebox.showerror("數據錯誤", f"無法處理數據：\n{e}", parent=self)
            return

        self.ax.clear()
        self.ax.scatter(x_data, y_data, label='原始數據', alpha=0.6)
        
        equation = ""
        r2 = 0.0
        
        # --- 計算通用相關係數 ---
        pearson_corr, pearson_p = stats.pearsonr(x_data, y_data)
        spearman_corr, spearman_p = stats.spearmanr(x_data, y_data)

        try:
            if model_type == "線性":
                popt, _ = curve_fit(lambda x, a, b: a * x + b, x_data, y_data)
                y_pred = popt[0] * x_data + popt[1]
                equation = f"y = {popt[0]:.4f}x + {popt[1]:.4f}"
            
            elif model_type == "多項式":
                degree = int(self.poly_degree_var.get())
                if degree < 1:
                    messagebox.showerror("錯誤", "多項式階數必須大於等於 1。", parent=self)
                    return
                
                popt = np.polyfit(x_data, y_data, degree)
                poly_func = np.poly1d(popt)
                y_pred = poly_func(x_data)
                
                terms = []
                for i, p in enumerate(popt):
                    power = degree - i
                    if power > 1:
                        terms.append(f"{p:.4f}x^{power}")
                    elif power == 1:
                        terms.append(f"{p:.4f}x")
                    else:
                        terms.append(f"{p:.4f}")
                equation = "y = " + " + ".join(terms).replace("+ -", "- ")

            elif model_type == "指數":
                popt, _ = curve_fit(lambda x, a, b, c: a * np.exp(b * x) + c, x_data, y_data, maxfev=5000)
                y_pred = popt[0] * np.exp(popt[1] * x_data) + popt[2]
                equation = f"y = {popt[0]:.4f} * exp({popt[1]:.4f}x) + {popt[2]:.4f}"

            elif model_type == "對數":
                if (x_data <= 0).any():
                    messagebox.showerror("數據錯誤", "對數模型要求所有 X 值都大於 0。", parent=self)
                    return
                popt, _ = curve_fit(lambda x, a, b: a * np.log(x) + b, x_data, y_data)
                y_pred = popt[0] * np.log(x_data) + popt[1]
                equation = f"y = {popt[0]:.4f} * log(x) + {popt[1]:.4f}"

            r2 = r2_score(y_data, y_pred)
            
            x_fit = np.linspace(x_data.min(), x_data.max(), 100)
            if model_type == "線性":
                y_fit = popt[0] * x_fit + popt[1]
            elif model_type == "多項式":
                y_fit = poly_func(x_fit)
            elif model_type == "指數":
                y_fit = popt[0] * np.exp(popt[1] * x_fit) + popt[2]
            elif model_type == "對數":
                if (x_fit <= 0).any():
                    x_fit = np.linspace(max(x_data.min(), 1e-9), x_data.max(), 100)
                y_fit = popt[0] * np.log(x_fit) + popt[1]

            self.ax.plot(x_fit, y_fit, color='red', label=f'{model_type}擬合')

        except Exception as e:
            messagebox.showerror("分析失敗", f"無法擬合模型：\n{e}", parent=self)
            return

        self.ax.set_xlabel(x_col)
        self.ax.set_ylabel(y_col)
        self.ax.set_title(f"{y_col} vs. {x_col} ({model_type}分析)")
        self.ax.legend()
        self.ax.grid(True)
        self.canvas.draw()

        # 更新結論
        conclusion = (
            f"模型擬合結果 ({model_type}):\n"
            f"  - 擬合方程式: {equation}\n"
            f"  - R² (決定係數): {r2:.4f}\n"
            f"通用相關性分析:\n"
            f"  - 皮爾森相關係數 (線性): {pearson_corr:.4f} (p-value: {pearson_p:.4f})\n"
            f"  - 斯皮爾曼相關係數 (等級): {spearman_corr:.4f} (p-value: {spearman_p:.4f})"
        )
        self.conclusion_text.config(state='normal')
        self.conclusion_text.delete(1.0, tk.END)
        self.conclusion_text.insert(tk.END, conclusion)
        self.conclusion_text.config(state='disabled')

class ChartWindow(tk.Toplevel):
    """
    一個用於顯示和編輯圖表的獨立視窗。
    """
    def __init__(self, parent, df, chart_type, x_col=None, y_col=None):
        super().__init__(parent)
        self.parent = parent
        self.df = df.copy()
        self.x_col = x_col
        self.y_col = y_col
        self.chart_type = chart_type

        self.title(f"{chart_type.capitalize()} 圖表")
        self.geometry("800x600")

        self._create_menu()
        self._draw_chart()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _create_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="編輯", menu=edit_menu)
        edit_menu.add_command(label="編輯標題...", command=self._edit_title)
        if self.x_col:
            edit_menu.add_command(label="編輯 X 軸名稱...", command=self._edit_xlabel)
        if self.y_col:
            edit_menu.add_command(label="編輯 Y 軸名稱...", command=self._edit_ylabel)

    def _draw_chart(self):
        """繪製所選類型的圖表。"""
        cols_to_check = []
        if self.x_col:
            cols_to_check.append(self.x_col)
            self.df[self.x_col] = pd.to_numeric(self.df[self.x_col], errors='coerce')
        if self.y_col:
            cols_to_check.append(self.y_col)
            self.df[self.y_col] = pd.to_numeric(self.df[self.y_col], errors='coerce')
        
        plot_df = self.df.dropna(subset=cols_to_check)

        if plot_df.empty:
            messagebox.showerror("數據錯誤", "所選列沒有足夠的有效數值數據來繪製圖表。", parent=self)
            self.destroy()
            return

        self.fig, self.ax = plt.subplots(figsize=(7, 5), dpi=100)

        try:
            if self.chart_type == 'scatter':
                sns.scatterplot(data=plot_df, x=self.x_col, y=self.y_col, ax=self.ax)
                self.ax.set_title(f"{self.y_col} vs. {self.x_col}")
            elif self.chart_type == 'line':
                sns.lineplot(data=plot_df, x=self.x_col, y=self.y_col, ax=self.ax)
                self.ax.set_title(f"{self.y_col} vs. {self.x_col}")
            elif self.chart_type == 'bar':
                sns.barplot(data=plot_df, x=self.x_col, y=self.y_col, ax=self.ax)
                self.ax.set_title(f"{self.y_col} vs. {self.x_col}")
            elif self.chart_type == 'hist':
                sns.histplot(data=plot_df, x=self.x_col, ax=self.ax, kde=True)
                self.ax.set_title(f"{self.x_col} 的直方圖")
            elif self.chart_type == 'box':
                sns.boxplot(data=plot_df, y=self.y_col, ax=self.ax)
                self.ax.set_title(f"{self.y_col} 的箱形圖")
            elif self.chart_type == 'kde':
                sns.kdeplot(data=plot_df, x=self.x_col, ax=self.ax, fill=True)
                self.ax.set_title(f"{self.x_col} 的核密度估計圖")

        except Exception as e:
            messagebox.showerror("繪圖錯誤", f"無法生成圖表：\n{e}", parent=self)
            self.destroy()
            return

        if self.x_col: self.ax.set_xlabel(self.x_col)
        if self.y_col: self.ax.set_ylabel(self.y_col)
        
        plt.tight_layout()

        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        toolbar = NavigationToolbar2Tk(self.canvas, self)
        toolbar.update()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def _redraw_canvas(self):
        self.canvas.draw()

    def _edit_title(self):
        new_title = simpledialog.askstring("編輯標題", "輸入新的圖表標題:", parent=self)
        if new_title is not None:
            self.ax.set_title(new_title)
            self._redraw_canvas()

    def _edit_xlabel(self):
        new_label = simpledialog.askstring("編輯 X 軸", "輸入新的 X 軸名稱:", parent=self)
        if new_label is not None:
            self.ax.set_xlabel(new_label)
            self._redraw_canvas()

    def _edit_ylabel(self):
        new_label = simpledialog.askstring("編輯 Y 軸", "輸入新的 Y 軸名稱:", parent=self)
        if new_label is not None:
            self.ax.set_ylabel(new_label)
            self._redraw_canvas()

class VisualizationDialog(tk.Toplevel):
    """
    一個用於選擇圖表類型和數據列的對話框。
    """
    def __init__(self, parent, columns):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        self.columns = columns
        self.title("創建圖表")
        self.geometry("350x200")

        self._create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="圖表類型:").grid(row=0, column=0, sticky="w", pady=5)
        chart_types = ['scatter', 'line', 'bar', 'hist', 'box', 'kde']
        self.chart_type_var = tk.StringVar(value='scatter')
        self.chart_combo = ttk.Combobox(main_frame, textvariable=self.chart_type_var, values=chart_types, state='readonly')
        self.chart_combo.grid(row=0, column=1, sticky="ew")
        self.chart_combo.bind("<<ComboboxSelected>>", self.on_chart_type_select)

        ttk.Label(main_frame, text="X 軸:").grid(row=1, column=0, sticky="w", pady=5)
        self.x_col_var = tk.StringVar()
        self.x_combo = ttk.Combobox(main_frame, textvariable=self.x_col_var, values=self.columns, state='readonly')
        self.x_combo.grid(row=1, column=1, sticky="ew")

        ttk.Label(main_frame, text="Y 軸:").grid(row=2, column=0, sticky="w", pady=5)
        self.y_col_var = tk.StringVar()
        self.y_combo = ttk.Combobox(main_frame, textvariable=self.y_col_var, values=self.columns, state='readonly')
        self.y_combo.grid(row=2, column=1, sticky="ew")

        ttk.Button(main_frame, text="生成圖表", command=self.generate_chart).grid(row=3, column=1, sticky="e", pady=20)

    def on_chart_type_select(self, event=None):
        """根據選擇的圖表類型，啟用/禁用軸選擇。"""
        chart_type = self.chart_type_var.get()
        if chart_type in ['hist', 'kde']:
            self.x_combo.config(state='readonly')
            self.y_combo.config(state='disabled')
            self.y_col_var.set('')
        elif chart_type == 'box':
            self.x_combo.config(state='disabled')
            self.y_combo.config(state='readonly')
            self.x_col_var.set('')
        else: # scatter, line, bar
            self.x_combo.config(state='readonly')
            self.y_combo.config(state='readonly')

    def generate_chart(self):
        x_col = self.x_col_var.get()
        y_col = self.y_col_var.get()
        chart_type = self.chart_type_var.get()

        if chart_type in ['scatter', 'line', 'bar'] and (not x_col or not y_col):
            messagebox.showwarning("選擇不完整", "此圖表類型需要同時選擇 X 軸和 Y 軸。", parent=self)
            return
        if chart_type in ['hist', 'kde'] and not x_col:
            messagebox.showwarning("選擇不完整", "此圖表類型需要選擇 X 軸。", parent=self)
            return
        if chart_type == 'box' and not y_col:
            messagebox.showwarning("選擇不完整", "此圖表類型需要選擇 Y 軸。", parent=self)
            return

        self.destroy()
        self.parent.create_chart_window(chart_type, x_col or None, y_col or None)

class PatternFillDialog(tk.Toplevel):
    """
    一個用於範式填充功能的獨立對話框。
    """
    def __init__(self, parent, start_value, start_index, start_column, all_indices, all_columns):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        
        self.start_value = str(start_value)
        self.start_index = start_index
        self.start_column = start_column
        self.all_indices = all_indices
        self.all_columns = all_columns
        
        self.title("範式填充")
        self.geometry("800x600")
        self.resizable(True, True)

        self._create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _create_widgets(self):
        dialog_main_frame = ttk.Frame(self)
        dialog_main_frame.pack(fill=tk.BOTH, expand=True)

        button_frame = ttk.Frame(dialog_main_frame, padding=(10, 10, 10, 10))
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(button_frame, text="應用", command=self._apply_fill).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="取消", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="預覽", command=self._update_preview).pack(side=tk.RIGHT, padx=5)

        main_pane = ttk.PanedWindow(dialog_main_frame, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

        controls_frame = ttk.Frame(main_pane, padding=10)
        main_pane.add(controls_frame, weight=2) 

        preview_container = ttk.Frame(main_pane, padding=10)
        main_pane.add(preview_container, weight=3)

        self._populate_controls_frame(controls_frame)
        self._populate_preview_frame(preview_container)


    def _populate_controls_frame(self, parent_frame):
        common_frame = ttk.LabelFrame(parent_frame, text="通用設置", padding=10)
        common_frame.pack(fill=tk.X, pady=(0, 5), side=tk.TOP)

        ttk.Label(common_frame, text="當前起始值:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.start_value_label = ttk.Label(common_frame, text=f"'{self.start_value}'", font=('Microsoft YaHei UI', 10, 'bold'))
        self.start_value_label.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(common_frame, text="填充方向:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.fill_direction_var = tk.StringVar(value="column")
        dir_frame = ttk.Frame(common_frame)
        ttk.Radiobutton(dir_frame, text="填充列 (向下)", variable=self.fill_direction_var, value="column").pack(side=tk.LEFT)
        ttk.Radiobutton(dir_frame, text="填充行 (向右)", variable=self.fill_direction_var, value="row").pack(side=tk.LEFT, padx=10)
        dir_frame.grid(row=1, column=1, sticky="w")

        ttk.Label(common_frame, text="填充數量:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.fill_count_var = tk.StringVar(value="10")
        ttk.Entry(common_frame, textvariable=self.fill_count_var, width=10).grid(row=2, column=1, sticky="w", padx=5)

        manual_frame = ttk.LabelFrame(parent_frame, text="手動選擇起始單元格", padding=10)
        manual_frame.pack(fill=tk.X, pady=5, side=tk.TOP)
        
        ttk.Label(manual_frame, text="行頭:").grid(row=0, column=0, padx=5, pady=5)
        self.row_var = tk.StringVar(value=str(self.start_index))
        self.row_combo = ttk.Combobox(manual_frame, textvariable=self.row_var, values=self.all_indices)
        self.row_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(manual_frame, text="列頭:").grid(row=1, column=0, padx=5, pady=5)
        self.col_var = tk.StringVar(value=self.start_column)
        self.col_combo = ttk.Combobox(manual_frame, textvariable=self.col_var, values=self.all_columns)
        self.col_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Button(manual_frame, text="更新", command=self._update_start_value_from_selection).grid(row=2, column=1, padx=5, pady=10, sticky="e")
        
        manual_frame.grid_columnconfigure(1, weight=1)

        self.notebook = ttk.Notebook(parent_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10, side=tk.TOP)

        self.copy_frame = ttk.Frame(self.notebook, padding=10)
        self.sequence_frame = ttk.Frame(self.notebook, padding=10)
        self.cycle_frame = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.copy_frame, text="複製")
        self.notebook.add(self.sequence_frame, text="遞推")
        self.notebook.add(self.cycle_frame, text="循環")

        self._populate_pattern_frames()

    def _populate_preview_frame(self, parent_frame):
        preview_frame = ttk.LabelFrame(parent_frame, text="預覽", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        tree_container = ttk.Frame(preview_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        self.preview_tree = ttk.Treeview(tree_container, show="headings", height=5)
        
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.preview_tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.preview_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

    def _populate_pattern_frames(self):
        ttk.Label(self.copy_frame, text="將起始值複製到指定數量的單元格中。").pack()

        ttk.Label(self.sequence_frame, text="步長:").grid(row=0, column=0, sticky="w", pady=5)
        self.step_var = tk.StringVar(value="1")
        ttk.Entry(self.sequence_frame, textvariable=self.step_var, width=10).grid(row=0, column=1, sticky="w")

        ttk.Label(self.cycle_frame, text="步長:").grid(row=0, column=0, sticky="w", pady=5)
        self.cycle_step_var = tk.StringVar(value="1")
        ttk.Entry(self.cycle_frame, textvariable=self.cycle_step_var, width=10).grid(row=0, column=1, sticky="w")
        
        ttk.Label(self.cycle_frame, text="下限:").grid(row=1, column=0, sticky="w", pady=5)
        self.lower_bound_var = tk.StringVar(value="1")
        ttk.Entry(self.cycle_frame, textvariable=self.lower_bound_var, width=10).grid(row=1, column=1, sticky="w")

        ttk.Label(self.cycle_frame, text="上限:").grid(row=2, column=0, sticky="w", pady=5)
        self.upper_bound_var = tk.StringVar(value="10")
        ttk.Entry(self.cycle_frame, textvariable=self.upper_bound_var, width=10).grid(row=2, column=1, sticky="w")

    def _update_start_value_from_selection(self):
        """從手動選擇的單元格更新起始值。"""
        selected_row_str = self.row_var.get()
        selected_col = self.col_var.get()

        if not selected_row_str or not selected_col:
            messagebox.showwarning("選擇不完整", "請選擇行和列。", parent=self)
            return
        
        try:
            index_dtype = self.parent.dataframe.index.dtype
            df_index = pd.Series([selected_row_str]).astype(index_dtype).iloc[0]
            
            new_start_value = self.parent.dataframe.loc[df_index, selected_col]
            self.start_value = str(new_start_value)
            self.start_value_label.config(text=f"'{self.start_value}'")
            messagebox.showinfo("更新成功", f"起始值已更新為 '{self.start_value}'。", parent=self)
        except (KeyError, IndexError, ValueError, TypeError):
            messagebox.showerror("錯誤", "無法找到所選的單元格。", parent=self)

    def _generate_fill_data(self):
        """根據用戶選擇的模式和參數生成填充數據列表。"""
        try:
            fill_count = int(self.fill_count_var.get())
            if fill_count <= 0:
                messagebox.showerror("輸入錯誤", "填充數量必須是正整數。", parent=self)
                return None
        except ValueError:
            messagebox.showerror("輸入錯誤", "填充數量必須是有效的整數。", parent=self)
            return None

        active_tab_text = self.notebook.tab(self.notebook.select(), "text")
        
        match = re.match(r'^(.*?)(\d+)$', self.start_value)
        if match:
            prefix, num_str = match.groups()
            current_num = int(num_str)
        else:
            prefix, current_num = self.start_value, None
            if active_tab_text != "複製":
                try:
                    current_num = int(self.start_value)
                    prefix = ""
                except (ValueError, TypeError):
                    messagebox.showerror("格式錯誤", f"起始值 '{self.start_value}' 無法用於遞推或循環模式。\n請使用純數字或 '文字+數字' 格式。", parent=self)
                    return None

        result_list = []
        if active_tab_text == "複製":
            result_list = [self.start_value] * fill_count
        
        elif active_tab_text == "遞推":
            try:
                step = int(self.step_var.get())
            except ValueError:
                messagebox.showerror("輸入錯誤", "步長必須是有效的整數。", parent=self)
                return None
            
            for _ in range(fill_count):
                result_list.append(f"{prefix}{current_num}")
                current_num += step

        elif active_tab_text == "循環":
            try:
                step = int(self.cycle_step_var.get())
                lower = int(self.lower_bound_var.get())
                upper = int(self.upper_bound_var.get())
            except ValueError:
                messagebox.showerror("輸入錯誤", "步長和循環範圍必須是有效的整數。", parent=self)
                return None

            if lower >= upper:
                messagebox.showerror("範圍錯誤", "循環範圍的上限必須大於下限。", parent=self)
                return None
            if current_num is None or not (lower <= current_num <= upper):
                messagebox.showerror("範圍錯誤", f"起始值 '{current_num}' 不在循環範圍 [{lower}, {upper}] 內。", parent=self)
                return None
            if (upper - lower + 1) <= abs(step):
                messagebox.showerror("步長錯誤", "循環範圍的區間必須大於步長的絕對值。", parent=self)
                return None

            for _ in range(fill_count):
                result_list.append(f"{prefix}{current_num}")
                current_num += step
                if current_num > upper:
                    current_num = lower + (current_num - upper - 1) % (upper - lower + 1)
                elif current_num < lower:
                    current_num = upper - (lower - current_num - 1) % (upper - lower + 1)

        return result_list

    def _update_preview(self):
        """更新預覽表格。"""
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        data = self._generate_fill_data()
        if data is None: return

        direction = self.fill_direction_var.get()
        if direction == 'column':
            self.preview_tree['columns'] = ('填充值',)
            self.preview_tree.heading('填充值', text='填充值')
            for val in data:
                self.preview_tree.insert('', 'end', values=(val,))
        else: # row
            self.preview_tree['columns'] = [f'單元格 {i+1}' for i in range(len(data))]
            for i, col_name in enumerate(self.preview_tree['columns']):
                self.preview_tree.heading(col_name, text=col_name)
            self.preview_tree.insert('', 'end', values=data)

    def _apply_fill(self):
        """應用填充並關閉對話框。"""
        data = self._generate_fill_data()
        if data is None: return
        
        direction = self.fill_direction_var.get()
        self.parent.apply_pattern_fill(direction, data)
        self.destroy()

class PySheetApp(tk.Tk):
    """
    一個使用 Python 和 Tkinter 開發的專業版表格編輯器應用程式。
    所有查找功能均已整合至主 UI。
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
        
        self.find_replace_frame_visible = False
        self.find_header_frame_visible = False
        self.last_find_coords = (-1, -1)
        self.last_header_find_index = -1
        self.last_header_search_term = ""
        self.last_header_search_type = ""
        
        self.active_column_id = None

        self._create_widgets()
        self._create_menu()
        self._create_context_menu()
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _setup_styles(self):
        self.bg_color = '#e0e0e0'
        self.fg_color = '#212121'
        self.entry_bg = '#ffffff'
        self.tree_bg_even = '#ffffff'
        self.tree_bg_odd = '#f5f5f5'
        self.select_bg = '#d9e8f9' 
        self.configure(bg=self.bg_color)
        self.style.configure('.', background=self.bg_color, foreground=self.fg_color, font=('Microsoft YaHei UI', 10))
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground=self.fg_color, padding=5)
        self.style.configure('TEntry', fieldbackground=self.entry_bg)
        self.style.configure('Treeview.Heading', font=('Microsoft YaHei UI', 10, 'bold'), padding=5)
        self.style.configure('Treeview', rowheight=25, fieldbackground=self.tree_bg_even)
        self.style.map('Treeview', 
                       background=[('selected', self.select_bg)],
                       foreground=[('selected', self.fg_color)])

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 5))
        self.cell_pos_label = ttk.Label(top_frame, text="儲存格:", font=('Microsoft YaHei UI', 10))
        self.cell_pos_label.pack(side=tk.LEFT, padx=(0, 5), anchor='w')
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(top_frame, textvariable=self.input_var, font=('Microsoft YaHei UI', 12))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.input_entry.bind("<Return>", self._update_cell_from_input)
        self.input_entry.bind("<FocusOut>", self._update_cell_from_input)

        self._create_find_replace_frame(main_frame)
        self._create_find_header_frame(main_frame)

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
        
        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self.tree.bind("<Button-1>", self._on_cell_click)
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Up>", self._on_key_navigate)
        self.tree.bind("<Down>", self._on_key_navigate)
        self.tree.bind("<Left>", self._on_key_navigate)
        self.tree.bind("<Right>", self._on_key_navigate)

    def _create_find_replace_frame(self, parent):
        self.find_replace_frame = ttk.Frame(parent, padding=5)
        ttk.Label(self.find_replace_frame, text="查找:").pack(side=tk.LEFT, padx=(0, 2))
        self.find_var = tk.StringVar()
        self.find_entry = ttk.Entry(self.find_replace_frame, textvariable=self.find_var, width=20)
        self.find_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(self.find_replace_frame, text="替換為:").pack(side=tk.LEFT, padx=(5, 2))
        self.replace_var = tk.StringVar()
        ttk.Entry(self.find_replace_frame, textvariable=self.replace_var, width=20).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(self.find_replace_frame, text="下一個", command=self.find_next).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.find_replace_frame, text="替換", command=self.replace_cell).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.find_replace_frame, text="全部替換", command=self.replace_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.find_replace_frame, text="×", command=self.toggle_find_replace_frame, width=3).pack(side=tk.RIGHT, padx=(10, 0))

    def _create_find_header_frame(self, parent):
        self.find_header_frame = ttk.Frame(parent, padding=5)
        ttk.Label(self.find_header_frame, text="查找標題:").pack(side=tk.LEFT, padx=(0, 2))
        self.header_search_var = tk.StringVar()
        self.header_search_entry = ttk.Entry(self.find_header_frame, textvariable=self.header_search_var, width=20)
        self.header_search_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.header_search_type_var = tk.StringVar(value="column")
        ttk.Radiobutton(self.find_header_frame, text="列頭", variable=self.header_search_type_var, value="column").pack(side=tk.LEFT)
        ttk.Radiobutton(self.find_header_frame, text="行頭", variable=self.header_search_type_var, value="row").pack(side=tk.LEFT, padx=(0, 10))
        self.header_fuzzy_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.find_header_frame, text="模糊搜索", variable=self.header_fuzzy_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(self.find_header_frame, text="查找下一個", command=self.find_next_header).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.find_header_frame, text="×", command=self.toggle_find_header_frame, width=3).pack(side=tk.RIGHT, padx=(10, 0))

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
        file_menu.add_command(label="導出為圖片...", command=self.export_as_image)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self._on_closing)
        
        self.edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="編輯", menu=self.edit_menu)
        self.edit_menu.add_command(label="回滾", command=self.undo_action, accelerator="Ctrl+Z", state="disabled")
        self.edit_menu.add_command(label="前滾", command=self.redo_action, accelerator="Ctrl+Y", state="disabled")
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="剪下", command=self._cut_cell, accelerator="Ctrl+X")
        self.edit_menu.add_command(label="複製", command=self._copy_cell, accelerator="Ctrl+C")
        self.edit_menu.add_command(label="貼上", command=self._paste_cell, accelerator="Ctrl+V")
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="範式填充...", command=self.open_pattern_fill_dialog)
        self.edit_menu.add_command(label="擴展表格...", command=self.open_extend_dialog)
        
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="查看", menu=view_menu)
        view_menu.add_command(label="查找與替換...", command=self.toggle_find_replace_frame, accelerator="Ctrl+F")
        view_menu.add_command(label="查找行頭/列頭...", command=self.toggle_find_header_frame)

        analysis_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="分析", menu=analysis_menu)
        analysis_menu.add_command(label="關係分析...", command=self.open_analysis_dialog)

        vis_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="可視化", menu=vis_menu)
        vis_menu.add_command(label="繪製圖表...", command=self.open_visualization_dialog)

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
        self.context_menu.add_separator()
        self.context_menu.add_command(label="重命名行頭...", command=self._rename_row_header)

    def _on_closing(self):
        """處理窗口關閉事件以確保乾淨退出。"""
        self.destroy()

    def _normalize_headers(self):
        """
        將行和列標題重置為標準的順序格式。
        - 行索引重置為從 0 開始的 RangeIndex (0, 1, 2, ...)。
        - 列名重置為 Excel 風格的名稱 (A, B, C, ...)。
        此操作可確保在修改後數據的完整性並防止定位錯誤。
        原始數據內容將被保留。
        """
        if self.dataframe.empty:
            return

        old_values = self.dataframe.values
        num_cols = self.dataframe.shape[1]
        new_col_names = [_to_excel_col(i + 1) for i in range(num_cols)]
        self.dataframe = pd.DataFrame(old_values, columns=new_col_names)

    # --- 標題編輯 ---
    def _on_double_click(self, event):
        """根據雙擊的位置，分派到標題編輯或儲存格編輯。"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            self._edit_column_header_popup(event)
        elif region == "cell":
            self._edit_cell_popup(event)

    def _edit_column_header_popup(self, event):
        """處理雙擊欄位標題以進行重命名的操作。"""
        messagebox.showinfo("功能變更", "為了確保數據穩定性，欄位標題現在由系統自動管理 (A, B, C...)，不支持手動重命名。", parent=self)
        return

    def _edit_cell_popup(self, event):
        """處理雙擊儲存格以彈出對話框進行編輯的操作。"""
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or not column_id: return

        try:
            col_index = int(column_id.replace('#', '')) - 1
            if col_index < 0: return
            column_name = self.tree['columns'][col_index]
            
            # Since index is always 0,1,2... we can convert item_id directly
            df_index = int(item_id)

            old_value = self.dataframe.loc[df_index, column_name]
        except (KeyError, IndexError, ValueError, TypeError):
            return

        new_value = simpledialog.askstring("編輯儲存格",
                                           f"為儲存格 {column_name}[{df_index}] 輸入新值:",
                                           initialvalue=old_value,
                                           parent=self)

        if new_value is not None and new_value != str(old_value):
            self._add_to_undo(('編輯儲存格', self.dataframe.copy()))
            self._apply_change(df_index, column_name, new_value)

    def _rename_row_header(self):
        """處理通過右鍵選單重命名行頭的邏輯。"""
        messagebox.showinfo("功能變更", "為了確保數據穩定性，行號現在由系統自動管理 (0, 1, 2...)，不支持手動重命名。", parent=self)
        return

    def open_analysis_dialog(self):
        if self.dataframe.empty or self.dataframe.shape[1] < 2:
            messagebox.showwarning("數據不足", "需要至少兩列數據才能進行分析。", parent=self)
            return
        AnalysisDialog(self, self.dataframe)

    # --- 可視化與查找邏輯 ---
    def open_visualization_dialog(self):
        if self.dataframe.empty:
            messagebox.showwarning("無數據", "請先加載或創建數據。")
            return
        # When opening visualization, use the current (potentially normalized) headers
        VisualizationDialog(self, list(self.dataframe.columns))

    def create_chart_window(self, chart_type, x_col, y_col):
        ChartWindow(self, self.dataframe, chart_type, x_col, y_col)

    def open_pattern_fill_dialog(self):
        """打開範式填充對話框。"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("未選擇", "請先選擇一個起始單元格。")
            return
        
        item_id = selected_items[0]
        if not self.active_column_id:
            messagebox.showwarning("未選擇", "無法確定起始列。請點擊一個單元格。")
            return
            
        try:
            col_index = int(self.active_column_id.replace('#', '')) - 1
            column_name = self.tree['columns'][col_index]
            df_index = int(item_id)
            start_value = self.dataframe.loc[df_index, column_name]
            
            all_columns = list(self.dataframe.columns)
            all_indices = list(self.dataframe.index.map(str))

            PatternFillDialog(self, start_value, df_index, column_name, all_indices, all_columns)
        except (KeyError, IndexError, ValueError, TypeError):
            messagebox.showerror("錯誤", "無法獲取所選單元格的數據。")
            return

    def apply_pattern_fill(self, direction, data):
        """應用範式填充的結果到 DataFrame，並在需要時自動擴展。"""
        selected_items = self.tree.selection()
        if not selected_items or not self.active_column_id: return
        
        item_id = selected_items[0]
        try:
            col_index = int(self.active_column_id.replace('#', '')) - 1
            start_col_name = self.tree['columns'][col_index]
            start_df_index = int(item_id)
            start_row_pos = self.dataframe.index.get_loc(start_df_index)
            start_col_pos = self.dataframe.columns.get_loc(start_col_name)
        except (KeyError, IndexError, ValueError, TypeError):
            return
        
        self._add_to_undo(('範式填充', self.dataframe.copy()))

        if direction == 'column':
            # Fill down
            end_row_pos = start_row_pos + len(data)
            self.dataframe.iloc[start_row_pos:end_row_pos, start_col_pos] = data
        else: # row
            # Fill right
            end_col_pos = start_col_pos + len(data)
            self.dataframe.iloc[start_row_pos, start_col_pos:end_col_pos] = data
        
        self._normalize_headers()
        self._load_data_to_treeview()

    def open_extend_dialog(self):
        ExtendDialog(self)

    def execute_extend(self, count, direction):
        if self.dataframe.empty:
            if direction == 'down':
                self.dataframe = pd.DataFrame(np.full((count, 1), ""))
            else: # right
                self.dataframe = pd.DataFrame(np.full((1, count), ""))
        else:
            self._add_to_undo(('擴展表格', self.dataframe.copy()))
            if direction == 'down':
                new_rows = pd.DataFrame(np.full((count, self.dataframe.shape[1]), ""), columns=self.dataframe.columns)
                self.dataframe = pd.concat([self.dataframe, new_rows])
            else: # right
                for i in range(count):
                    self.dataframe[f'__TEMP__{i}'] = ""
        
        self._normalize_headers()
        self._load_data_to_treeview()

    def toggle_find_replace_frame(self):
        if self.find_header_frame_visible: self.toggle_find_header_frame()
        if self.find_replace_frame_visible:
            self.find_replace_frame.pack_forget()
        else:
            self.find_replace_frame.pack(fill=tk.X, pady=(0, 5), before=self.tree.master)
            self.find_entry.focus_set()
        self.find_replace_frame_visible = not self.find_replace_frame_visible

    def toggle_find_header_frame(self):
        if self.find_replace_frame_visible: self.toggle_find_replace_frame()
        if self.find_header_frame_visible:
            self.find_header_frame.pack_forget()
        else:
            self.find_header_frame.pack(fill=tk.X, pady=(0, 5), before=self.tree.master)
            self.header_search_entry.focus_set()
        self.find_header_frame_visible = not self.find_header_frame_visible

    def find_next(self):
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

    def find_next_header(self):
        term = self.header_search_var.get()
        search_type = self.header_search_type_var.get()
        if not term: return
        if term != self.last_header_search_term or search_type != self.last_header_search_type:
            self.last_header_find_index = -1
        self.last_header_search_term = term
        self.last_header_search_type = search_type
        start_index = self.last_header_find_index + 1
        headers = list(self.dataframe.columns) if search_type == 'column' else list(self.dataframe.index.map(str))
        if start_index >= len(headers):
            if messagebox.askyesno("查找完畢", "已搜索到結尾。是否從頭開始搜索？", parent=self):
                start_index = 0
            else:
                return
        for i in range(start_index, len(headers)):
            header = headers[i]
            match = False
            if self.header_fuzzy_var.get():
                if process.fuzz.ratio(term.lower(), str(header).lower()) > 70: match = True
            else:
                if term.lower() in str(header).lower(): match = True
            if match:
                self.last_header_find_index = i
                if search_type == 'column': self.highlight_column(i)
                else: self.highlight_row(i)
                return
        self.last_header_find_index = -1
        messagebox.showinfo("未找到", f"找不到匹配 '{term}' 的標題。", parent=self)

    def replace_cell(self):
        selected_items = self.tree.selection()
        if not selected_items: self.find_next(); return
        row_id_str, row_idx, col_idx = selected_items[0], *self.last_find_coords
        if self.dataframe.index.get_loc(int(row_id_str)) != row_idx or col_idx == -1:
            self.find_next(); return
        find_term, replace_term = self.find_var.get(), self.replace_var.get()
        col_name = self.dataframe.columns[col_idx]
        old_value = str(self.dataframe.iat[row_idx, col_idx])
        new_value = old_value.replace(find_term, replace_term, 1)
        self._add_to_undo(('edit', self.dataframe.copy()))
        self._apply_change(int(row_id_str), col_name, new_value)
        self.find_next()

    def replace_all(self):
        find_term, replace_term = self.find_var.get(), self.replace_var.get()
        if not find_term: return
        self._add_to_undo(('replace_all', self.dataframe.copy()))
        new_df = self.dataframe.astype(str).applymap(lambda x: x.replace(find_term, replace_term))
        found_count = (self.dataframe.astype(str) != new_df.astype(str)).sum().sum()
        if found_count > 0:
            self.dataframe = new_df
            self._load_data_to_treeview()
            messagebox.showinfo("完成", f"已完成 {found_count} 處替換。", parent=self)
        else:
            self.undo_stack.pop() # No change, so remove undo record
            messagebox.showinfo("未找到", f"找不到 '{find_term}'。", parent=self)

    def highlight_cell(self, row_idx, col_idx):
        self.last_find_coords = (row_idx, col_idx)
        row_id_str = str(self.dataframe.index[row_idx])
        self.tree.selection_set(row_id_str)
        self.tree.focus(row_id_str)
        self.tree.see(row_id_str)

    def highlight_column(self, col_idx):
        col_id = f"#{col_idx + 1}"
        self.tree.xview_moveto(col_idx / len(self.dataframe.columns))
        all_rows = self.tree.get_children()
        if all_rows:
            first_row_id = all_rows[0]
            self.tree.selection_set(first_row_id)
            self.tree.focus(first_row_id)
            col_name = self.dataframe.columns[col_idx]
            self.cell_pos_label.config(text=f"欄: {col_name}")

    def highlight_row(self, row_idx):
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
        description, old_dataframe = self.undo_stack.pop()
        self.redo_stack.append((description, self.dataframe.copy()))
        self.dataframe = old_dataframe
        self._load_data_to_treeview()
        self._update_edit_menu_state()

    def redo_action(self, event=None):
        if not self.redo_stack: return
        description, new_dataframe = self.redo_stack.pop()
        self.undo_stack.append((description, self.dataframe.copy()))
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
        df_index = int(row_id)
        self.clipboard_data = self.dataframe.loc[df_index, col_name]
        self._update_edit_menu_state()

    def _cut_cell(self, event=None):
        row_id, col_name = self._get_context_cell()
        if row_id is None: return
        self._copy_cell()
        df_index = int(row_id)
        self._add_to_undo(('剪下', self.dataframe.copy()))
        self._apply_change(df_index, col_name, "")

    def _paste_cell(self, event=None):
        if self.clipboard_data is None: return
        row_id, col_name = self._get_context_cell()
        if row_id is None: return
        df_index = int(row_id)
        self._add_to_undo(('貼上', self.dataframe.copy()))
        self._apply_change(df_index, col_name, self.clipboard_data)

    def _delete_row(self):
        if not self.context_menu_row_id: return
        self._add_to_undo(('刪除行', self.dataframe.copy()))
        try:
            df_index_to_drop = int(self.context_menu_row_id)
            self.dataframe = self.dataframe.drop(df_index_to_drop)
        except (ValueError, KeyError):
            return
        self._normalize_headers()
        self._load_data_to_treeview()

    def _insert_row(self, above=True):
        self._add_to_undo(('插入行', self.dataframe.copy()))
        insert_pos = 0
        if self.dataframe.empty:
            self.dataframe = pd.DataFrame([[""]])
            self._normalize_headers()
            self._load_data_to_treeview()
            return
            
        if self.context_menu_row_id:
            try:
                row_pos = self.tree.index(self.context_menu_row_id)
                insert_pos = row_pos if above else row_pos + 1
            except tk.TclError:
                insert_pos = len(self.dataframe)
        else:
            insert_pos = len(self.dataframe)

        new_row_df = pd.DataFrame([[""] * len(self.dataframe.columns)], columns=self.dataframe.columns)
        part1 = self.dataframe.iloc[:insert_pos]
        part2 = self.dataframe.iloc[insert_pos:]
        self.dataframe = pd.concat([part1, new_row_df, part2])
        self._normalize_headers()
        self._load_data_to_treeview()

    def _delete_column(self):
        _, col_name = self._get_context_cell()
        if col_name is None: return
        if messagebox.askyesno("確認刪除", f"您確定要刪除 '{col_name}' 這一整列嗎？"):
            self._add_to_undo(('刪除列', self.dataframe.copy()))
            self.dataframe.drop(columns=[col_name], inplace=True)
            self._normalize_headers()
            self._load_data_to_treeview()

    def _insert_column(self, left=True):
        if self.dataframe.empty:
            self.dataframe = pd.DataFrame([[""]])
            self._normalize_headers()
            self._load_data_to_treeview()
            return

        self._add_to_undo(('插入列', self.dataframe.copy()))
        insert_pos = 0
        if self.context_menu_col_id:
            try:
                col_index = int(self.context_menu_col_id.replace('#', '')) - 1
                if col_index >= 0:
                    insert_pos = col_index if left else col_index + 1
            except (ValueError, IndexError):
                insert_pos = len(self.dataframe.columns)
        else:
            insert_pos = len(self.dataframe.columns)
            
        self.dataframe.insert(insert_pos, f"__TEMP__{pd.Timestamp.now().isoformat()}", "")
        self._normalize_headers()
        self._load_data_to_treeview()

    def _apply_change(self, df_index, col_name, value):
        self.dataframe.loc[df_index, col_name] = value
        item_id = str(df_index)
        updated_row_values = [str(v) for v in self.dataframe.loc[df_index].values]
        try:
            self.tree.item(item_id, values=updated_row_values)
        except tk.TclError:
            pass 

    def _update_cell_from_input(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items: return
        item_id = selected_items[0]
        try:
            if "儲存格:" in self.cell_pos_label.cget("text") and self.active_column_id:
                column_name = self.tree.heading(self.active_column_id, "text")
            else: return
            df_index = int(item_id)
        except (IndexError, KeyError, ValueError, TypeError): return
        new_value = self.input_var.get()
        self._add_to_undo(('edit', self.dataframe.copy()))
        self._apply_change(df_index, column_name, new_value)
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
        for i, (index, row) in enumerate(self.dataframe.iterrows()):
            values = [str(v) for v in row.values]
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", values=values, iid=str(index), tags=(tag,))

    def new_file(self, event=None):
        self.dataframe = pd.DataFrame([['', ''], ['', '']])
        self._normalize_headers()
        self.file_path = None
        self.title("Python 表格編輯器 (PySheet) - 未命名")
        self._load_data_to_treeview()
        self.undo_stack.clear(); self.redo_stack.clear(); self._update_edit_menu_state()

    def open_file(self, event=None):
        path = filedialog.askopenfilename(filetypes=[("支援的檔案", "*.xlsx *.csv *.json *.md"), ("Excel 檔案", "*.xlsx"), ("CSV 檔案", "*.csv"), ("JSON 檔案", "*.json"), ("Markdown 檔案", "*.md"), ("所有檔案", "*.*")])
        if not path: return
        self.file_path = path
        try:
            if path.endswith('.xlsx'): self.dataframe = pd.read_excel(path, index_col=0)
            elif path.endswith('.csv'): self.dataframe = pd.read_csv(path, index_col=0)
            elif path.endswith('.json'):
                try: self.dataframe = pd.read_json(path, orient='records')
                except ValueError:
                    with open(path, 'r', encoding='utf-8') as f: data = pd.read_json(f); self.dataframe = pd.json_normalize(data)
            elif path.endswith('.md'):
                self.dataframe = self._read_markdown_table(path)

            for col in self.dataframe.columns: self.dataframe[col] = self.dataframe[col].astype(str)
            self.title(f"Python 表格編輯器 - {path.split('/')[-1]}")
            self._normalize_headers()
            self._load_data_to_treeview()
            self.undo_stack.clear(); self.redo_stack.clear(); self._update_edit_menu_state()
        except Exception as e: messagebox.showerror("開啟檔案錯誤", f"無法讀取檔案：\n{e}"); self.dataframe = pd.DataFrame(); self._clear_treeview()

    def _read_markdown_table(self, path):
        """從 Markdown 文件中解析第一個表格。"""
        with open(path, 'r', encoding='utf-8') as f:
            lines = [line.strip() for line in f if line.strip()]
        
        if len(lines) < 2: raise ValueError("Markdown 文件不包含有效的表格。")

        separator_index = -1
        for i, line in enumerate(lines):
            if re.match(r'^[|: -]+$', line) and '---' in line:
                separator_index = i
                break
        
        if separator_index == -1 or separator_index == 0:
            raise ValueError("找不到 Markdown 表格的標頭分隔符。")

        header_line = lines[separator_index - 1]
        columns = [h.strip() for h in header_line.split('|') if h.strip()]
        
        data = []
        for line in lines[separator_index + 1:]:
            if re.match(r'^[|: -]+$', line) or not line.startswith('|'): break
            row = [d.strip() for d in line.split('|')]
            if len(row) > 1 and row[0] == '' and row[-1] == '':
                row = row[1:-1]
            if len(row) == len(columns):
                data.append(row)

        if not data: return pd.DataFrame(columns=columns)
        
        df = pd.DataFrame(data, columns=columns)
        df.set_index(df.columns[0], inplace=True)
        return df

    def save_file(self, event=None):
        if self.file_path: self._execute_save(self.file_path)
        else: self.save_file_as()

    def save_file_as(self, event=None):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 檔案", "*.xlsx"), ("CSV 檔案", "*.csv"), ("JSON 檔案", "*.json"), ("Markdown 檔案", "*.md")])
        if path: self._execute_save(path)

    def _execute_save(self, path):
        if self.dataframe.empty: messagebox.showwarning("無資料", "表格中沒有資料可以儲存。"); return
        try:
            df_to_save = self.dataframe.copy()
            for col in df_to_save.columns: df_to_save[col] = pd.to_numeric(df_to_save[col], errors='ignore')
            if path.endswith('.xlsx'): df_to_save.to_excel(path, index=True)
            elif path.endswith('.csv'): df_to_save.to_csv(path, index=True, encoding='utf-8-sig')
            elif path.endswith('.json'): df_to_save.to_json(path, orient='index', indent=4)
            elif path.endswith('.md'):
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(df_to_save.to_markdown())
            self.file_path = path; self.title(f"Python 表格編輯器 - {path.split('/')[-1]}"); messagebox.showinfo("儲存成功", f"檔案已成功儲存至：\n{path}")
        except Exception as e: messagebox.showerror("儲存檔案錯誤", f"無法儲存檔案：\n{e}")
    
    def export_as_image(self):
        """將當前表格導出為圖片。"""
        if self.dataframe.empty:
            messagebox.showwarning("無數據", "沒有數據可以導出。")
            return
            
        dpi = simpledialog.askinteger("設置 DPI", "請輸入導出圖片的 DPI (例如 150):", initialvalue=150, minvalue=50, maxvalue=600)
        if not dpi: return

        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG 圖片", "*.png")])
        if not path: return
        
        try:
            fig, ax = plt.subplots(figsize=(12, 8)) 
            ax.axis('tight')
            ax.axis('off')
            
            table = ax.table(cellText=self.dataframe.values,
                             colLabels=self.dataframe.columns,
                             rowLabels=self.dataframe.index,
                             cellLoc='center', 
                             loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.2)
            
            fig.savefig(path, dpi=dpi, bbox_inches='tight', pad_inches=0.5)
            plt.close(fig) 
            
            messagebox.showinfo("導出成功", f"表格已成功導出為圖片：\n{path}")
        except Exception as e:
            messagebox.showerror("導出錯誤", f"無法導出圖片：\n{e}")

    def _change_background_color(self):
        color_data = colorchooser.askcolor(title="選擇背景顏色", initialcolor=self.bg_color)
        if not color_data: return
        new_color = color_data[1]; self.bg_color = new_color; self.configure(bg=self.bg_color)
        self.style.configure('.', background=self.bg_color); self.style.configure('TFrame', background=self.bg_color); self.style.configure('TLabel', background=self.bg_color)

    def _change_resolution(self, resolution): self.geometry(resolution.split(" ")[0])

    def _clear_treeview(self): self.tree.delete(*self.tree.get_children()); self.tree["columns"] = ()
    
    def _on_row_select(self, event):
        """僅在行選擇變化時更新活動儲存格。"""
        selected_items = self.tree.selection()
        if not selected_items: return
        item_id = selected_items[0]
        if not self.active_column_id:
            try:
                self.active_column_id = "#1" 
            except IndexError:
                return 
        self._update_active_cell_display(item_id, self.active_column_id)

    def _on_cell_click(self, event):
        """處理儲存格的單次點擊事件。"""
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or not column_id: return
        
        if self.tree.selection() != (item_id,):
            self.tree.selection_set(item_id)

        self._update_active_cell_display(item_id, column_id)
        self.input_entry.focus_set()
        self.input_entry.select_range(0, tk.END)

    def _on_key_navigate(self, event):
        """處理方向鍵導航。"""
        selected_items = self.tree.selection()
        if not selected_items: return
        
        current_item = selected_items[0]
        
        if not self.active_column_id:
            self.active_column_id = "#1"

        if event.keysym == 'Up':
            next_item = self.tree.prev(current_item)
            if next_item:
                self.tree.selection_set(next_item)
                self.tree.see(next_item)
        elif event.keysym == 'Down':
            next_item = self.tree.next(current_item)
            if next_item:
                self.tree.selection_set(next_item)
                self.tree.see(next_item)
        elif event.keysym in ['Left', 'Right']:
            cols = list(self.tree['columns'])
            try:
                current_col_index = cols.index(self.tree.heading(self.active_column_id, 'text'))
            except (ValueError, tk.TclError):
                current_col_index = 0

            if event.keysym == 'Left':
                next_col_index = max(0, current_col_index - 1)
            else: # Right
                next_col_index = min(len(cols) - 1, current_col_index + 1)
            
            new_col_id = f"#{next_col_index + 1}"
            self._update_active_cell_display(current_item, new_col_id)
        
        return "break" 

    def _update_active_cell_display(self, item_id, column_id):
        """更新頂部輸入框和標籤以反映當前的活動儲存格。"""
        self.active_column_id = column_id
        try:
            col_index = int(column_id.replace('#', '')) - 1
            if col_index < 0: return
            column_name = self.tree['columns'][col_index]
            
            df_index = int(item_id)

            value = self.dataframe.loc[df_index, column_name]
            self.input_var.set(str(value))
            self.cell_pos_label.config(text=f"儲存格: {column_name}[{df_index}]")
        except (KeyError, IndexError, ValueError, TypeError):
             self.input_var.set("")
             self.cell_pos_label.config(text="儲存格:")

    def show_about(self): messagebox.showinfo("關於 PySheet", "Python 表格編輯器 (PySheet) v8.1\n\n一個使用 Tkinter 和 Pandas 製作的全功能表格應用程式。\n分析功能現包含皮爾森和斯皮爾曼相關性分析。\n開發者：Gemini")

class ExtendDialog(tk.Toplevel):
    """一個用於擴展表格的對話框。"""
    def __init__(self, parent):
        super().__init__(parent)
        self.transient(parent)
        self.parent = parent
        self.title("擴展表格")
        self.geometry("300x150")
        self._create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="數量:").grid(row=0, column=0, sticky="w", pady=5)
        self.count_var = tk.StringVar(value="10")
        ttk.Entry(main_frame, textvariable=self.count_var, width=10).grid(row=0, column=1, sticky="w")

        ttk.Label(main_frame, text="方向:").grid(row=1, column=0, sticky="w", pady=5)
        self.direction_var = tk.StringVar(value="down")
        dir_frame = ttk.Frame(main_frame)
        ttk.Radiobutton(dir_frame, text="向下 (增加行)", variable=self.direction_var, value="down").pack(side=tk.LEFT)
        ttk.Radiobutton(dir_frame, text="向右 (增加列)", variable=self.direction_var, value="right").pack(side=tk.LEFT, padx=10)
        dir_frame.grid(row=1, column=1, sticky="w")

        ttk.Button(main_frame, text="擴展", command=self.apply_extend).grid(row=2, column=1, sticky="e", pady=20)

    def apply_extend(self):
        try:
            count = int(self.count_var.get())
            if count <= 0:
                messagebox.showerror("輸入錯誤", "數量必須是正整數。", parent=self)
                return
        except ValueError:
            messagebox.showerror("輸入錯誤", "數量必須是有效的整數。", parent=self)
            return
        
        direction = self.direction_var.get()
        self.parent.execute_extend(count, direction)
        self.destroy()

if __name__ == "__main__":
    # 為了讓所有功能能運作，需要安裝以下函式庫
    # pip install pandas openpyxl fuzzywuzzy python-Levenshtein matplotlib seaborn scikit-learn scipy
    app = PySheetApp()
    app.mainloop()
