import os

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

AGG_MAP = {"Sum": "sum", "Average": "mean", "Count": "count"}
MAX_CATEGORIES = 20


class DataApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Data Forge v2.4")

        self.center_window(self.root, 1400, 850)

        self.df = None
        self.original_df = None
        self.file_path = None
        self.history = [] #for back button

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=0)
        self.root.grid_columnconfigure(0, weight=0)
        self.root.grid_columnconfigure(1, weight=1)

        self.sidebar = tk.Frame(self.root, width=200, bg="#34495e")
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self.sidebar.grid_propagate(False)  #fixed

        #status bar at the bottom.
        self.status_frame = tk.Frame(self.root, bg="#e6e6e6", height=30)
        self.status_frame.grid(row=1, column=1, sticky="ew")
        self.status_frame.grid_propagate(False)

        self.status_message = tk.Label(self.status_frame, text="Ready", bg="#e6e6e6", anchor="e")
        self.status_message.pack(side="right", padx=10)

        btn_style = {"bg": "#2c3e50", "fg": "white", "activebackground": "#1abc9c", "relief": "flat"}

        #buttons
        tk.Button(self.sidebar, text="Load File", command=self.load_file, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Select Columns", command=self.select_columns, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Filter Builder", command=self.open_filter_builder, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Pivot", command=self.open_pivot, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Quick Analysis", command=self.open_quick_analysis, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Save Excel", command=self.save, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Remove Duplicates", command=self.remove_duplicates, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Back", command=self.go_back, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Reset", command=self.reset, **btn_style).pack(fill="x", pady=2)

        self.tree = ttk.Treeview(self.root)
        self.tree.grid(row=0, column=1, sticky="nsew")

        tk.Label(
            self.sidebar, text="Developed by BaHaa", font=("Segoe UI", 9), fg="grey", bg="#34495e"
        ).pack(side="bottom", pady=10)

    #open and loads csv/txt/excel (3types of big files) to self.df
    def load_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[
                ("Supported Files", "*.txt *.csv *.xlsx *.xls"),
                ("Text Files", "*.txt"),
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls"),
            ]
        )
        if not self.file_path:
            return #stop here if not use a file

        ext = os.path.splitext(self.file_path)[1].lower()
        if ext == ".txt":
            self.df = pd.read_csv(self.file_path, sep="\t")
        elif ext == ".csv":
            self.df = pd.read_csv(self.file_path)
        elif ext in (".xlsx", ".xls"):
            self.df = pd.read_excel(self.file_path)
        else:
            messagebox.showerror("Error", "Unsupported file type")
            return

        self.original_df = self.df.copy() #original copy of the file.
        self.show_table(self.df)
        self.update_status_bar()
        self.set_status("File loaded")

    # renders df into the treeview, only first 200 rows for speed
    def show_table(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, stretch=False)

        for _, row in df.head(200).iterrows():  #view only 1st 200 row
            self.tree.insert("", "end", values=list(row))

    #open checkboxes to choose needed columns
    def select_columns(self):
        if self.df is None:
            return

        cols = self.df.columns.tolist()
        win = tk.Toplevel(self.root)
        win.withdraw()
        win.title("Select Columns")

        height = min(len(cols) * 25 + 250, 600)
        self.center_window(win, 350, height)
        win.deiconify()

        vars_dict = {}
        canvas = tk.Canvas(win)
        scrollbar = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="top", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for col in cols:
            var = tk.BooleanVar(value=False)
            tk.Checkbutton(frame, text=col, variable=var).pack(anchor="w")
            vars_dict[col] = var

        def apply_selection():
            selected = [c for c, v in vars_dict.items() if v.get()] #columns selected
            if not selected:
                messagebox.showwarning("Warning", "No columns selected")
                return

            self.history.append(self.df.copy())
            self.df = self.df[selected]
            self.show_table(self.df)
            self.update_status_bar()
            self.set_status(f"{len(selected)} columns kept")
            win.destroy()

        bottom_frame = tk.Frame(win)
        bottom_frame.pack(side="bottom", fill="x")
        tk.Button(bottom_frame, text="Apply", command=apply_selection).pack(anchor="w", padx=10, pady=5)
        self.update_status_bar()

    # multiple filters (exact/contains/startswith/endswith)
    def open_filter_builder(self):
        if self.df is None:
            return

        win = tk.Toplevel(self.root)
        win.title("Filter Builder")
        self.center_window(win, 400, 280)

        filters = []
        listbox = tk.Listbox(win, width=60)
        listbox.pack(pady=10)

        def add_filter(): #add new filter
            popup = tk.Toplevel(win)
            popup.title("Add Filter")
            self.center_window(popup, 350, 600)

            col_var = tk.StringVar(popup)
            col_var.set(self.df.columns[0])

            type_var = tk.StringVar(popup)
            type_var.set("contains")

            tk.Label(popup, text="Column").pack()
            tk.OptionMenu(popup, col_var, *self.df.columns).pack()

            tk.Label(popup, text="Type").pack()
            tk.OptionMenu(
                popup, type_var, "exact", "contains", "startswith", "endswith",
                command=lambda x: build_value_ui(),
            ).pack()

            value_frame = tk.Frame(popup)
            value_frame.pack(fill="both", expand=True)
            value_widget = {"type": None, "data": None}

            def build_value_ui(*args):
                for w in value_frame.winfo_children():
                    w.destroy()

                col = col_var.get()
                ftype = type_var.get()

                if ftype == "exact":
                    # listbox not checkbox avoid freezing on big columns
                    all_values = sorted(str(v) for v in self.df[col].dropna().unique())

                    search_entry = tk.Entry(value_frame)
                    search_entry.pack(fill="x", padx=10, pady=(10, 5))

                    list_frame = tk.Frame(value_frame)
                    list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

                    scrollbar = tk.Scrollbar(list_frame, orient="vertical")
                    value_listbox = tk.Listbox(list_frame, selectmode="extended", yscrollcommand=scrollbar.set)
                    scrollbar.config(command=value_listbox.yview)
                    value_listbox.pack(side="left", fill="both", expand=True)
                    scrollbar.pack(side="right", fill="y")

                    selected_values = set()

                    def populate_listbox(filter_text=""):
                        value_listbox.delete(0, "end")
                        filter_text = filter_text.lower()
                        for v in all_values:
                            if filter_text in v.lower():
                                value_listbox.insert("end", v)
                        for i, v in enumerate(value_listbox.get(0, "end")):
                            if v in selected_values:
                                value_listbox.selection_set(i)

                    def on_search(*args):
                        populate_listbox(search_entry.get())

                    def on_select(event):
                        visible = value_listbox.get(0, "end")
                        picked = {visible[i] for i in value_listbox.curselection()}
                        for v in visible:
                            selected_values.discard(v)
                        selected_values.update(picked)

                    search_entry.bind("<KeyRelease>", on_search)
                    value_listbox.bind("<<ListboxSelect>>", on_select)
                    populate_listbox()

                    value_widget["type"] = "exact"
                    value_widget["data"] = selected_values
                else:
                    tk.Label(value_frame, text="Enter Value").pack(anchor="w", padx=10, pady=(10, 0))
                    entry = tk.Entry(value_frame)
                    entry.pack(fill="x", padx=10, pady=10)
                    value_widget["type"] = "text"
                    value_widget["data"] = entry

            build_value_ui()

            def save_filter():
                col = col_var.get()
                ftype = type_var.get()

                if ftype == "exact":
                    selected = list(value_widget["data"])
                    if not selected:
                        messagebox.showwarning("Warning", "No values selected")
                        return
                    filters.append({"col": col, "type": "exact", "value": selected})
                else:
                    val = value_widget["data"].get()
                    if not val:
                        messagebox.showwarning("Warning", "Enter value")
                        return
                    filters.append({"col": col, "type": ftype, "value": val})

                listbox.insert("end", f"Column: {col} | Type: {ftype} | Value: {filters[-1]['value']}")
                popup.destroy()

            tk.Button(popup, text="Add Filter", command=save_filter).pack(pady=10)

        def run_filters():
            if not filters:
                messagebox.showwarning("Warning", "No filters added")
                return

            df = self.df.copy()
            for f in filters:
                col = f["col"]
                if f["type"] == "exact":
                    df = df[df[col].astype(str).isin(list(map(str, f["value"])))]
                elif f["type"] == "contains":
                    df = df[df[col].astype(str).str.contains(f["value"], na=False)]
                elif f["type"] == "startswith":
                    df = df[df[col].astype(str).str.startswith(f["value"], na=False)]
                elif f["type"] == "endswith":
                    df = df[df[col].astype(str).str.endswith(f["value"], na=False)]

            self.history.append(self.df.copy())
            self.df = df
            self.show_table(self.df)
            self.update_status_bar()
            self.set_status(f"{len(filters)} filter(s) applied")
            win.destroy()

        tk.Button(win, text="Add Filter", command=add_filter).pack(pady=5)
        tk.Button(win, text="Apply Filters", command=run_filters).pack(pady=5)

    # reshape data: fixed column, column with features, column values
    def open_pivot(self):
        if self.df is None:
            return

        win = tk.Toplevel(self.root)
        win.title("Pivot Tool")
        self.center_window(win, 300, 560)

        cols = self.df.columns.tolist()
        tk.Label(win, text="Index Columns (keep fixed)").pack()

        fixed_vars = {}
        fixed_frame = tk.Frame(win)
        fixed_frame.pack()
        for col in cols:
            var = tk.BooleanVar(value=False)
            tk.Checkbutton(fixed_frame, text=col, variable=var).pack(anchor="w")
            fixed_vars[col] = var

        tk.Label(win, text="Pivot Column (Features)").pack()
        pivot_col = tk.StringVar(win)
        pivot_col.set(cols[0])
        tk.OptionMenu(win, pivot_col, *cols).pack()

        tk.Label(win, text="Values Column").pack()
        value_col = tk.StringVar(win)
        value_col.set(cols[1] if len(cols) > 1 else cols[0])
        tk.OptionMenu(win, value_col, *cols).pack()

        tk.Label(win, text="Combine repeated values using").pack()
        agg_display = tk.StringVar(win)
        agg_display.set("Average")
        tk.OptionMenu(win, agg_display, *AGG_MAP.keys()).pack()

        def run_pivot():
            index_cols = [c for c, v in fixed_vars.items() if v.get()]
            if not index_cols:
                messagebox.showwarning("Warning", "Select at least one index column")
                return
            if pivot_col.get() == value_col.get():
                messagebox.showwarning("Warning", "Pivot and Value columns must differ")
                return

            try:
                agg_func = AGG_MAP[agg_display.get()]
                df_pivot = pd.pivot_table(
                    self.df, index=index_cols, columns=pivot_col.get(),
                    values=value_col.get(), aggfunc=agg_func,
                ).reset_index()

                df_pivot.columns = [
                    str(c) if isinstance(c, str) else str(c[1]) for c in df_pivot.columns
                ]

                self.history.append(self.df.copy())
                self.df = df_pivot
                self.show_table(self.df)
                self.update_status_bar()
                self.set_status(f"Pivot applied ({agg_display.get()})")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        tk.Button(win, text="Apply Pivot", command=run_pivot).pack(pady=10)

    #bar chart of the current data, X/Y
    def open_quick_analysis(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first.")
            return

        win = tk.Toplevel(self.root)
        win.title("Quick Analysis")
        self.center_window(win, 800, 650)

        columns = list(self.df.columns)
        controls = tk.Frame(win)
        controls.pack(fill="x", padx=10, pady=10)

        tk.Label(controls, text="X axis").grid(row=0, column=0, padx=5, sticky="w")
        x_var = tk.StringVar(value=columns[0])
        tk.OptionMenu(controls, x_var, *columns).grid(row=0, column=1, padx=5)

        tk.Label(controls, text="Y axis").grid(row=0, column=2, padx=5, sticky="w")
        y_var = tk.StringVar(value=columns[0])
        tk.OptionMenu(controls, y_var, *columns).grid(row=0, column=3, padx=5)

        tk.Label(controls, text="Aggregate").grid(row=0, column=4, padx=5, sticky="w")
        agg_display = tk.StringVar(value="Average")
        tk.OptionMenu(controls, agg_display, *AGG_MAP.keys()).grid(row=0, column=5, padx=5)

        chart_frame = tk.Frame(win)
        chart_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        def plot_chart():
            for w in chart_frame.winfo_children():
                w.destroy()

            x_col = x_var.get()
            y_col = y_var.get()
            agg_label = agg_display.get()
            agg_func = AGG_MAP[agg_label]

            try:
                fig = Figure(figsize=(7, 5), dpi=100)
                ax = fig.add_subplot(111)

                y_data = pd.to_numeric(self.df[y_col], errors="coerce")
                grouped = self.df.assign(_y=y_data).groupby(x_col)["_y"].agg(agg_func)
                grouped = grouped.sort_values(ascending=False)
                if len(grouped) > MAX_CATEGORIES:
                    grouped = grouped.iloc[:MAX_CATEGORIES]
                    ax.set_title(f"Top {MAX_CATEGORIES} of {x_col} by {agg_label.lower()} {y_col}")

                ax.bar(grouped.index.astype(str), grouped.values)
                ax.set_xlabel(x_col)
                ax.set_ylabel(f"{agg_label} of {y_col}")
                ax.tick_params(axis="x", rotation=45)
                fig.tight_layout()

                canvas = FigureCanvasTkAgg(fig, master=chart_frame)
                canvas.draw()
                canvas.get_tk_widget().pack(fill="both", expand=True)
            except Exception as e:
                messagebox.showerror("Chart error", str(e))

        tk.Button(controls, text="Plot", command=plot_chart).grid(row=0, column=6, padx=10)

    # back to last step from self.history.
    def go_back(self):
        if not self.history:
            messagebox.showinfo("Info", "No previous step available")
            return

        self.df = self.history.pop()
        self.show_table(self.df)
        self.update_status_bar()
        self.set_status("Back to previous step")

    #back to the original loaded file
    def reset(self):
        if self.original_df is not None:
            self.history.append(self.df.copy())
            self.df = self.original_df.copy()
            self.show_table(self.df)
            self.update_status_bar()
            self.set_status("Reset to original data")

    # export to excel, and open automaticly.
    def save(self):
        if self.df is None:
            return

        folder = os.path.dirname(self.file_path)
        output = os.path.join(folder, "output.xlsx")

        self.df.drop_duplicates().to_excel(output, index=False)
        os.startfile(output)
        self.set_status(f"Saved to {output}")

    #build the status bar with value counts per column
    def update_status_bar(self):
        for w in self.status_frame.winfo_children():
            if w != self.status_message:
                w.destroy()

        if self.df is None:
            return

        for col in self.df.columns:
            val = self.df[col].nunique()
            tk.Label(self.status_frame, text=f"{col}: {val}", bg="#e6e6e6", padx=10).pack(side="left")

    def set_status(self, text):
        self.status_message.config(text=text)

    def center_window(self, win, width, height): #centers the window on the screen
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = (sw // 2) - (width // 2)
        y = (sh // 2) - (height // 2)
        win.geometry(f"{width}x{height}+{x}+{y}")

    # remove duplicate rows and shows count of removed rows
    def remove_duplicates(self):
        if self.df is None:
            return

        before = len(self.df)
        self.history.append(self.df.copy())
        self.df = self.df.drop_duplicates()
        removed = before - len(self.df)

        self.show_table(self.df)
        self.update_status_bar()
        self.set_status(f"Removed {removed} duplicate rows")


if __name__ == "__main__":
    root = tk.Tk()
    app = DataApp(root)
    root.mainloop()
