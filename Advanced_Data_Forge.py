"""
Advanced Data Forge
--------------------
A lightweight desktop GUI (built with Tkinter + pandas) for quickly
loading, cleaning, and reshaping tabular data (CSV / TXT / Excel)
without writing code for every small transformation.

Features:
    - Load CSV, TXT (tab-delimited), and Excel files
    - Select a subset of columns
    - Build multi-condition filters (exact / contains / starts with / ends with)
    - Pivot data around a chosen column
    - Remove duplicate rows
    - Undo (Back) and Reset to the originally loaded data
    - Export the current view to Excel

Author: BaHaa
"""

import os

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


class DataApp:
    """Main application window for Advanced Data Forge."""

    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Data Forge v2.4")

        # Default window size, centered on the user's screen.
        width = 1400
        height = 850

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        self.root.geometry(f"{width}x{height}+{x}+{y}")

        # ===== Application State =====
        self.df = None            # Currently displayed / active DataFrame
        self.original_df = None   # Untouched copy of the first loaded file
        self.file_path = None     # Path of the currently loaded file
        self.history = []         # Stack of previous DataFrame states (for "Back")

        # ===== Layout System =====
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=0)
        self.root.grid_columnconfigure(0, weight=0)
        self.root.grid_columnconfigure(1, weight=1)     

        self.sidebar = tk.Frame(self.root, width=200, bg="#34495e")
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self.sidebar.grid_propagate(False)

        self.status_frame = tk.Frame(self.root, bg="#e6e6e6", height=30)
        self.status_frame.grid(row=1, column=1, sticky="ew")
        self.status_frame.grid_propagate(False)

        self.status_message = tk.Label(self.status_frame, text="Ready", bg="#e6e6e6", anchor="e")
        self.status_message.pack(side="right", padx=10)

        btn_style = {"bg": "#2c3e50","fg": "white","activebackground": "#1abc9c","relief": "flat"}
    

        # ===== Buttons =====
        tk.Button(self.sidebar, text="Load File", command=self.load_file, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Select Columns", command=self.select_columns, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Filter Builder", command=self.open_filter_builder, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Pivot", command=self.open_pivot, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Quick Analysis", command=self.open_quick_analysis, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Save Excel", command=self.save, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Remove Duplicates", command=self.remove_duplicates, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Back", command=self.go_back, **btn_style).pack(fill="x", pady=2)
        tk.Button(self.sidebar, text="Reset", command=self.reset, **btn_style).pack(fill="x", pady=2)
        

        # ===== Table =====
        self.tree = ttk.Treeview(self.root)
        self.tree.grid(row=0, column=1, sticky="nsew")

                # ===== style =====
        self.tree.tag_configure("summary", background="#e6e6e6", font=("Segoe UI", 10, "bold"))

        # ===== Developer Credit =====
        credit_label = tk.Label(
            self.sidebar,
            text="Developed by BaHaa",
            font=("Segoe UI", 9),
            fg="grey",
            bg="#34495e"
        )

        credit_label.pack(side="bottom", pady=10)

    # ================= Load File =================
    def load_file(self):
        """Open a file dialog, load the selected file into a DataFrame,
        and refresh the table view."""

        self.file_path = filedialog.askopenfilename(
            filetypes=[
                ("Supported Files", "*.txt *.csv *.xlsx *.xls"),
                ("Text Files", "*.txt"),
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls")
            ]
        )

        if not self.file_path:
            return

        ext = os.path.splitext(self.file_path)[1].lower()

        if ext == ".txt":
            self.df = pd.read_csv(self.file_path, sep="\t")

        elif ext == ".csv":
            self.df = pd.read_csv(self.file_path)

        elif ext in [".xlsx", ".xls"]:
            self.df = pd.read_excel(self.file_path)

        else:
            messagebox.showerror("Error", "Unsupported file type")
            return        

        self.original_df = self.df.copy()

        self.show_table(self.df)
        self.update_status_bar()
        self.set_status("File loaded successfully")

    # ================= Show Table =================
    def show_table(self, df):
        """Render a DataFrame into the Treeview widget.

        Only the first 200 rows are displayed for performance; the
        underlying DataFrame (self.df) still holds all rows.
        """
        self.tree.delete(*self.tree.get_children())

        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        # ===== columns =====
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, stretch=False)

        # ===== data rows =====
        for _, row in df.head(200).iterrows():
            self.tree.insert("", "end", values=list(row))

    # ================= Select Columns =================
    def select_columns(self):
        """Open a checklist popup letting the user keep only a subset
        of columns from the current DataFrame."""
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
            selected = [c for c, v in vars_dict.items() if v.get()]

            if not selected:
                messagebox.showwarning("Warning", "No columns selected")
                return

            self.history.append(self.df.copy())
            self.df = self.df[selected]
            self.show_table(self.df)
            self.update_status_bar()
            self.set_status("Columns selected successfully")
            win.destroy()

        bottom_frame = tk.Frame(win)
        bottom_frame.pack(side="bottom", fill="x")
        tk.Button(bottom_frame, text="Apply", command=apply_selection).pack(anchor="w", padx=10, pady=5)
        self.update_status_bar()

    # ================= FILTER BUILDER =================
    def open_filter_builder(self):
        """Open the Filter Builder window, where the user can stack
        multiple column filters (exact / contains / startswith /
        endswith) and apply them all at once."""
        if self.df is None:
            return

        win = tk.Toplevel(self.root)
        win.title("Filter Builder")
        self.center_window(win, 400, 280)

        filters = []

        listbox = tk.Listbox(win, width=60)
        listbox.pack(pady=10)

        # ===== ADD FILTER WINDOW =====
        def add_filter():
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
            type_menu = tk.OptionMenu(
                popup,
                type_var,
                "exact", "contains", "startswith", "endswith",
                command=lambda x: build_value_ui()
            )
            type_menu.pack() 

            value_frame = tk.Frame(popup)
            value_frame.pack(fill="both", expand=True)

            value_widget = {"type": None, "data": None}

            def build_value_ui(*args):
                for w in value_frame.winfo_children():
                    w.destroy()

                col = col_var.get()
                ftype = type_var.get()

                # ===== EXACT: searchable multi-select list =====
                # Uses a single Listbox instead of one Checkbutton per
                # unique value. Creating thousands of individual
                # widgets (the old approach) is what caused the UI to
                # freeze on high-cardinality columns.
                if ftype == "exact":
                    all_values = sorted(str(v) for v in self.df[col].dropna().unique())

                    search_entry = tk.Entry(value_frame)
                    search_entry.pack(fill="x", padx=10, pady=(10, 5))

                    list_frame = tk.Frame(value_frame)
                    list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

                    scrollbar = tk.Scrollbar(list_frame, orient="vertical")
                    value_listbox = tk.Listbox(
                        list_frame,
                        selectmode="extended",
                        yscrollcommand=scrollbar.set,
                    )
                    scrollbar.config(command=value_listbox.yview)

                    value_listbox.pack(side="left", fill="both", expand=True)
                    scrollbar.pack(side="right", fill="y")

                    # Keep the selected values (not just indices) so
                    # a selection survives while the list is filtered.
                    selected_values = set()

                    def populate_listbox(filter_text=""):
                        value_listbox.delete(0, "end")
                        filter_text = filter_text.lower()
                        for v in all_values:
                            if filter_text in v.lower():
                                value_listbox.insert("end", v)
                        # Re-apply selection highlighting for values
                        # that are still visible after filtering.
                        for i, v in enumerate(value_listbox.get(0, "end")):
                            if v in selected_values:
                                value_listbox.selection_set(i)

                    def on_search(*args):
                        populate_listbox(search_entry.get())

                    def on_select(event):
                        # Update selected_values based on what's
                        # currently checked in the (possibly filtered) list.
                        visible = value_listbox.get(0, "end")
                        currently_selected = {
                            visible[i] for i in value_listbox.curselection()
                        }
                        # Remove visible items that got deselected, add newly selected ones.
                        for v in visible:
                            selected_values.discard(v)
                        selected_values.update(currently_selected)

                    search_entry.bind("<KeyRelease>", on_search)
                    value_listbox.bind("<<ListboxSelect>>", on_select)

                    populate_listbox()

                    value_widget["type"] = "exact"
                    value_widget["data"] = selected_values

                # ===== TEXT INPUT =====
                else:
                    label = tk.Label(value_frame, text="Enter Value")
                    label.pack(anchor="w", padx=10, pady=(10, 0))
                    entry = tk.Entry(value_frame)
                    entry.pack(fill="x", padx=10, pady=10)

                    value_widget["type"] = "text"
                    value_widget["data"] = entry
            
            build_value_ui()

            # ===== SAVE FILTER =====
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

        # ===== APPLY ALL FILTERS =====
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
            self.set_status("Filters applied successfully")
            win.destroy()

        tk.Button(win, text="Add Filter", command=add_filter).pack(pady=5)
        tk.Button(win, text="Apply Filters", command=run_filters).pack(pady=5)

    # ================= Pivot =================
    def open_pivot(self):
        """Open the Pivot Tool window, letting the user reshape the
        data by choosing index columns, a pivot column, a values
        column, and how to aggregate values when a combination of
        index + pivot column repeats (e.g. the same customer buying
        from the same category more than once)."""
        if self.df is None:
            return

        win = tk.Toplevel(self.root)
        win.title("Pivot Tool")
        self.center_window(win, 300, 560)

        cols = self.df.columns.tolist()

        tk.Label(win, text="Index Columns (keep fixed)").pack()
        
        # fixed columns (PDF_URL + SUPPLIER)
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
        if len(cols) > 1:
            value_col.set(cols[1])
        else:
            value_col.set(cols[0])
        tk.OptionMenu(win, value_col, *cols).pack()

        # When the same index+pivot combination repeats (e.g. one
        # customer buying from the same category twice), this decides
        # how those repeated values get combined into a single cell
        # instead of producing duplicate rows.
        tk.Label(win, text="Combine repeated values using").pack()
        agg_display = tk.StringVar(win)
        agg_display.set("Average")
        AGG_MAP = {"Sum": "sum", "Average": "mean", "Count": "count"}
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
                pivot_name = pivot_col.get()
                value_name = value_col.get()
                agg_func = AGG_MAP[agg_display.get()]

                # pivot_table aggregates repeated index+pivot
                # combinations automatically (sum/mean/count), so each
                # index value gets exactly one row in the result.
                df_pivot = pd.pivot_table(
                    self.df,
                    index=index_cols,
                    columns=pivot_name,
                    values=value_name,
                    aggfunc=agg_func,
                ).reset_index()

                df_pivot.columns = [
                    str(c) if isinstance(c, str) else str(c[1])
                    for c in df_pivot.columns
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

    # ================= QUICK ANALYSIS =================
    def open_quick_analysis(self):
        """Open a small charting window: pick an X column and a Y
        column, then plot a bar chart (average Y per X) against the
        currently displayed data (self.df) — so any active filters
        are reflected in the chart too.
        """
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
        x_menu = tk.OptionMenu(controls, x_var, *columns)
        x_menu.grid(row=0, column=1, padx=5)

        tk.Label(controls, text="Y axis").grid(row=0, column=2, padx=5, sticky="w")
        y_var = tk.StringVar(value=columns[0])
        y_menu = tk.OptionMenu(controls, y_var, *columns)
        y_menu.grid(row=0, column=3, padx=5)

        tk.Label(controls, text="Aggregate").grid(row=0, column=4, padx=5, sticky="w")
        agg_display = tk.StringVar(value="Average")
        AGG_MAP = {"Sum": "sum", "Average": "mean", "Count": "count"}
        agg_menu = tk.OptionMenu(controls, agg_display, *AGG_MAP.keys())
        agg_menu.grid(row=0, column=5, padx=5)

        # Container for the matplotlib canvas, recreated on every plot.
        chart_frame = tk.Frame(win)
        chart_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        MAX_CATEGORIES = 20  # Cap categorical X-axis groups to keep the chart readable

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

    # ================= BACK =================
    def go_back(self):
        """Undo the last operation by restoring the previous DataFrame
        state from the history stack."""
        if not self.history:
            messagebox.showinfo("Info", "No previous step available")
            return

        self.df = self.history.pop()

        self.show_table(self.df)
        self.update_status_bar()
        self.set_status("Returned to previous step")        

    # ================= RESET =================
    def reset(self):
        """Discard all changes and restore the originally loaded file."""
        if self.original_df is not None:
            self.history.append(self.df.copy())
            self.df = self.original_df.copy()
            self.show_table(self.df)
            self.update_status_bar()
            self.set_status("Reset applied successfully")

    # ================= SAVE =================
    def save(self):
        """Export the current DataFrame to 'output.xlsx' in the same
        folder as the loaded file, then open it automatically.

        Note: os.startfile() is Windows-only. On macOS/Linux the file
        will still be saved, but it won't open automatically.
        """
        if self.df is None:
            return

        folder = os.path.dirname(self.file_path)
        output = os.path.join(folder, "output.xlsx")

        df_to_save = self.df.drop_duplicates()
        df_to_save.to_excel(output, index=False)

        os.startfile(output)
        self.set_status("Export completed successfully")

    def get_unique_counts_row(self):
        """Return the count of unique values per column (used for the
        status bar summary)."""
        if self.df is None:
            return []

        return [self.df[col].nunique() for col in self.df.columns]

    def update_status_bar(self):
        """Refresh the bottom status bar with a unique-value count per
        column."""
        for w in self.status_frame.winfo_children():
            if w != self.status_message:
                w.destroy()

        if self.df is None:
            return

        for col in self.df.columns:
            val = self.df[col].nunique()

            lbl = tk.Label(self.status_frame,text=f"{col}: {val}",bg="#e6e6e6",padx=10)
            lbl.pack(side="left")        

    def set_status(self, text):
        """Update the status bar message on the right side of the window."""
        self.status_message.config(text=text)

    def center_window(self, win, width, height):
        """Center a Toplevel window on the screen given its target size."""
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()

        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        win.geometry(f"{width}x{height}+{x}+{y}")

    def remove_duplicates(self):
        """Drop duplicate rows from the current DataFrame and report
        how many were removed."""
        if self.df is None:
            return

        before = len(self.df)

        self.history.append(self.df.copy())

        self.df = self.df.drop_duplicates()

        after = len(self.df)
        removed = before - after

        self.show_table(self.df)
        self.update_status_bar()
        self.set_status(f"Removed {removed} duplicate rows")


# ================= RUN =================
if __name__ == "__main__":
    root = tk.Tk()
    app = DataApp(root)
    root.mainloop()