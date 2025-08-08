import pandas as pd
import statsmodels.api as sm
from itertools import combinations
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Progressbar
import threading

# Function to calculate AICc from an OLS model
def calculate_aicc(model, X):
    n = X.shape[0]
    k = X.shape[1]
    if n <= k + 1:
        return np.inf
    aic = model.aic
    return aic + (2 * k * (k + 1)) / (n - k - 1)

# Function to calculate VIF values for a set of variables
def calculate_vif(X):
    vif_dict = {}
    for i in range(X.shape[1]):
        var_name = X.columns[i]
        vif = sm.OLS(X.iloc[:, i], sm.add_constant(X.drop(X.columns[i], axis=1))).fit().rsquared
        vif_value = 1 / (1 - vif) if vif < 1 else np.inf
        vif_dict[var_name] = vif_value
    return vif_dict

# Main analysis function to run OLS on all variable combinations
def run_analysis(df, y_column, x_columns, output_path, progress_callback=None, cancel_flag=None):
    y = df[y_column]
    X_full = df[x_columns]

    filtered_results = []
    total = sum(1 for k in range(2, len(X_full.columns) + 1) for _ in combinations(X_full.columns, k))
    current = 0

    for k in range(2, len(X_full.columns) + 1):
        for combo in combinations(X_full.columns, k):
            if cancel_flag and cancel_flag["stop"]:
                if progress_callback:
                    progress_callback(0)
                messagebox.showinfo("Cancelled", "Operation was cancelled by the user.")
                return

            current += 1
            if progress_callback:
                progress_callback(current / total * 100)

            X_combo = sm.add_constant(X_full[list(combo)])
            try:
                model = sm.OLS(y, X_combo).fit()
                pvals = model.pvalues.drop('const')
                vif_values = calculate_vif(X_combo.drop('const', axis=1))

                result = {
                    'num_vars': len(combo),
                    'adjusted_r2': model.rsquared_adj,
                    'aic': model.aic,
                    'aicc': calculate_aicc(model, X_combo),
                    'bic': model.bic,
                    'variables': combo,
                    'p_values': pvals.to_dict(),
                    'vif': vif_values
                }
                filtered_results.append(result)
            except Exception as e:
                print(f"❌ Error with combination: {combo}\n{e}")

    output_data = []
    for idx, res in enumerate(filtered_results, start=1):
        pvals = res['p_values']
        significant_vars = [var for var, p in pvals.items() if p < 0.05]
        num_significant = len(significant_vars)

        output_data.append({
            'Model_ID': idx,
            'Num_Variables': res['num_vars'],
            'Num_Significant': num_significant,
            'Variables': ', '.join(res['variables']),
            'Significant_Variables': ', '.join(significant_vars),
            'Adjusted_R2': res['adjusted_r2'],
            'AIC': res['aic'],
            'AICc': res['aicc'],
            'BIC': res['bic'],
            'VIFs': ', '.join([f'{k}: {v:.2f}' for k, v in res['vif'].items()]),
            'P_Values': ', '.join([f'{k}: {v:.6f}' for k, v in res['p_values'].items()])
        })

    df_results = pd.DataFrame(output_data)
    df_results.to_csv(output_path, index=False, encoding='utf-8-sig')

    if progress_callback:
        progress_callback(0)
    messagebox.showinfo("Done", f"Output saved to:   {output_path}")

# GUI setup function
def start_gui():
    def browse_file():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            input_entry.delete(0, tk.END)
            input_entry.insert(0, path)
            load_columns(path)

    def load_columns(path):
        nonlocal df, y_combo, check_vars
        df = pd.read_excel(path)
        columns = df.columns.tolist()
        y_combo['values'] = columns
        y_combo.current(0)
        for widget in check_vars_frame.winfo_children():
            widget.destroy()
        check_vars.clear()
        for col in columns:
            var = tk.BooleanVar(value=False)
            cb = tk.Checkbutton(check_vars_frame, text=col, variable=var)
            cb.pack(anchor='w')
            check_vars[col] = var

    def run():
        if df is None:
            messagebox.showerror("Error", "No data file loaded.")
            return
        y_column = y_combo.get()
        x_columns = [col for col, var in check_vars.items() if var.get() and col != y_column]
        if not x_columns:
            messagebox.showerror("Error", "No independent variables selected.")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not output_path:
            return

        cancel_flag["stop"] = False
        disable_widgets()
        progress_label.place(in_=progress_bar, relx=0.5, rely=0.5, anchor='center')
        progress_msg.place(relx=0.5, rely=0.5, anchor='center')

        if len(x_columns) > 10:
            progress_warning.place(relx=0.5, rely=0.57, anchor='center')

        def update_progress(val):
            progress.set(val)
            progress_label.config(text=f"{int(val)}%")

        def analysis_wrapper():
            run_analysis(df, y_column, x_columns, output_path, update_progress, cancel_flag)
            progress.set(0)
            progress_label.config(text="")
            progress_label.place_forget()
            progress_msg.place_forget()
            progress_warning.place_forget()
            enable_widgets()

        threading.Thread(target=analysis_wrapper).start()

    def cancel_processing():
        cancel_flag["stop"] = True
        progress.set(0)
        progress_label.config(text="")
        progress_label.place_forget()
        progress_msg.place_forget()
        progress_warning.place_forget()
        enable_widgets()

    def disable_widgets():
        browse_btn.config(state="disabled")
        input_entry.config(state="disabled")
        y_combo.config(state="disabled")
        run_button.config(state="disabled")
        for cb in check_vars_frame.winfo_children():
            cb.config(state="disabled")
        cancel_button.config(state="normal")

    def enable_widgets():
        browse_btn.config(state="normal")
        input_entry.config(state="normal")
        y_combo.config(state="readonly")
        run_button.config(state="normal")
        for cb in check_vars_frame.winfo_children():
            cb.config(state="normal")
        cancel_button.config(state="disabled")

    df = None
    check_vars = {}
    cancel_flag = {"stop": False}

    root = tk.Tk()
    root.title("ExhausOLS")
### root.iconbitmap("ExhausOLS_Logo.ico")

    # Center the window on screen
    window_width = 600
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = int((screen_width / 2) - (window_width / 2))
    y = int((screen_height / 2) - (window_height / 2))
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    tk.Label(root, text="Select input Excel file:").pack(pady=5)
    input_entry = tk.Entry(root, width=70)
    input_entry.pack()
    browse_btn = tk.Button(root, text="Browse", command=browse_file)
    browse_btn.pack(pady=5)

    tk.Label(root, text="Select  DEPENDENT  Variable:" ,  font=("Arial", 9, "bold") ).pack(pady=(20,5))
    y_combo = Combobox(root, state="readonly", width=50)
    y_combo.pack()

    tk.Label(root, text="Select INDEPENDENT Variables:", anchor='w', font=("Arial", 9, "bold")).pack(fill='x', padx=10, pady=(20, 5))

    check_vars_container = tk.Frame(root)
    check_vars_container.pack(pady=5, fill='both', expand=True)

    scrollbar = tk.Scrollbar(check_vars_container)
    scrollbar.pack(side='right', fill='y')

    canvas = tk.Canvas(check_vars_container, yscrollcommand=scrollbar.set)
    canvas.pack(side='left', fill='both', expand=True)
    scrollbar.config(command=canvas.yview)

    def on_mouse_wheel(event): canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    def bind_mousewheel_to_canvas(event): canvas.bind_all("<MouseWheel>", on_mouse_wheel)
    def unbind_mousewheel_from_canvas(event): canvas.unbind_all("<MouseWheel>")

    check_vars_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=check_vars_frame, anchor='nw')
    check_vars_frame.bind("<Enter>", bind_mousewheel_to_canvas)
    check_vars_frame.bind("<Leave>", unbind_mousewheel_from_canvas)
    check_vars_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    run_button = tk.Button(root, text="Run Analysis",  font=("Arial", 9, "bold"),command=run)
    run_button.pack(pady=10)

    cancel_button = tk.Button(root, text="Cancel", state="disabled", command=cancel_processing)
    cancel_button.pack(pady=2)

    progress = tk.DoubleVar()
    progress_bar = Progressbar(root, orient='horizontal', length=500, mode='determinate', variable=progress)
    progress_bar.pack(pady=(5, 0))

    progress_label = tk.Label(root, text="", font=("Arial", 9), bg="white")
    progress_label.place_forget()

    progress_msg = tk.Label(root, text="Please wait...", font=("Arial", 12, "bold"), fg="blue")
    progress_msg.place_forget()

    progress_warning = tk.Label(root, text="A large number of variables has been selected. The analysis may take longer than usual.",
                                font=("Arial", 10), fg="darkred", wraplength=275, justify='center')
    progress_warning.place_forget()

    footer = tk.Frame(root)
    footer.pack(side='bottom', fill='x')
    tk.Label(footer, text="© 2025 ExhausOLS - Zarei, Rafieian, Soltani", font=("Arial", 8), fg="gray").pack(fill='x', expand=True, pady=0)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
