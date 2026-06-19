import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

file_path = ""


def select_file():
    global file_path

    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if file_path:
        path_label.config(text=file_path)


def process_excel():
    global file_path

    if not file_path:
        messagebox.showerror(
            "Error",
            "Please select an Excel file first."
        )
        return

    try:
        df = pd.read_excel(file_path)

        if "Sales" not in df.columns:
            messagebox.showerror(
                "Error",
                "Excel file must contain a 'Sales' column."
            )
            return

        total_sales = df["Sales"].sum()

        total_row = pd.DataFrame(
            [["TOTAL", total_sales]],
            columns=["Name", "Sales"]
        )

        result_df = pd.concat(
            [df, total_row],
            ignore_index=True
        )

        output_file = os.path.join(
            os.path.dirname(file_path),
            "processed_sales_report.xlsx"
        )

        result_df.to_excel(
            output_file,
            index=False
        )

        messagebox.showinfo(
            "Success",
            f"Report created:\n{output_file}"
        )

    except Exception as e:
        messagebox.showerror(
            "Error",
            str(e)
        )


root = tk.Tk()

root.title("Excel Automation Tool")
root.geometry("800x500")

title = tk.Label(
    root,
    text="Excel Automation Tool",
    font=("Arial", 20, "bold")
)

title.pack(pady=20)

btn_select = tk.Button(
    root,
    text="Select Excel File",
    font=("Arial", 12),
    command=select_file
)

btn_select.pack(pady=10)

path_label = tk.Label(
    root,
    text="No file selected",
    wraplength=700
)

path_label.pack()

btn_process = tk.Button(
    root,
    text="Generate Report",
    font=("Arial", 12),
    command=process_excel
)

btn_process.pack(pady=20)

result_label = tk.Label(
    root,
    text="Ready"
)

result_label.pack()

root.mainloop()