import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox

def update_status(text):
    status_label.config(text=text)
    status_label.update_idletasks()

def run_process(report_path, raw_path):
    if not report_path or not raw_path:
        messagebox.showwarning(
            "Missing File",
            "Please select both the report workbook and the raw workbook."
        )
        return

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False

    try:
        update_status("Opening report workbook...")
        report_wb = excel.Workbooks.Open(report_path, UpdateLinks=0)
        report_ws = report_wb.Worksheets("Raw Data")

        update_status("Clearing Raw Data sheet...")
        if report_ws.UsedRange.Cells.Count > 1:
            report_ws.UsedRange.ClearContents()

        update_status("Opening Raw workbook...")
        raw_wb = excel.Workbooks.Open(raw_path, UpdateLinks=0)
        raw_ws = raw_wb.Worksheets(1)  # Assume first sheet

        # Get data range
        used_range = raw_ws.UsedRange
        rows = used_range.Rows.Count
        cols = used_range.Columns.Count

        update_status("Copying data from Raw workbook...")
        report_ws.Range(
            report_ws.Cells(1,1),
            report_ws.Cells(rows, cols)
        ).Value = used_range.Value

        raw_wb.Close(False)

        update_status("Refreshing pivot tables...")
        # Pivot table(s) are on "Table" sheet
        pivot_ws = report_wb.Worksheets("Table")
        report_wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()

        update_status("Saving report workbook...")
        report_wb.Save()

        update_status("Process completed successfully.")
        messagebox.showinfo(
            "Success",
            "Raw data loaded and pivot tables refreshed."
        )

    except Exception as e:
        update_status("Error occurred.")
        messagebox.showerror("Error", str(e))

    finally:
        excel.ScreenUpdating = True
        excel.EnableEvents = True
        excel.DisplayAlerts = True

def browse(entry, filetypes):
    path = filedialog.askopenfilename(filetypes=filetypes)
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

# ---------------- UI ---------------- #

root = tk.Tk()
root.title("Amtrak v1.0.01")
root.geometry("520x260")
root.resizable(False, False)

tk.Label(root, text="Report Workbook").pack(pady=(10, 0))
report_entry = tk.Entry(root, width=70)
report_entry.pack()
tk.Button(
    root,
    text="Browse",
    command=lambda: browse(report_entry, [("Excel Files", "*.xlsx *.xlsm *.xls")])
).pack()

tk.Label(root, text="Raw Workbook").pack(pady=(10, 0))
raw_entry = tk.Entry(root, width=70)
raw_entry.pack()
tk.Button(
    root,
    text="Browse",
    command=lambda: browse(raw_entry, [("Excel Files", "*.xlsx *.xlsm *.xls")])
).pack()

tk.Button(
    root,
    text="Run Process",
    height=2,
    width=20,
    command=lambda: run_process(report_entry.get(), raw_entry.get())
).pack(pady=10)

status_label = tk.Label(root, text="Waiting for files...", anchor="w")
status_label.pack(fill="x", padx=10, pady=(5,0))

root.mainloop()
