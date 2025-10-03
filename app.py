# app.py
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import shutil
import pandas as pd
import datetime
import xlwings as xw

from generate import generate_output, get_preview, push_to_database

# helper to convert month names to numbers for sorting
def parse_month_to_num(month_name):
    try:
        return datetime.datetime.strptime(month_name, "%B").month
    except:
        return 0

class EventTemplateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Event Template Generator")

        # paths
        self.source_path = None
        self.staff_path = None
        self.roster_path = None
        self.output_path = None          # generated output path
        self.uploaded_output_path = None # corrected file path (if user uploads)

        # preview data (dict of dataframes)
        self.preview_data = None

        # cancel state
        self.cancel_push = False
        self.btn_cancel = None

        # filtered months sheet-wise
        self.filtered_months_dict = {}

        # build UI
        self._build_ui()

    def _build_ui(self):
        frm_top = tk.Frame(self.root)
        frm_top.pack(padx=10, pady=10, fill="x")

        tk.Button(frm_top, text="Select Source File", command=self.load_source).pack(side="left", padx=5)
        tk.Button(frm_top, text="Select Staff File", command=self.load_staff).pack(side="left", padx=5)
        tk.Button(frm_top, text="Select Roster File", command=self.load_roster).pack(side="left", padx=5)
        tk.Button(frm_top, text="Generate Output", command=self.generate_output).pack(side="left", padx=5)

        # status
        self.lbl_status = tk.Label(self.root, text="Please upload files and generate output.", fg="blue")
        self.lbl_status.pack(pady=5)

        # progress bar (determinate)
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(padx=10, pady=5, fill="x")
        self.progress['value'] = 0
        self.progress.pack_forget()

        # substatus label for step updates
        self.lbl_substatus = tk.Label(self.root, text="", fg="gray")
        self.lbl_substatus.pack(pady=2)

        # ---------- filtered months label ----------
        self.lbl_filtered_months = tk.Label(self.root, text="", fg="purple")
        self.lbl_filtered_months.pack(pady=2)

        # preview notebook
        self.preview_notebook = ttk.Notebook(self.root)
        self.preview_notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # DB buttons + download/upload area
        self.frm_db = tk.Frame(self.root)
        self.frm_db.pack(padx=10, pady=10, fill="x")

        self.btn_download = tk.Button(self.frm_db, text="Download Output File", command=self.download_output, state="disabled")
        self.btn_download.pack(side="left", padx=5)

        self.btn_upload = tk.Button(self.frm_db, text="Upload Corrected Output", command=self.upload_corrected_output, state="disabled")
        self.btn_upload.pack(side="left", padx=5)

        self.btn_push = tk.Button(self.frm_db, text="Push Records to Database", command=self.push_to_db, state="disabled")
        self.btn_push.pack(side="left", padx=5)

        # Cancel push button will be created but not packed; pack it when push starts
        self.btn_cancel = tk.Button(self.frm_db, text="Cancel Push", command=self.cancel_push_action)
        self.btn_cancel.pack_forget()

    # ---------- File Loaders ----------
    def load_source(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.source_path = path
        self.lbl_status.config(text=f"Source file loaded: {os.path.basename(path)}", fg="blue")
        self.lbl_filtered_months.config(text="Loading filtered months...")
        
        # ---------- Load filtered months sheet-wise using xlwings ----------
        def worker():
            try:
                app = xw.App(visible=False)
                wb = app.books.open(self.source_path)
                sheet_months = {}

                for sheet in wb.sheets:
                    last_row = sheet.used_range.last_cell.row
                    last_col = sheet.used_range.last_cell.column
                    header = sheet.range((1,1),(1,last_col)).value
                    if 'Month' not in header:
                        continue
                    month_idx = header.index('Month') + 1  # 1-based index
                    visible_months = []
                    for row in range(2, last_row+1):
                        if not sheet.range(f"A{row}").api.EntireRow.Hidden:
                            month_val = sheet.range((row, month_idx)).value
                            if month_val:
                                visible_months.append(str(month_val))
                    if visible_months:
                        sheet_months[sheet.name] = sorted(set(visible_months), key=parse_month_to_num)

                wb.close()
                app.quit()
                self.filtered_months_dict = sheet_months

                def update_label():
                    if sheet_months:
                        texts = []
                        for s, months in sheet_months.items():
                            texts.append(f"{s}: {', '.join(months)}")
                        self.lbl_filtered_months.config(
                            text="Filtered Months (source file only):\n" + "\n".join(texts)
                        )
                    else:
                        self.lbl_filtered_months.config(text="No filtered months found in source file.")
                self.safe_ui_update(update_label)
            except Exception as e:
                def update_err():
                    self.lbl_filtered_months.config(text=f"Error loading filtered months: {e}")
                self.safe_ui_update(update_err)

        threading.Thread(target=worker, daemon=True).start()

    def load_staff(self):
        self.staff_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.staff_path:
            self.lbl_status.config(text=f"Staff file loaded: {os.path.basename(self.staff_path)}", fg="blue")

    def load_roster(self):
        self.roster_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.roster_path:
            self.lbl_status.config(text=f"Roster file loaded: {os.path.basename(self.roster_path)}", fg="blue")

    # ---------- Generate Output ----------
    def generate_output(self):
        if not (self.source_path and self.staff_path and self.roster_path):
            messagebox.showerror("Error", "Please select all three files first.")
            return

        self.uploaded_output_path = None
        self.output_path = os.path.join(os.getcwd(), "output.xlsx")
        self.lbl_status.config(text="Generating output file...", fg="blue")
        self.lbl_substatus.config(text="")
        self.progress['value'] = 0
        self.progress.pack(padx=10, pady=5, fill="x")
        self.progress.start(10)

        def worker():
            try:
                def status_callback(current, total, message):
                    try:
                        pct = int((current / total) * 100)
                    except Exception:
                        pct = 0
                    def u():
                        try:
                            self.progress.stop()
                            self.progress['maximum'] = 100
                            self.progress['value'] = pct
                            self.lbl_substatus.config(text=message)
                        except:
                            pass
                    self.safe_ui_update(u)

                generate_output(self.source_path, self.staff_path, self.roster_path, self.output_path, status_callback=status_callback)
                preview = get_preview(self.output_path)

                def finish_ok():
                    self.lbl_status.config(text="✅ Output generated successfully!", fg="green")
                    self.preview_data = preview
                    self.show_preview()
                    self.btn_push.config(state="normal")
                    self.btn_download.config(state="normal")
                    self.btn_upload.config(state="normal")
                    self.progress['value'] = 100
                    self.progress.pack_forget()
                    self.lbl_substatus.config(text="")
                self.safe_ui_update(finish_ok)
            except Exception as e:
                def finish_err():
                    messagebox.showerror("Error", f"Failed to generate output: {e}")
                    self.lbl_status.config(text="❌ Failed to generate output.", fg="red")
                    self.progress.pack_forget()
                    self.lbl_substatus.config(text="")
                self.safe_ui_update(finish_err)

        threading.Thread(target=worker, daemon=True).start()

    # ---------- Preview ----------
    def show_preview(self):
        for tab_id in self.preview_notebook.tabs():
            self.preview_notebook.forget(tab_id)
        if not self.preview_data:
            return
        for sheet_name, df in self.preview_data.items():
            frame = tk.Frame(self.preview_notebook)
            self.preview_notebook.add(frame, text=sheet_name)
            cols = list(df.columns)
            tree = ttk.Treeview(frame, columns=cols, show="headings")
            vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            tree.configure(yscroll=vsb.set, xscroll=hsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            frame.grid_rowconfigure(0, weight=1)
            frame.grid_columnconfigure(0, weight=1)
            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=120, anchor="center")
            for _, row in df.iterrows():
                values = [row.get(c) for c in cols]
                tree.insert("", "end", values=values)

    # ---------- Download Output ----------
    def download_output(self):
        if not self.output_path or not os.path.exists(self.output_path):
            messagebox.showerror("Error", "No generated output file to download. Please generate output first.")
            return
        dest = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not dest:
            return
        try:
            shutil.copyfile(self.output_path, dest)
            messagebox.showinfo("Saved", f"Output saved to: {dest}")
            self.lbl_status.config(text=f"Output downloaded to {os.path.basename(dest)}", fg="green")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

    # ---------- Upload Corrected Output ----------
    def upload_corrected_output(self):
        path = filedialog.askopenfilename(title="Select corrected output file", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            pd.ExcelFile(path)
        except Exception as e:
            messagebox.showerror("Error", f"Selected file is not a valid Excel file: {e}")
            return
        self.uploaded_output_path = path
        self.preview_data = get_preview(path)
        self.show_preview()
        self.lbl_status.config(text=f"Uploaded corrected file: {os.path.basename(path)}", fg="blue")
        self.btn_push.config(state="normal")

    # ---------- Cancel Push ----------
    def cancel_push_action(self):
        self.cancel_push = True
        self.lbl_status.config(text="⏹ Cancelling push... please wait.", fg="orange")
        try:
            self.btn_cancel.config(state="disabled")
        except:
            pass

    # ---------- safe UI helper ----------
    def safe_ui_update(self, update_fn):
        try:
            if self.root and self.root.winfo_exists():
                self.root.after(0, update_fn)
        except:
            pass

    # ---------- Push to Database ----------
    def push_to_db(self):
        if self.uploaded_output_path:
            data_source = self.uploaded_output_path
        elif self.output_path and os.path.exists(self.output_path):
            data_source = self.output_path
        else:
            messagebox.showerror("Error", "No data available to push. Generate output or upload corrected sheet first.")
            return

        self.btn_push.config(state="disabled")
        self.lbl_status.config(text="Pushing records to database...", fg="blue")
        self.cancel_push = False
        try:
            self.btn_cancel.pack(side="left", padx=5)
            self.btn_cancel.config(state="normal")
        except:
            pass

        def worker():
            try:
                preview = get_preview(data_source)
                def cancel_check(): return self.cancel_push
                count, errors = push_to_database(preview, cancel_check=cancel_check)

                def finish_ui():
                    if self.cancel_push:
                        messagebox.showinfo("Cancelled", f"⏹ Push cancelled. {count} records inserted before stopping.")
                        self.lbl_status.config(text=f"⏹ Push cancelled — {count} inserted", fg="orange")
                    else:
                        messagebox.showinfo("Success", f"✅ {count} records pushed successfully!")
                        self.lbl_status.config(text=f"✅ {count} records pushed successfully!", fg="green")
                    self.btn_push.config(state="normal")
                    try: self.btn_cancel.pack_forget()
                    except: pass
                    if errors:
                        summary = "\n".join([f"{s} row {r}: {m}" for s, r, m in errors[:50]])
                        extra = f"\n...and {len(errors)-50} more." if len(errors) > 50 else ""
                        messagebox.showwarning("Some rows failed", f"{len(errors)} rows failed to push.\n\n{summary}{extra}")
                self.safe_ui_update(finish_ui)
            except Exception as e:
                def update_error():
                    messagebox.showerror("Error", f"Database push failed:\n{e}")
                    self.lbl_status.config(text="Database push failed!", fg="red")
                    self.btn_push.config(state="normal")
                    try: self.btn_cancel.pack_forget()
                    except: pass
                self.safe_ui_update(update_error)

        threading.Thread(target=worker, daemon=True).start()

# ---------- Run App ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = EventTemplateApp(root)
    root.geometry("1000x600")
    root.mainloop()
