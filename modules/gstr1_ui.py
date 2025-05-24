import tkinter as tk
from tkinter import filedialog, messagebox
import os
import datetime
import traceback  # For detailed error reporting
from collections import Counter  # Import Counter

# Assuming telemetry and license_util are in the same directory or accessible via PYTHONPATH
try:
    from telemetry import send_event
    # from license_util import get_machine_guid # Assuming this is available if needed
except ImportError:
    print("[WARN] Telemetry or license_util module not found. Telemetry will be disabled.")


    def send_event(event_name, payload):  # Dummy function if telemetry is not available
        pass

from gstr1_processor import process_gstr1, parse_filename, get_tax_period, parse_large_filename


class CustomErrorDialog(tk.Toplevel):
    def __init__(self, parent, title, message, error_details_to_copy):
        super().__init__(parent)
        self.title(title)
        self.error_details = error_details_to_copy
        self.parent = parent

        self.transient(parent)
        self.grab_set()

        main_frame = tk.Frame(self, padx=10, pady=10)
        main_frame.pack(expand=True, fill=tk.BOTH)

        icon_label = tk.Label(main_frame, text="❌", font=("Arial", 24), fg="red")
        icon_label.pack(side=tk.LEFT, padx=(0, 10))

        message_frame = tk.Frame(main_frame)
        message_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        tk.Label(message_frame, text=title, font=("Arial", 12, "bold")).pack(anchor="w")

        self.msg_text_widget = tk.Text(message_frame, wrap=tk.WORD, height=6, width=50, borderwidth=0,
                                       bg=self.cget('bg'))
        self.msg_text_widget.insert(tk.END, message)
        self.msg_text_widget.config(state=tk.DISABLED)
        self.msg_text_widget.pack(pady=5, fill=tk.X, expand=True)

        self.copy_status_label = tk.Label(message_frame, text="", fg="green")
        self.copy_status_label.pack(pady=(0, 5))

        button_frame = tk.Frame(self, pady=10)
        button_frame.pack()

        tk.Button(button_frame, text="OK", width=10, command=self.destroy).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Copy Error Details", width=15, command=self.copy_error).pack(side=tk.LEFT, padx=5)

        self.center_window()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    def copy_error(self):
        try:
            self.clipboard_clear()
            self.clipboard_append(self.error_details)
            self.copy_status_label.config(text="Error details copied to clipboard!")
            self.after(2000, lambda: self.copy_status_label.config(text=""))
        except tk.TclError:
            self.copy_status_label.config(text="Could not access clipboard.", fg="red")

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')


class GSTR1ProcessorUI:
    def __init__(self, root_window):  # Renamed root to root_window for clarity
        self.root = root_window
        self.root.title("GSTR1 Processing")
        self.root.geometry("500x480")
        self.small_files = []
        self.large_files = []
        self.template_file = None
        self.excluded_sections_by_month = {}
        self.base_height = 450

        tk.Label(self.root, text="GSTR1 Processing", font=("Arial", 16, "bold")).pack(pady=5)
        main_frame = tk.Frame(self.root)
        main_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        tk.Label(left_frame, text="GSTR1 JSON (<500)", font=("Arial", 10, "bold")).pack()
        self.small_listbox = tk.Listbox(left_frame, height=13, width=38, selectmode=tk.MULTIPLE)
        self.small_listbox.pack(pady=2, fill=tk.Y, expand=True)
        self.small_listbox.bind("<Button-1>", self.single_click_small)
        self.small_listbox.bind("<Shift-Button-1>", self.shift_click_small)
        self.small_listbox.bind("<Control-Button-1>", self.ctrl_click_small)
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="+ Add", command=self.add_small_file).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="- Remove", command=self.delete_small_file).pack(side=tk.LEFT, padx=5)

        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        tk.Label(right_frame, text="GSTR1 JSON Zip (>500)", font=("Arial", 10, "bold")).pack()
        self.large_listbox = tk.Listbox(right_frame, height=13, width=38, selectmode=tk.MULTIPLE)
        self.large_listbox.pack(pady=2, fill=tk.Y, expand=True)
        self.large_listbox.bind("<Button-1>", self.single_click_large)
        self.large_listbox.bind("<Shift-Button-1>", self.shift_click_large)
        self.large_listbox.bind("<Control-Button-1>", self.ctrl_click_large)
        large_btn_frame = tk.Frame(right_frame)
        large_btn_frame.pack(pady=5)
        tk.Button(large_btn_frame, text="+ Add", command=self.add_large_file).pack(side=tk.LEFT, padx=5)
        tk.Button(large_btn_frame, text="- Remove", command=self.delete_large_file).pack(side=tk.LEFT, padx=5)

        # Warning Frame (initially hidden, packed before template_frame when shown)
        self.warning_frame = tk.Frame(self.root, borderwidth=1, relief="solid")
        self.warning_title = tk.Label(self.warning_frame, text="Warning !", fg="red", font=("Arial", 10, "underline"))
        self.warning_text = tk.Label(self.warning_frame, text="", fg="red", justify=tk.LEFT, wraplength=450)
        self.ignore_var = tk.BooleanVar()
        # Changed text to match original
        self.ignore_check = tk.Checkbutton(self.warning_frame, text="Ignore All Warnings", variable=self.ignore_var,
                                           command=self.update_process_button)

        self.template_frame = tk.Frame(self.root)
        self.template_frame.pack(pady=5, fill=tk.X, padx=10)
        tk.Label(self.template_frame, text="Template Excel (Optional):").pack(side=tk.LEFT)
        self.template_label = tk.Label(self.template_frame, text="No file selected", width=25, anchor="w")
        self.template_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(self.template_frame, text="Select", command=self.select_template).pack(side=tk.LEFT, padx=2)
        tk.Button(self.template_frame, text="Clear", command=self.clear_template).pack(side=tk.LEFT, padx=2)

        self.process_btn = tk.Button(self.root, text="Process GSTR1", font=("Arial", 12), command=self.process_files,
                                     state=tk.DISABLED, bg="light grey")
        self.process_btn.pack(pady=10)

        self.update_process_button()

    def single_click_small(self, event):
        index = self.small_listbox.nearest(event.y)
        self.small_listbox.selection_clear(0, tk.END)
        self.small_listbox.selection_set(index)
        return "break"

    def shift_click_small(self, event):
        index = self.small_listbox.nearest(event.y)
        if not self.small_listbox.curselection():
            self.small_listbox.selection_set(index)
        else:
            anchor_tuple = self.small_listbox.curselection()
            if not anchor_tuple:
                self.small_listbox.selection_set(index)
                return "break"
            anchor = anchor_tuple[0]
            start, end = min(anchor, index), max(anchor, index)
            self.small_listbox.selection_clear(0, tk.END)
            for i in range(start, end + 1):
                self.small_listbox.selection_set(i)
        return "break"

    def ctrl_click_small(self, event):
        index = self.small_listbox.nearest(event.y)
        if self.small_listbox.selection_includes(index):
            self.small_listbox.selection_clear(index)
        else:
            self.small_listbox.selection_set(index)
        return "break"

    def single_click_large(self, event):
        index = self.large_listbox.nearest(event.y)
        self.large_listbox.selection_clear(0, tk.END)
        self.large_listbox.selection_set(index)
        return "break"

    def shift_click_large(self, event):
        index = self.large_listbox.nearest(event.y)
        if not self.large_listbox.curselection():
            self.large_listbox.selection_set(index)
        else:
            anchor_tuple = self.large_listbox.curselection()
            if not anchor_tuple:
                self.large_listbox.selection_set(index)
                return "break"
            anchor = anchor_tuple[0]
            start, end = min(anchor, index), max(anchor, index)
            self.large_listbox.selection_clear(0, tk.END)
            for i in range(start, end + 1):
                self.large_listbox.selection_set(i)
        return "break"

    def ctrl_click_large(self, event):
        index = self.large_listbox.nearest(event.y)
        if self.large_listbox.selection_includes(index):
            self.large_listbox.selection_clear(index)
        else:
            self.large_listbox.selection_set(index)
        return "break"

    def add_small_file(self):
        files = filedialog.askopenfilenames(filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")])
        financial_order = ["April", "May", "June", "July", "August", "September", "October", "November", "December",
                           "January", "February", "March"]
        new_files_added = False
        for file_path in files:
            month, excluded = parse_filename(file_path)
            if any(f[0] == file_path for f in self.small_files):
                print(f"Skipping already added file: {file_path}")
                continue
            if not month:
                messagebox.showwarning("Invalid Filename",
                                       f"Could not determine month for {os.path.basename(file_path)}. File not added.")
                continue

            self.small_files.append((file_path, month))
            new_files_added = True
            if excluded:
                if month not in self.excluded_sections_by_month:
                    self.excluded_sections_by_month[month] = []
                for ex_item in excluded:
                    if ex_item not in self.excluded_sections_by_month[month]:
                        self.excluded_sections_by_month[month].append(ex_item)

        if new_files_added:
            self.small_files.sort(key=lambda x: financial_order.index(get_tax_period(x[1])) if get_tax_period(
                x[1]) in financial_order else 999)
            self.small_listbox.delete(0, tk.END)
            for f_path, _ in self.small_files:
                self.small_listbox.insert(tk.END, os.path.basename(f_path))
        self.update_process_button()

    def delete_small_file(self):
        selections = self.small_listbox.curselection()
        if selections:
            for index in reversed(sorted(selections)):
                file_path, month = self.small_files.pop(index)
                self.small_listbox.delete(index)
                if month in self.excluded_sections_by_month and not any(m == month for _, m in self.small_files):
                    del self.excluded_sections_by_month[month]
            self.update_process_button()

    def add_large_file(self):
        files = filedialog.askopenfilenames(filetypes=[("ZIP Files", "*.zip"), ("All Files", "*.*")])
        financial_order = ["April", "May", "June", "July", "August", "September", "October", "November", "December",
                           "January", "February", "March"]
        new_files_added = False
        for file_path in files:
            month = parse_large_filename(file_path)
            if any(f[0] == file_path for f in self.large_files):
                print(f"Skipping already added file: {file_path}")
                continue
            if not month:
                messagebox.showwarning("Invalid Filename",
                                       f"Could not determine month for {os.path.basename(file_path)}. File not added.")
                continue

            self.large_files.append((file_path, month))
            new_files_added = True

        if new_files_added:
            self.large_files.sort(key=lambda x: financial_order.index(get_tax_period(x[1])) if get_tax_period(
                x[1]) in financial_order else 999)
            self.large_listbox.delete(0, tk.END)
            for f_path, _ in self.large_files:
                self.large_listbox.insert(tk.END, os.path.basename(f_path))
        self.update_process_button()

    def delete_large_file(self):
        selections = self.large_listbox.curselection()
        if selections:
            for index in reversed(sorted(selections)):
                self.large_files.pop(index)
                self.large_listbox.delete(index)
            self.update_process_button()

    def select_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if file_path:
            self.template_file = file_path
            self.template_label.config(text=os.path.basename(file_path))

    def clear_template(self):
        self.template_file = None
        self.template_label.config(text="No file selected")

    def update_process_button(self):
        warnings = []

        # --- Restored/Combined Warning Logic ---
        required_months_from_exclusions = set(self.excluded_sections_by_month.keys())
        selected_large_months = {month for _, month in self.large_files}
        selected_small_months = {month for _, month in self.small_files}

        # Original warning: If exclusions imply a month needs a >500 file, but it's missing
        missing_large_for_excluded = required_months_from_exclusions - selected_large_months
        if missing_large_for_excluded:
            warnings.append(
                f"'>500' JSON file for month(s) {', '.join(sorted(missing_large_for_excluded))} not selected.")

        # Duplicate <500 files for the same month
        small_month_counts = Counter(month for _, month in self.small_files)
        duplicate_small = [month for month, count in small_month_counts.items() if count > 1]
        if duplicate_small:
            warnings.append(f"Multiple '<500' JSON files selected for month(s): {', '.join(sorted(duplicate_small))}")

        # Duplicate >500 files for the same month
        large_month_counts = Counter(month for _, month in self.large_files)
        duplicate_large = [month for month, count in large_month_counts.items() if count > 1]
        if duplicate_large:
            warnings.append(f"Multiple '>500' JSON files selected for month(s): {', '.join(sorted(duplicate_large))}")

        # Original warning: If >500 file is present, but corresponding <500 is missing (or no <500 files at all)
        if self.large_files:  # Only check this if there are large files selected
            missing_small_for_large = selected_large_months - selected_small_months
            # The condition "or not self.small_files" from original code seems to mean:
            # if there are large files, AND (either specific small files are missing OR no small files at all are present)
            # This can be simplified to just checking missing_small_for_large if we assume <500 is always needed if >500 is present.
            if missing_small_for_large:  # Simplified this condition slightly from original
                months_str = ', '.join(sorted(missing_small_for_large))
                warnings.append(
                    f"No '<500' JSON file for month(s) {months_str} (to accompany >500 files). Details from <500 JSON will be missing.")
            elif not self.small_files and self.large_files:  # If large files are present, but absolutely no small files
                warnings.append(
                    f"No '<500' JSON files selected. Details from <500 JSON will be missing for all >500 files.")

        has_files = bool(self.small_files or self.large_files)

        current_warning_text = "\n".join(warnings)
        if warnings:
            if not self.warning_frame.winfo_ismapped():
                # Using original packing method, not `before=self.template_frame`
                self.warning_frame.pack(pady=5, padx=10, fill=tk.X)
            self.warning_title.pack(pady=(5, 0))
            self.warning_text.config(text=current_warning_text)
            self.warning_text.pack(pady=(0, 5))
            self.ignore_check.pack(pady=2)
            num_warning_lines = current_warning_text.count('\n') + 1
            # Original code used fixed +100, this dynamic one is generally better
            extra_height = 60 + (num_warning_lines * 15)
            self.root.geometry(f"500x{self.base_height + extra_height}")

            if self.ignore_var.get() and has_files:
                self.process_btn.config(state=tk.NORMAL, bg="light green")
            else:
                self.process_btn.config(state=tk.DISABLED, bg="light grey")
        else:
            if self.warning_frame.winfo_ismapped():
                self.warning_frame.pack_forget()
                self.warning_title.pack_forget()
                self.warning_text.pack_forget()
                self.ignore_check.pack_forget()
            self.root.geometry(f"500x{self.base_height}")
            if has_files:
                self.process_btn.config(state=tk.NORMAL, bg="light green")
            else:
                self.process_btn.config(state=tk.DISABLED, bg="light grey")

    def process_files(self):
        if not self.small_files and not self.large_files:
            messagebox.showerror("Error", "No files selected for processing.")
            return

        save_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Save GSTR1 Report As"
        )
        if not save_file:
            return

        self.process_btn.config(text="Processing...", state=tk.DISABLED, bg="light grey")
        self.root.update_idletasks()

        small_file_paths = [f for f, _ in self.small_files]
        large_files_map_for_processor = {}
        for file_path, month_key in self.large_files:
            large_files_map_for_processor[month_key] = (file_path, [])

        try:
            process_gstr1(
                small_file_paths,
                large_files_map_for_processor,
                self.excluded_sections_by_month,
                self.template_file,
                save_file,
                ignore_warnings=self.ignore_var.get()
            )

            send_event("gstr1_complete", {
                "input_small_files_count": len(small_file_paths),
                "input_large_files_count": len(large_files_map_for_processor),
                "template_used": bool(self.template_file),
                "output_file_extension": os.path.splitext(save_file)[1]
            })

            messagebox.showinfo("Success", f"GSTR1 report saved successfully at:\n{save_file}")
            self.small_files.clear()
            self.large_files.clear()
            self.excluded_sections_by_month.clear()
            self.small_listbox.delete(0, tk.END)
            self.large_listbox.delete(0, tk.END)
            self.clear_template()
            self.ignore_var.set(False)
            self.update_process_button()

        except Exception as e:
            detailed_error_info = traceback.format_exc()
            print("--- ERROR DETAILS ---")
            print(detailed_error_info)
            print("---------------------")

            send_event("error", {
                "module": "gstr1_ui.process_files",
                "error_type": type(e).__name__,
                "error_message": str(e),
                "input_small_files_count": len(small_file_paths),
                "input_large_files_count": len(large_files_map_for_processor),
            })

            CustomErrorDialog(self.root,
                              "Processing Error",
                              f"An error occurred during processing:\n\n{type(e).__name__}: {str(e)}\n\nSee console for full traceback if run from command line.",
                              detailed_error_info)
        finally:
            self.process_btn.config(text="Process GSTR1")
            self.update_process_button()


if __name__ == "__main__":
    root = tk.Tk()
    app = GSTR1ProcessorUI(root)
    root.mainloop()
