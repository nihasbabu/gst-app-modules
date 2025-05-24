import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from openpyxl import Workbook, load_workbook
import logging  # For consistency if processors use it
import traceback  # For detailed error reporting
from collections import Counter  # Import Counter

# Attempt to import processor functions
try:
    from credit_note_processor import process_credit_data
except ImportError:
    messagebox.showerror("Import Error",
                         "Could not import 'process_credit_data' from 'credit_note_processor.py'. Make sure the file exists and is in the Python path.")
    process_credit_data = None

try:
    from debit_note_processor import process_debit_note_data
except ImportError:
    messagebox.showerror("Import Error",
                         "Could not import 'process_debit_note_data' from 'debit_note_processor.py'. Make sure the file exists and is in the Python path.")
    process_debit_note_data = None

# Optional: For telemetry (wrapped in try-except)
try:
    from utils.telemetry import send_event
    # from utils.license_util import get_machine_guid # Assuming this is available if needed
except ImportError:
    print("[WARN] Telemetry or license_util module not found. Telemetry will be disabled.")


    def send_event(event_name, payload):  # Dummy function if telemetry is not available
        pass

PLACEHOLDER_TEXT = "Code"  # Or any other suitable placeholder like "Enter Code"


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
        icon_label.pack(side=tk.LEFT, padx=(0, 10), anchor='n')

        message_frame = tk.Frame(main_frame)
        message_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        tk.Label(message_frame, text=title, font=("Arial", 12, "bold")).pack(anchor="w")

        self.msg_text_widget = tk.Text(message_frame, wrap=tk.WORD, height=5, width=50, borderwidth=0,
                                       bg=self.cget('bg'))
        self.msg_text_widget.insert(tk.END, message)
        self.msg_text_widget.config(state=tk.DISABLED)
        self.msg_text_widget.pack(pady=5, fill=tk.X, expand=True)

        self.copy_status_label = tk.Label(message_frame, text="", fg="green")
        self.copy_status_label.pack(pady=(0, 5))

        button_frame = tk.Frame(self, pady=10)
        button_frame.pack(fill=tk.X)

        def copy_error_to_clipboard_action():
            try:
                self.clipboard_clear()  # Use self for Toplevel's clipboard
                self.clipboard_append(self.error_details)
                self.copy_status_label.config(text="Error details copied to clipboard!")
                self.after(2000, lambda: self.copy_status_label.config(text=""))
            except tk.TclError:
                self.copy_status_label.config(text="Could not access clipboard.", fg="red")
                messagebox.showwarning("Clipboard Error", "Could not access the clipboard on this system.", parent=self)

        ok_button = tk.Button(button_frame, text="OK", width=10, command=self.destroy)
        copy_button = tk.Button(button_frame, text="Copy Error Details", width=15,
                                command=copy_error_to_clipboard_action)

        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=0)
        button_frame.columnconfigure(2, weight=0)
        button_frame.columnconfigure(3, weight=1)

        copy_button.grid(row=0, column=1, padx=5)
        ok_button.grid(row=0, column=2, padx=5)

        self.center_window()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    def center_window(self):
        self.update_idletasks()
        # Get parent window geometry
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()

        dialog_width = self.winfo_reqwidth()
        dialog_height = self.winfo_reqheight()

        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        self.geometry(f'+{x}+{y}')


class CreditDebitNoteProcessorUI:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Process Credit / Debit Notes")

        self.credit_note_files = []
        self.debit_note_files = []
        self.template_file = None
        self.base_height = 500
        self.warning_frame_height_addition = 100

        self.root.geometry(f"500x{self.base_height}")

        tk.Label(self.root, text="Process Credit / Debit Notes", font=("Arial", 16, "bold")).pack(pady=5)
        main_frame = tk.Frame(self.root)
        main_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        # Credit Note Section (Left)
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        tk.Label(left_frame, text="Credit Note Files", font=("Arial", 10, "bold")).pack()

        self.credit_note_tree = ttk.Treeview(left_frame, columns=("File Name", "Branch Code"),
                                             show="headings", height=13)
        self.credit_note_tree.heading("File Name", text="File Name")
        self.credit_note_tree.heading("Branch Code", text="Branch Code")
        self.credit_note_tree.column("File Name", width=140, stretch=True)
        self.credit_note_tree.column("Branch Code", width=85, stretch=True)
        self.credit_note_tree.pack(pady=2, fill=tk.BOTH, expand=True)
        self.credit_note_tree.bind("<Double-1>", self.edit_credit_note_branch_code)

        btn_frame_credit = tk.Frame(left_frame)
        btn_frame_credit.pack(pady=5)
        tk.Button(btn_frame_credit, text="+ Add", command=self.add_credit_note_file).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame_credit, text="- Remove", command=self.delete_credit_note_file).pack(side=tk.LEFT, padx=5)

        # Debit Note Section (Right)
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        tk.Label(right_frame, text="Debit Note Files", font=("Arial", 10, "bold")).pack()

        self.debit_note_tree = ttk.Treeview(right_frame, columns=("File Name", "Branch Code"),
                                            show="headings", height=13)
        self.debit_note_tree.heading("File Name", text="File Name")
        self.debit_note_tree.heading("Branch Code", text="Branch Code")
        self.debit_note_tree.column("File Name", width=140, stretch=True)
        self.debit_note_tree.column("Branch Code", width=85, stretch=True)
        self.debit_note_tree.pack(pady=2, fill=tk.BOTH, expand=True)
        self.debit_note_tree.bind("<Double-1>", self.edit_debit_note_branch_code)

        btn_frame_debit = tk.Frame(right_frame)
        btn_frame_debit.pack(pady=5)
        tk.Button(btn_frame_debit, text="+ Add", command=self.add_debit_note_file).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame_debit, text="- Remove", command=self.delete_debit_note_file).pack(side=tk.LEFT, padx=5)

        # Template File Section
        self.template_frame = tk.Frame(self.root)
        self.template_frame.pack(pady=5, fill=tk.X, padx=10)
        tk.Label(self.template_frame, text="Template Excel File (Optional):").pack(side=tk.LEFT)
        self.template_label = tk.Label(self.template_frame, text="No file selected", width=25, anchor="w")
        self.template_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(self.template_frame, text="Select", command=self.select_template).pack(side=tk.LEFT, padx=2)
        tk.Button(self.template_frame, text="Clear", command=self.clear_template).pack(side=tk.LEFT, padx=2)

        # Process Button
        self.process_btn = tk.Button(self.root, text="Process Credit / Debit Notes", font=("Arial", 12),
                                     command=self.process_files, state=tk.DISABLED, bg="light grey")
        self.process_btn.pack(pady=10)

        # Warning Frame - will be packed by update_process_button_state if needed
        self.warning_frame = tk.Frame(self.root, borderwidth=1, relief="solid")
        self.warning_title = tk.Label(self.warning_frame, text="Warning!", fg="red",
                                      font=("Arial", 10, "underline"))
        self.warning_text = tk.Label(self.warning_frame, text="", fg="red",
                                     justify=tk.LEFT, wraplength=450)
        self.ignore_var = tk.BooleanVar()
        self.ignore_check = tk.Checkbutton(self.warning_frame, text="Ignore Warning and Proceed",
                                           variable=self.ignore_var,
                                           command=self.update_process_button_state)

        self.update_process_button_state()

    def _add_file_to_tree(self, file_list, tree_widget, file_type_name):
        files = filedialog.askopenfilenames(
            title=f"Select {file_type_name} Excel Files",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        for file_path in files:
            if not any(f[0] == file_path for f in file_list):
                base = os.path.basename(file_path)
                extracted_branch_key = base.split('_')[0].split('-')[0].split('.')[0] if base else ""
                if extracted_branch_key == os.path.splitext(base)[0]:
                    stored_branch_code = ""
                else:
                    stored_branch_code = extracted_branch_key.strip()

                file_list.append((file_path, stored_branch_code))
                display_branch_code = stored_branch_code if stored_branch_code else PLACEHOLDER_TEXT
                tree_widget.insert("", tk.END, values=(os.path.basename(file_path), display_branch_code))
        self.update_process_button_state()

    def _delete_file_from_tree(self, file_list, tree_widget):
        selected_items = tree_widget.selection()
        if not selected_items:
            return
        items_to_remove_from_list = []
        for item_id in selected_items:
            filename_in_tree = tree_widget.item(item_id, "values")[0]
            for i, (fp, stored_bc) in enumerate(file_list):
                if os.path.basename(fp) == filename_in_tree:
                    items_to_remove_from_list.append(file_list[i])
                    break
            tree_widget.delete(item_id)
        for item_tuple in items_to_remove_from_list:
            if item_tuple in file_list:
                file_list.remove(item_tuple)
        self.update_process_button_state()

    def _edit_branch_code(self, event, file_list, tree_widget):
        item_id = tree_widget.identify_row(event.y)
        column_id = tree_widget.identify_column(event.x)

        if not item_id or column_id != "#2":
            return

        current_filename_display = tree_widget.item(item_id, "values")[0]
        list_idx = -1
        try:
            for i, (fp, bc) in enumerate(file_list):
                if os.path.basename(fp) == current_filename_display:
                    list_idx = i
                    break
        except ValueError:
            logging.error(f"Item ID {item_id} not found in tree children. This should not happen.")
            return

        if list_idx == -1:
            logging.error(f"Could not find {current_filename_display} in internal file list for editing branch code.")
            return

        stored_branch_code = file_list[list_idx][1]
        entry = tk.Entry(tree_widget)
        entry.insert(0, stored_branch_code if stored_branch_code else "")
        if stored_branch_code:
            entry.select_range(0, tk.END)

        x, y, width, height = tree_widget.bbox(item_id, column_id)
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus_set()

        def save_branch_code(evt):
            new_branch_code_input = entry.get().strip()
            file_list[list_idx] = (file_list[list_idx][0], new_branch_code_input)
            display_text_for_tree = new_branch_code_input if new_branch_code_input else PLACEHOLDER_TEXT
            tree_widget.item(item_id, values=(current_filename_display, display_text_for_tree))
            entry.destroy()
            self.update_process_button_state()

        entry.bind("<Return>", save_branch_code)
        entry.bind("<FocusOut>", save_branch_code)
        entry.bind("<Escape>", lambda e: entry.destroy())

    def add_credit_note_file(self):
        self._add_file_to_tree(self.credit_note_files, self.credit_note_tree, "Credit Note")

    def delete_credit_note_file(self):
        self._delete_file_from_tree(self.credit_note_files, self.credit_note_tree)

    def edit_credit_note_branch_code(self, event):
        self._edit_branch_code(event, self.credit_note_files, self.credit_note_tree)

    def add_debit_note_file(self):
        self._add_file_to_tree(self.debit_note_files, self.debit_note_tree, "Debit Note")

    def delete_debit_note_file(self):
        self._delete_file_from_tree(self.debit_note_files, self.debit_note_tree)

    def edit_debit_note_branch_code(self, event):
        self._edit_branch_code(event, self.debit_note_files, self.debit_note_tree)

    def select_template(self):
        file = filedialog.askopenfilename(
            title="Select Template Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file:
            self.template_file = file
            self.template_label.config(text=os.path.basename(file))

    def clear_template(self):
        self.template_file = None
        self.template_label.config(text="No file selected")

    def update_process_button_state(self):
        has_any_files = bool(self.credit_note_files or self.debit_note_files)
        all_files_with_codes = self.credit_note_files + self.debit_note_files
        missing_branch_code = any(not code.strip() or code == PLACEHOLDER_TEXT for _, code in all_files_with_codes)

        if missing_branch_code and has_any_files:
            if not self.warning_frame.winfo_ismapped():
                # Pack warning frame *after* the process button to match original intent
                self.warning_frame.pack(pady=5, padx=10, fill=tk.X, after=self.process_btn)
            self.warning_title.pack(pady=(5, 0))
            self.warning_text.config(
                text="Warning: Branch Code is missing. Please double-click to edit or check 'Ignore Warning'.")
            self.warning_text.pack(pady=2)
            self.ignore_check.pack(pady=2)

            self.root.update_idletasks()
            current_warning_frame_height = self.warning_frame.winfo_reqheight()
            self.root.geometry(f"500x{self.base_height + current_warning_frame_height + 10}")

            if self.ignore_var.get():
                self.process_btn.config(state=tk.NORMAL, bg="light green")
            else:
                self.process_btn.config(state=tk.DISABLED, bg="light grey")
        else:
            if self.warning_frame.winfo_ismapped():
                self.warning_frame.pack_forget()
                self.warning_title.pack_forget()
                self.warning_text.pack_forget()
                self.ignore_check.pack_forget()
                self.root.update_idletasks()
            self.root.geometry(f"500x{self.base_height}")

            if has_any_files:
                self.process_btn.config(state=tk.NORMAL, bg="light green")
            else:
                self.process_btn.config(state=tk.DISABLED, bg="light grey")

    def process_files(self):
        if not (self.credit_note_files or self.debit_note_files):
            messagebox.showerror("Error", "No Credit Note or Debit Note files selected for processing.")
            return

        all_files_with_codes = self.credit_note_files + self.debit_note_files
        missing_branch_code_strict = any(
            not code.strip() or code == PLACEHOLDER_TEXT for _, code in all_files_with_codes)
        if missing_branch_code_strict and not self.ignore_var.get():
            messagebox.showwarning("Branch Code Missing",
                                   "Please review branch codes (cannot be empty or placeholder) or check 'Ignore Warning' to proceed.")
            self.update_process_button_state()
            return

        default_branch_for_processor = "Default_Branch"
        credit_notes_to_process = [
            (f, c.strip() if (c.strip() and c != PLACEHOLDER_TEXT) else default_branch_for_processor)
            for f, c in self.credit_note_files]
        debit_notes_to_process = [
            (f, c.strip() if (c.strip() and c != PLACEHOLDER_TEXT) else default_branch_for_processor)
            for f, c in self.debit_note_files]

        output_save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Combined Report As",
            initialfile="Processed_Credit_Debit_Notes.xlsx"
        )
        if not output_save_path:
            return

        self.process_btn.config(text="Processing...", state=tk.DISABLED, bg="light grey")
        self.root.update_idletasks()

        try:
            if self.template_file:
                logging.info(f"Loading template: {self.template_file}")
                wb = load_workbook(self.template_file)
            else:
                logging.info("Creating new workbook.")
                wb = Workbook()
                if 'Sheet' in wb.sheetnames and len(wb.sheetnames) == 1:
                    del wb['Sheet']

            if credit_notes_to_process and process_credit_data:
                logging.info(f"Processing {len(credit_notes_to_process)} credit note file(s)...")
                wb = process_credit_data(credit_notes_to_process, template_file=None, existing_wb=wb)
                logging.info("Credit note processing complete.")

            if debit_notes_to_process and process_debit_note_data:
                logging.info(f"Processing {len(debit_notes_to_process)} debit note file(s)...")
                wb = process_debit_note_data(debit_notes_to_process, template_file=None, existing_wb=wb)
                logging.info("Debit note processing complete.")

            wb.save(output_save_path)
            logging.info(f"Combined report saved successfully at: {output_save_path}")
            messagebox.showinfo("Success", f"Report saved successfully at:\n{output_save_path}")

            self.credit_note_files.clear()
            self.debit_note_files.clear()
            self.credit_note_tree.delete(*self.credit_note_tree.get_children())
            self.debit_note_tree.delete(*self.debit_note_tree.get_children())
            self.ignore_var.set(False)
            self.clear_template()

        except Exception as e:
            detailed_error_info = traceback.format_exc()
            logging.error(f"An error occurred during processing: {str(e)}", exc_info=True)
            # Call the modified show_error_with_copy
            self.show_error_with_copy(
                "Processing Error",
                f"An error occurred:\n{type(e).__name__}: {str(e)}",  # Short message
                detailed_error_info  # Full traceback for copying
            )
        finally:
            self.process_btn.config(text="Process Credit / Debit Notes")
            self.update_process_button_state()

    def show_error_with_copy(self, title, short_message, detailed_message_to_copy):
        """Displays a custom error dialog with a copy button for detailed error info."""
        logging.error(f"Displaying error to user: {title} - {short_message}")

        error_window = tk.Toplevel(self.root)
        error_window.title(title)
        error_window.resizable(False, False)

        main_frame = tk.Frame(error_window, padx=10, pady=10)
        main_frame.pack(expand=True, fill=tk.BOTH)

        icon_label = tk.Label(main_frame, text="❌", font=("Arial", 24), fg="red")
        icon_label.pack(side=tk.LEFT, padx=(0, 10), anchor='n')

        message_frame = tk.Frame(main_frame)
        message_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        tk.Label(message_frame, text=title, font=("Arial", 12, "bold")).pack(anchor="w")

        msg_text_widget = tk.Text(message_frame, wrap=tk.WORD, height=5, width=50, borderwidth=0,
                                  bg=error_window.cget('bg'))
        msg_text_widget.insert(tk.END, short_message)
        msg_text_widget.config(state=tk.DISABLED)
        msg_text_widget.pack(pady=5, fill=tk.X, expand=True)

        copy_status_label = tk.Label(message_frame, text="", fg="green")
        copy_status_label.pack(pady=(0, 5))

        button_frame = tk.Frame(error_window, pady=10)
        button_frame.pack(fill=tk.X)

        def copy_error_to_clipboard_action():
            try:
                error_window.clipboard_clear()
                error_window.clipboard_append(detailed_message_to_copy)
                copy_status_label.config(text="Error details copied to clipboard!")
                error_window.after(2000, lambda: copy_status_label.config(text=""))
            except tk.TclError:
                copy_status_label.config(text="Could not access clipboard.", fg="red")
                messagebox.showwarning("Clipboard Error", "Could not access the clipboard on this system.",
                                       parent=error_window)

        ok_button = tk.Button(button_frame, text="OK", width=10, command=error_window.destroy)
        copy_button = tk.Button(button_frame, text="Copy Error Details", width=15,
                                command=copy_error_to_clipboard_action)

        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=0)
        button_frame.columnconfigure(2, weight=0)
        button_frame.columnconfigure(3, weight=1)

        copy_button.grid(row=0, column=1, padx=5)
        ok_button.grid(row=0, column=2, padx=5)

        error_window.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        dialog_width = error_window.winfo_reqwidth()
        dialog_height = error_window.winfo_reqheight()

        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        error_window.geometry(f'+{x}+{y}')

        error_window.transient(self.root)
        error_window.grab_set()
        self.root.wait_window(error_window)


if __name__ == "__main__":
    if not logging.getLogger().hasHandlers():
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - UI - %(levelname)s - %(message)s')

    if process_credit_data is None or process_debit_note_data is None:
        logging.critical("One or both processor functions could not be imported. UI cannot function.")
    else:
        root = tk.Tk()
        app = CreditDebitNoteProcessorUI(root)
        root.mainloop()
