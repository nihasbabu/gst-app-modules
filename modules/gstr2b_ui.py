# gstr2b_ui.py

import tkinter as tk
from tkinter import filedialog, messagebox
import os
# import logging # Not used, can be removed
import traceback  # For detailed error reporting
import re  # For sorting logic in add_json_file

# Assuming telemetry is in the same directory or accessible via PYTHONPATH
try:
    from utils.telemetry import send_event
except ImportError:
    print("[WARN] Telemetry module not found in gstr2b_ui. Telemetry will be disabled.")


    def send_event(event_name, payload):  # Dummy function if telemetry is not available
        pass

try:
    # Ensure gstr2b_processor is in the python path or same directory
    from gstr2b_processor import process_gstr2b
except ImportError as e:
    print(f"[ERROR] Could not import 'process_gstr2b': {e}. Ensure 'gstr2b_processor.py' is accessible.")
    messagebox.showerror("Import Error",
                         f"Could not import 'process_gstr2b' from 'gstr2b_processor.py'.\nError: {e}\nEnsure the file exists and is accessible in PYTHONPATH.")
    process_gstr2b = None  # Define it as None so the UI can still load partially


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
                                       bg=self.cget('bg'))  # Use system background color
        self.msg_text_widget.insert(tk.END, message)
        self.msg_text_widget.config(state=tk.DISABLED)
        self.msg_text_widget.pack(pady=5, fill=tk.X, expand=True)

        self.copy_status_label = tk.Label(message_frame, text="", fg="green")
        self.copy_status_label.pack(pady=(0, 5))

        button_frame = tk.Frame(self, pady=10)
        button_frame.pack(fill=tk.X)

        def copy_error_to_clipboard_action():
            try:
                self.clipboard_clear()
                self.clipboard_append(self.error_details)
                self.copy_status_label.config(text="Error details copied to clipboard!")
                self.after(2000, lambda: self.copy_status_label.config(text=""))
            except tk.TclError:
                self.copy_status_label.config(text="Could not access clipboard.", fg="red")
                messagebox.showwarning("Clipboard Error", "Could not access the clipboard on this system.", parent=self)

        ok_button = tk.Button(button_frame, text="OK", width=10, command=self.destroy)
        copy_button = tk.Button(button_frame, text="Copy Error Details", width=15,
                                command=copy_error_to_clipboard_action)

        # Center buttons in the frame
        button_frame.columnconfigure(0, weight=1)  # Spacer
        button_frame.columnconfigure(1, weight=0)  # copy_button
        button_frame.columnconfigure(2, weight=0)  # ok_button
        button_frame.columnconfigure(3, weight=1)  # Spacer

        copy_button.grid(row=0, column=1, padx=5)
        ok_button.grid(row=0, column=2, padx=5)

        self.center_window()
        self.protocol("WM_DELETE_WINDOW", self.destroy)  # Handle window close button
        self.wait_window(self)  # Make it modal

    def center_window(self):
        self.update_idletasks()  # Ensure dimensions are calculated
        # Get parent window dimensions and position
        parent_x = self.parent.winfo_rootx()  # Use rootx/rooty for screen coordinates
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()

        # Get dialog's requested size
        dialog_width = self.winfo_reqwidth()
        dialog_height = self.winfo_reqheight()

        # Calculate position for centering
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2

        self.geometry(f'{dialog_width}x{dialog_height}+{x}+{y}')


class GSTR2BProcessorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("GSTR‑2B Processing")
        self.root.geometry("500x480")  # Initial size
        self.json_files = []  # List of GSTR‑2B JSON file paths
        self.template_file = None

        # Title
        tk.Label(root, text="GSTR‑2B Processing", font=("Arial", 16, "bold")).pack(pady=10)

        # Main Frame for Listbox
        main_frame = tk.Frame(root)
        main_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="GSTR‑2B JSON Files", font=("Arial", 10, "bold")).pack()
        self.json_list = tk.Listbox(main_frame, height=15, width=60, selectmode=tk.EXTENDED)
        self.json_list.pack(pady=0, fill=tk.BOTH, expand=True)  # Fill both X and Y

        # Add/Remove Buttons
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="+ Add", command=self.add_json_file, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="- Remove", command=self.delete_json_file, width=8).pack(side=tk.LEFT, padx=5)

        # Template File Section
        template_frame = tk.Frame(root)
        template_frame.pack(pady=5, fill=tk.X, padx=10)  # Fill X for better alignment
        tk.Label(template_frame, text="Template Excel (Optional):").pack(side=tk.LEFT)

        self.template_label_var = tk.StringVar(value="No file selected")
        self.template_label = tk.Label(template_frame, textvariable=self.template_label_var, anchor="w",
                                       width=25)  # Use anchor and width
        self.template_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        tk.Button(template_frame, text="Select", command=self.select_template, width=6).pack(side=tk.LEFT, padx=(0, 2))
        tk.Button(template_frame, text="Clear", command=self.clear_template, width=6).pack(side=tk.LEFT, padx=(0, 2))

        # Process Button
        self.process_btn = tk.Button(
            root,
            text="Process GSTR‑2B",
            font=("Arial", 12),
            command=self.process_files,
            state=tk.DISABLED,
            bg="light grey",
            width=20  # Give it a decent width
        )
        self.process_btn.pack(pady=10)

        self.update_process_button()

    def add_json_file(self):
        files = filedialog.askopenfilenames(
            title="Select GSTR-2B JSON Files",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        financial_months_order = ["04", "05", "06", "07", "08", "09", "10", "11", "12", "01", "02", "03"]
        new_files_added = False
        for file_path in files:
            if file_path not in self.json_files:
                self.json_files.append(file_path)
                new_files_added = True
            else:
                print(f"File already selected: {os.path.basename(file_path)}")

        if new_files_added:
            def sort_key_gstr2b(path):
                name = os.path.basename(path).lower()
                # Regex to find MMYYYY or MMMYYYY, e.g., 012023 or JAN2023 or gstr2b_042023_....json
                # Prioritize MMYYYY if found directly from known GSTR-2B naming patterns (e.g., rtnprd in content or filename convention)
                # This example sorts by filename which might contain the period.
                # A more robust sort would involve reading 'rtnprd' from each JSON.
                # For UI simplicity, filename sort is common.

                # Attempt to extract MMYYYY from typical GSTR-2B filenames like "gstr2b_MMYYYY_*.json" or just "MMYYYY.json"
                match_period = re.search(r'(?:gstr2b_)?(\d{2})(\d{4})', name)
                if not match_period:  # Try to find MMYYYY if it's just numbers
                    match_period = re.search(r'^(\d{2})(\d{4})\.json$', name)

                if match_period:
                    month_str, year_str = match_period.group(1), match_period.group(2)
                    if month_str in financial_months_order:
                        return (int(year_str), financial_months_order.index(month_str), name)
                # Fallback sort if no clear period in filename
                return (9999, 99, name)

            self.json_files.sort(key=sort_key_gstr2b)
            self.json_list.delete(0, tk.END)
            for file_item in self.json_files:
                self.json_list.insert(tk.END, os.path.basename(file_item))

        self.update_process_button()

    def delete_json_file(self):
        selected_indices = self.json_list.curselection()
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select file(s) to remove.", parent=self.root)
            return
        # Iterate in reverse to avoid index shifting issues
        for idx in reversed(selected_indices):
            self.json_files.pop(idx)
            self.json_list.delete(idx)
        self.update_process_button()

    def select_template(self):
        file = filedialog.askopenfilename(
            title="Select Template Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if file:
            self.template_file = file
            self.template_label_var.set(os.path.basename(file))

    def clear_template(self):
        self.template_file = None
        self.template_label_var.set("No file selected")

    def update_process_button(self):
        if self.json_files:
            self.process_btn.config(state=tk.NORMAL, bg="light green")
        else:
            self.process_btn.config(state=tk.DISABLED, bg="light grey")

    def process_files(self):
        if not self.json_files:
            messagebox.showerror("Error", "No JSON files selected for processing.", parent=self.root)
            return

        if process_gstr2b is None:  # Check if processor was imported
            messagebox.showerror("Error", "GSTR-2B processor is not available. Cannot proceed.", parent=self.root)
            return

        save_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Save GSTR‑2B Report As",
            initialfile="GSTR2B_Consolidated_Report.xlsx"  # Default filename
        )
        if not save_file:  # User cancelled save dialog
            return

        self.process_btn.config(text="Processing...", state=tk.DISABLED, bg="light grey")
        self.root.update_idletasks()  # Ensure UI updates before long operation

        base_message_from_processor = ""
        unexpected_details_list = []

        try:
            # MODIFIED: process_gstr2b now returns (message, unexpected_details_list)
            base_message_from_processor, unexpected_details_list = process_gstr2b(
self.json_files, self.template_file, save_file
            )

            final_display_message = base_message_from_processor

            # --- MODIFIED: Append warning to user message if unexpected sections are found ---
            if unexpected_details_list:
                warning_str_parts = [
                    "\n\nWARNING: Processing completed, but the following unexpected subsections were encountered in the JSON data. These might not have been extracted by the standard GSTR-2B logic. Please inform the app manager:"
                ]
                unique_warnings = set()  # To show each unique warning once
                for detail in unexpected_details_list:
                    # Check for file load errors specifically logged in unexpected_details_list
                    if detail.get("file_type") == "gstr2b_json_load_error":
                        warn_msg = f"- Error loading/parsing file '{detail['filename']}'."
                    else:
                        warn_msg = f"- Path '{detail['section_path']}' in file '{detail['filename']}' (Period: {detail.get('reporting_month', detail.get('raw_period', 'N/A'))})"
                    if warn_msg not in unique_warnings:
                        warning_str_parts.append(warn_msg)
                        unique_warnings.add(warn_msg)

                if len(unique_warnings) > 0:  # Only add if there are actual subsection warnings
                    final_display_message += "\n".join(warning_str_parts)

            # The detailed telemetry event `gstr2b_complete` is now sent from the processor.
            # The UI can send a simpler event if needed, or rely on the processor's event.
            # For now, let's assume processor handles the detailed completion telemetry.
            send_event("gstr2b_ui_process_attempt_complete", {  # A UI-specific event
                "input_files_count": len(self.json_files),
                "output_file": save_file,
                "template_used": bool(self.template_file),
                "status": "success",  # Assuming success if no exception before this point
                "had_unexpected_subsections": bool(unexpected_details_list)
            })

            messagebox.showinfo("Success", final_display_message, parent=self.root)

            # Reset UI elements after successful processing
            self.json_files.clear()
            self.json_list.delete(0, tk.END)
            self.clear_template()

        except PermissionError as pe:  # Catch permission errors specifically for saving
            detailed_error_info = traceback.format_exc()
            print(f"--- GSTR-2B UI PERMISSION ERROR ---")
            print(detailed_error_info)
            send_event("error", {
                "module": "gstr2b_ui.process_files", "error_type": "PermissionError",
                "error_message": str(pe), "filename": save_file
            })
            CustomErrorDialog(self.root, "File Save Error (GSTR-2B)",
                              f"Could not save the report file:\n{str(pe)}\n\nPlease ensure the file is not open and you have write permissions to the location.",
                              detailed_error_info)
        except Exception as e:
            detailed_error_info = traceback.format_exc()
            print(f"--- GSTR-2B UI UNEXPECTED ERROR ---")
            print(detailed_error_info)

            # Send detailed error telemetry from UI as a fallback or primary error event
            send_event("error", {
                "module": "gstr2b_ui.process_files",
                "error_type": type(e).__name__,
                "error_message": str(e),
                "input_files_count": len(self.json_files),
                "traceback_snippet": detailed_error_info[-1000:]  # Send last 1000 chars of traceback
            })

            CustomErrorDialog(self.root,
                              "Processing Error (GSTR-2B)",
                              f"An unexpected error occurred during GSTR-2B processing:\n\n{type(e).__name__}: {str(e)}\n\nMore details might be in the console if run from a command line.",
                              detailed_error_info)
        finally:
            self.process_btn.config(text="Process GSTR‑2B")  # Reset button text
            self.update_process_button()  # Re-evaluate button state


if __name__ == "__main__":
    root = tk.Tk()
    app = GSTR2BProcessorUI(root)
    root.mainloop()
