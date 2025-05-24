# gstr3b_ui.py

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import traceback  # For detailed error reporting

# Assuming telemetry is in the same directory or accessible via PYTHONPATH
try:
    from telemetry import send_event
except ImportError:
    print("[WARN] Telemetry module not found. Telemetry will be disabled.")


    def send_event(event_name, payload):  # Dummy function if telemetry is not available
        pass

from gstr3b_processor import process_gstr3b


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
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()

        dialog_width = self.winfo_reqwidth()
        dialog_height = self.winfo_reqheight()

        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        self.geometry(f'+{x}+{y}')


class GSTR3BProcessorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("GSTR-3B Processor")
        self.root.geometry("500x480")
        self.json_files = []  # List of file paths
        self.template_file = None

        # Title
        tk.Label(root, text="GSTR-3B JSON Processing", font=("Arial", 16, "bold")).pack(pady=10)

        # Main Frame for List Box
        main_frame = tk.Frame(root)
        main_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        # GSTR3B JSON Section
        tk.Label(main_frame, text="GSTR-3B JSON Files", font=("Arial", 10, "bold")).pack()
        self.json_list = tk.Listbox(main_frame, height=15, width=60, selectmode=tk.EXTENDED)
        self.json_list.pack(pady=0, fill=tk.Y, expand=True)

        # Add/Remove Buttons
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="+ Add", command=self.add_json_file).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="- Remove", command=self.delete_json_file).pack(side=tk.LEFT, padx=5)

        # Template File Section
        template_frame = tk.Frame(root)
        template_frame.pack(pady=5)
        tk.Label(template_frame, text="Template Excel File (Optional):").pack(side=tk.LEFT)
        self.template_label = tk.Label(template_frame, text="No file selected")
        self.template_label.pack(side=tk.LEFT, padx=5)
        tk.Button(template_frame, text="Select", command=self.select_template).pack(side=tk.LEFT, padx=2)
        tk.Button(template_frame, text="Clear", command=self.clear_template).pack(side=tk.LEFT, padx=2)

        # Process Button
        self.process_btn = tk.Button(
            root,
            text="Process GSTR-3B",
            font=("Arial", 12),
            command=self.process_files,
            state=tk.DISABLED,
            bg="light grey"
        )
        self.process_btn.pack(pady=10)

        self.update_process_button()

    def add_json_file(self):
        files = filedialog.askopenfilenames(filetypes=[("JSON Files", "*.json")])
        financial_months = ["04", "05", "06", "07", "08", "09", "10", "11", "12", "01", "02", "03"]

        new_files_added = False
        for file_path in files:
            if file_path not in self.json_files:
                self.json_files.append(file_path)
                new_files_added = True

        if new_files_added:
            def sort_key(path):
                name = os.path.basename(path).lower()
                if name.endswith('.json'):
                    name = name[:-5]
                parts = name.split('_')

                month_year_part = None
                if len(parts) > 1 and len(parts[1]) == 6 and parts[1].isdigit():
                    month_year_part = parts[1]
                elif len(parts[0]) == 6 and parts[0].isdigit():
                    month_year_part = parts[0]

                if month_year_part:
                    m, y = month_year_part[:2], month_year_part[2:]
                    if m in financial_months and y.isdigit():
                        return (financial_months.index(m), int(y))
                return (len(financial_months) + 1, name)

            self.json_files.sort(key=sort_key)
            self.json_list.delete(0, tk.END)
            for file_item in self.json_files:
                self.json_list.insert(tk.END, os.path.basename(file_item))

        self.update_process_button()

    def delete_json_file(self):
        selected = self.json_list.curselection()
        if not selected:
            return
        for idx in reversed(selected):
            self.json_files.pop(idx)
            self.json_list.delete(idx)
        self.update_process_button()

    def select_template(self):
        file = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if file:
            self.template_file = file
            self.template_label.config(text=os.path.basename(file))

    def clear_template(self):
        self.template_file = None
        self.template_label.config(text="No file selected")

    def update_process_button(self):
        if self.json_files:
            self.process_btn.config(state=tk.NORMAL, bg="light green")
        else:
            self.process_btn.config(state=tk.DISABLED, bg="light grey")

    def process_files(self):
        if not self.json_files:
            messagebox.showerror("Error", "No JSON files selected for processing.")
            return

        save_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Save GSTR-3B Summary Report As",
            initialfile="GSTR3B_Consolidated_Report.xlsx"
        )
        if not save_file:
            return

        self.process_btn.config(text="Processing...", state=tk.DISABLED, bg="light grey")
        self.root.update_idletasks()

        try:
            result_message = process_gstr3b(self.json_files, self.template_file, save_file)

            send_event("gstr3b_complete", {
                "input_files_count": len(self.json_files),
                "output_file": save_file,
                "template_used": bool(self.template_file),
                "message": result_message
            })

            messagebox.showinfo("Success", result_message)
            self.json_files.clear()
            self.json_list.delete(0, tk.END)
            self.clear_template()

        except Exception as e:
            detailed_error_info = traceback.format_exc()
            print("--- GSTR-3B UI ERROR DETAILS ---")
            print(detailed_error_info)
            print("--------------------------------")

            send_event("error", {
                "module": "gstr3b_ui.process_files",
                "error_type": type(e).__name__,
                "error_message": str(e),
                "input_files_count": len(self.json_files)
            })

            CustomErrorDialog(self.root,
                              "Processing Error (GSTR-3B)",
                              f"An error occurred during GSTR-3B processing:\n\n{type(e).__name__}: {str(e)}\n\nSee console for full traceback if run from command line.",
                              detailed_error_info)
        finally:
            self.process_btn.config(text="Process GSTR-3B")
            self.update_process_button()


if __name__ == "__main__":
    root = tk.Tk()
    app = GSTR3BProcessorUI(root)
    root.mainloop()
