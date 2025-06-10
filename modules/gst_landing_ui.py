# src/gst_landing_ui.py
# (Based on the version with external license check, simplified error message, and hardcoded support info)

import sys
import os
import subprocess
import json
import traceback
import tkinter as tk
from tkinter import messagebox
import logging

# Define the current version of THIS application
CURRENT_APP_VERSION = "1.0.0"  # From gst_landing_ui_py_versioned_update, kept for consistency


# ─────────────────────────────────────────────────────────────────────────────
# Helper function to determine resource paths
# ─────────────────────────────────────────────────────────────────────────────
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# ─────────────────────────────────────────────────────────────────────────────
# Determine base and modules directory
# ─────────────────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    executable_location_dir = os.path.dirname(sys.executable)
    app_base_dir = sys._MEIPASS
else:
    executable_location_dir = os.path.dirname(os.path.abspath(__file__))
    app_base_dir = os.path.dirname(os.path.abspath(__file__))

if getattr(sys, "frozen", False):
    actual_app_root_for_modules = os.path.dirname(sys.executable)
else:
    actual_app_root_for_modules = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

modules_dir = os.path.join(actual_app_root_for_modules, "modules")
os.makedirs(modules_dir, exist_ok=True)

if modules_dir not in sys.path:
    sys.path.insert(0, modules_dir)

if not getattr(sys, "frozen", False):
    src_dir_path = os.path.dirname(os.path.abspath(__file__))
    if src_dir_path not in sys.path:
        sys.path.insert(1, src_dir_path)
else:
    if sys._MEIPASS not in sys.path:
        sys.path.insert(1, sys._MEIPASS)

# ─────────────────────────────────────────────────────────────────────────────
# Configure Logging
# ─────────────────────────────────────────────────────────────────────────────
log_file_path = os.path.join(actual_app_root_for_modules, 'gst_processor_app.log')
if not logging.getLogger().hasHandlers():
    file_handler = logging.FileHandler(log_file_path, mode='a')
    file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    console_handler = logging.StreamHandler(sys.stdout)
    console_formatter = logging.Formatter('%(levelname)s: %(name)s: %(message)s')
    console_handler.setFormatter(console_formatter)
    console_handler.setLevel(logging.INFO)
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Bootstrap modules folder & auto‑update logic
# ─────────────────────────────────────────────────────────────────────────────
try:
    from utils.updater import update_modules

    logger.info(f"Attempting to call update_modules for current app version: {CURRENT_APP_VERSION}")
    # Pass the CURRENT_APP_VERSION to the updater
    update_modules(modules_dir, CURRENT_APP_VERSION)  # Assuming updater takes version
    logger.info("update_modules call completed.")
except ImportError as e_updater_imp:
    logger.warning(
        f"Updater module (utils/updater.py) not found or import error: {e_updater_imp}. Cannot check for updates.")
except Exception as e_updater:
    logger.error(f"Updater failed during initialization: {e_updater}\n{traceback.format_exc()}")
    pass

# ─────────────────────────────────────────────────────────────────────────────
# Import utilities
# ─────────────────────────────────────────────────────────────────────────────
try:
    from utils.license_util import get_machine_guid
    from utils.telemetry import send_event  # This send_event will be used

    logger.info("Successfully imported license_util and telemetry from utils.")
except ImportError as e_utils_imp:
    logger.critical(f"Could not import from utils (license_util or telemetry): {e_utils_imp}\n{traceback.format_exc()}")


    # Define dummy functions if imports fail, so the app can try to show an error
    def get_machine_guid():
        return "dummy_guid_import_failed"


    def send_event(event_name, payload):  # This dummy send_event will be used if import fails
        logger.warning(f"Telemetry disabled: Could not send event '{event_name}' due to import error.")
        pass


    messagebox.showerror("Critical Error",
                         f"Failed to load core utilities: {e_utils_imp}. Application cannot continue.")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# Global exception hook
# ─────────────────────────────────────────────────────────────────────────────
def global_exception_hook(exc_type, exc_value, exc_tb):
    tb_str = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    logger.error(f"Global unhandled exception: {tb_str}")
    try:
        # KEPT: Telemetry for global errors
        send_event("error", {
            "module": "global_exception_hook",
            "error_type": str(exc_type.__name__),
            "error_message": str(exc_value),
            "traceback": tb_str
        })
    except Exception as telemetry_ex:
        logger.error(f"Failed to send telemetry for global exception: {telemetry_ex}")
    sys.__excepthook__(exc_type, exc_value, exc_tb)
    messagebox.showerror("Unhandled Exception", f"An unexpected error occurred: {exc_value}\nDetails have been logged.")


sys.excepthook = global_exception_hook


# ─────────────────────────────────────────────────────────────────────────────
# License check - MODIFIED for Reduced Telemetry & Enhanced Failure Reporting
# ─────────────────────────────────────────────────────────────────────────────
def require_valid_license():
    SUPPORT_EMAIL = "nihasbabu.t3@gmail.com"
    SUPPORT_PHONE = "+91-7558057790 (WhatsApp)"

    def show_generic_license_error_and_exit(failure_reason_for_telemetry="Unknown license failure"):
        # KEPT: Telemetry for license check failed
        send_event("license_check_failed", {
            "reason": failure_reason_for_telemetry,
            "machine_guid_attempted": get_machine_guid()  # Send the machine GUID that failed
        })
        messagebox.showerror("License Error",
                             f"This app is not licensed to run on this machine. Please contact support:\n\n"
                             f"Email: {SUPPORT_EMAIL}\n"
                             f"Phone: {SUPPORT_PHONE}")
        sys.exit(1)

    if getattr(sys, "frozen", False):
        app_dir = os.path.dirname(sys.executable)
        lic_path = os.path.join(app_dir, "license.json")
    else:
        lic_path = resource_path(os.path.join("config", "license.json"))
    logger.info(f"Attempting to load license from: {lic_path}")
    allowed_guid = ""
    try:
        with open(lic_path, "r") as f:
            lic_data = json.load(f)
        allowed_guid = lic_data.get("machine_guid", "").strip().lower()
    except FileNotFoundError:
        reason = f"License file not found: {lic_path}"
        logger.error(f"License Error - {reason}")
        show_generic_license_error_and_exit(reason)
    except json.JSONDecodeError as e_json_decode:
        reason = f"Could not parse license file (invalid JSON): {lic_path}"
        logger.error(f"License Error - {reason}\n{e_json_decode}\n{traceback.format_exc()}")
        show_generic_license_error_and_exit(f"{reason} - Details: {e_json_decode}")
    except Exception as e:
        reason = f"Could not read license file: {lic_path}"
        logger.error(f"License Error - {reason}\n{e}\n{traceback.format_exc()}")
        show_generic_license_error_and_exit(f"{reason} - Error: {e}")

    local_guid = get_machine_guid()
    if not local_guid or local_guid == "dummy_guid_import_failed":
        reason = "Could not retrieve this machine's unique identifier."
        logger.error(f"License Error - {reason}")
        show_generic_license_error_and_exit(reason)
    if not allowed_guid:
        reason = f"'machine_guid' in license file ('{lic_path}') is empty or missing."
        logger.warning(f"License Warning - {reason}")
        show_generic_license_error_and_exit(reason)
    if local_guid != allowed_guid:
        reason = f"Machine GUID mismatch. Local: {local_guid}, Allowed: {allowed_guid}"
        logger.error(f"License Error - {reason}")
        show_generic_license_error_and_exit(reason)

    logger.info(f"License check passed for machine GUID: {local_guid}")
    # REMOVED: send_event("license_check_passed", ...)


# ─────────────────────────────────────────────────────────────────────────────
# Import UI classes and Recon script
# ─────────────────────────────────────────────────────────────────────────────
try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None; ImageTk = None; logger.warning("Pillow (PIL) library not found.")
try:
    logger.info("Attempting to import UI modules and Recon...")
    from gstr1_ui import GSTR1ProcessorUI
    from gstr3b_ui import GSTR3BProcessorUI
    from gstr2b_ui import GSTR2BProcessorUI
    from sales_purchase_ui import SalesPurchaseProcessorUI
    from credit_debit_ui import CreditDebitNoteProcessorUI
    import Recon

    logger.info("Successfully imported UI modules and Recon.")
except ImportError as e_ui_imp:
    logger.critical(f"Failed to import UI/Recon: {e_ui_imp}\n{traceback.format_exc()}")
    messagebox.showerror("Application Error", f"Failed to load essential components: {e_ui_imp}")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# GSTLandingUI class
# ─────────────────────────────────────────────────────────────────────────────
class GSTLandingUI:
    def __init__(self, root):
        self.root = root;
        self.root.title("GST Processor");
        self.root.geometry("300x450")
        logo_relative_path_in_src = os.path.join("assets", "gst_logo.png")
        if getattr(sys, "frozen", False):
            logo_path_resolved = os.path.join(sys._MEIPASS, "assets", "gst_logo.png")
        else:
            logo_path_resolved = resource_path(logo_relative_path_in_src)
        logger.info(f"Attempting to load logo from: {logo_path_resolved}")
        if Image and ImageTk:
            try:
                img = Image.open(logo_path_resolved).resize((100, 100),
                                                            Image.Resampling.LANCZOS); self.logo = ImageTk.PhotoImage(
                    img); tk.Label(self.root, image=self.logo).pack(pady=10)
            except FileNotFoundError:
                logger.warning(f"Logo file not found at {logo_path_resolved}"); tk.Label(self.root,
                                                                                         text="Logo not found",
                                                                                         fg="red").pack(pady=10)
            except Exception as e_logo:
                logger.warning(f"Could not load logo: {e_logo}\n{traceback.format_exc()}"); tk.Label(self.root,
                                                                                                     text="Logo error",
                                                                                                     fg="red").pack(
                    pady=10)
        else:
            tk.Label(self.root, text="GST Processor", font=("Arial", 10)).pack(pady=10)
        tk.Label(self.root, text="GST Processor", font=("Arial", 16, "bold")).pack(pady=10)
        btn_cfg = [("Process GSTR-1", self.open_gstr1_ui), ("Process GSTR-3B", self.open_gstr3b_ui),
                   ("Process GSTR-2B", self.open_gstr2b_ui), ("Process Sales / Purchase", self.open_sales_purchase_ui),
                   ("Process Credit / Debit Notes", self.open_credit_debit_ui),
                   ("Reconciliation", self.run_reconciliation_script), ]
        for txt, cmd in btn_cfg: tk.Button(self.root, text=txt, font=("Arial", 12), command=cmd, width=25).pack(pady=5)

    def _open_processor(self, ProcessorUIClass, title="Processor"):
        if ProcessorUIClass is None: logger.error(f"UI module for {title} not loaded."); messagebox.showerror("Error",
                                                                                                              f"The {title} module could not be loaded."); return
        try:
            top = tk.Toplevel(self.root);
            top.title(title);
            top.transient(self.root);
            top.grab_set();
            self.root.update_idletasks();
            main_x, main_y = self.root.winfo_x(), self.root.winfo_y();
            main_w, main_h = self.root.winfo_width(), self.root.winfo_height();
            top.update_idletasks();
            top_w, top_h = top.winfo_reqwidth(), top.winfo_reqheight();
            min_width, min_height = 500, 480;
            final_w, final_h = max(top_w, min_width), max(top_h, min_height);
            x_pos = main_x + (main_w // 2) - (final_w // 2);
            y_pos = main_y + (main_h // 2) - (final_h // 2);
            top.geometry(f"{final_w}x{final_h}+{x_pos}+{y_pos}");
            ProcessorUIClass(top);
            top.wait_window()
        except Exception as e_proc_ui:
            logger.error(f"Error opening {title}: {e_proc_ui}\n{traceback.format_exc()}"); messagebox.showerror(
                "UI Error", f"Could not open {title}: {e_proc_ui}");

    # REMOVED: send_event("ui_open_attempt", ...) from all open_X_ui methods and run_reconciliation_script
    def open_gstr1_ui(self):
        self._open_processor(GSTR1ProcessorUI, title="GSTR-1 Processor")

    def open_gstr3b_ui(self):
        self._open_processor(GSTR3BProcessorUI, title="GSTR-3B Processor")

    def open_gstr2b_ui(self):
        self._open_processor(GSTR2BProcessorUI, title="GSTR-2B Processor")

    def open_sales_purchase_ui(self):
        self._open_processor(SalesPurchaseProcessorUI, title="Sales/Purchase Processor")

    def open_credit_debit_ui(self):
        self._open_processor(CreditDebitNoteProcessorUI, title="Credit/Debit Note Processor")

    def run_reconciliation_script(self):
        logger.info("Attempting to run Reconciliation script...")
        try:
            Recon.main(); logger.info("Reconciliation script finished.")
        except Exception as e_recon:
            logger.error(f"Error running Recon script: {e_recon}\n{traceback.format_exc()}");
            messagebox.showerror("Reconciliation Error", f"Error: {e_recon}");
            # KEPT: Telemetry for specific error in Recon script
            send_event("error",
                       {"module": "Recon.main", "error_message": str(e_recon), "traceback": traceback.format_exc()})


# ─────────────────────────────────────────────────────────────────────────────
# Main execution block
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if os.environ.get('GST_PROCESSOR_MAIN_RUNNING') == 'true': sys.exit(0)
    os.environ['GST_PROCESSOR_MAIN_RUNNING'] = 'true'
    require_valid_license()
    logger.info(f"Application starting. Python: {sys.version}. PID: {os.getpid()}. App Version: {CURRENT_APP_VERSION}")
    logger.info(f"CWD: {os.getcwd()}");
    logger.info(f"Actual app root for modules: {actual_app_root_for_modules}");
    logger.info(f"Modules directory set to: {modules_dir}");
    logger.info(f"App base directory (sys._MEIPASS or script dir): {app_base_dir}");
    logger.info(f"sys.path: {sys.path}")

    # REMOVED: send_event("app_start", ...)

    root = tk.Tk()
    try:
        app = GSTLandingUI(root); root.mainloop()
    except Exception as e_mainloop:
        logger.critical(f"Fatal error: {e_mainloop}\n{traceback.format_exc()}");
        messagebox.showerror("Fatal Application Error", f"A critical error occurred: {e_mainloop}");
        # KEPT: Telemetry for app crash
        send_event("app_crash",
                   {"error_message": str(e_mainloop), "stage": "mainloop", "traceback": traceback.format_exc()})
    finally:
        logger.info("Application finished.");
        # REMOVED: send_event("app_stop", {})
        if 'GST_PROCESSOR_MAIN_RUNNING' in os.environ: del os.environ['GST_PROCESSOR_MAIN_RUNNING']
