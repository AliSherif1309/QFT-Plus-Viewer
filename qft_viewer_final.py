# -*- coding: utf-8 -*-
import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont, simpledialog, colorchooser
import traceback
import pandas as pd
from reportlab.lib import colors as reportlab_colors # Renamed to avoid clash
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import Color # Needed for custom RGB colors in PDF
import datetime
import platform
from operator import itemgetter
import csv
import xlsxwriter
import sqlite3
import json
import time

# --- Configuration & Helpers (From Script 1) ---

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        # Use base_path relative to the executable location
        base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.abspath(".")
        resource_dir = os.path.join(base_path, "resources") # Look in a 'resources' subfolder
    except Exception:
        # Fallback if finding executable path fails
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "resources"))
        resource_dir = base_path

    # Ensure the resource directory exists if needed (optional)
    # os.makedirs(resource_dir, exist_ok=True)

    return os.path.join(resource_dir, relative_path)


# Use placeholder paths if logos don't exist or cannot be found
DEFAULT_LEFT_LOGO = resource_path("left_logo.png")
DEFAULT_RIGHT_LOGO = resource_path("right_logo.png")

# Check if logo files exist, otherwise use None or a placeholder mechanism
LEFT_LOGO_PATH = DEFAULT_LEFT_LOGO if os.path.exists(DEFAULT_LEFT_LOGO) else None
RIGHT_LOGO_PATH = DEFAULT_RIGHT_LOGO if os.path.exists(DEFAULT_RIGHT_LOGO) else None
# Ensure app_icon.ico is also in the resources folder or adjust path
APP_ICON_PATH = resource_path("app_icon.ico")


def get_app_data_dir():
    """Get the appropriate application data directory for settings and database."""
    try:
        # Prioritize directory next to executable for portability if frozen
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
            app_dir = os.path.join(base_dir, "qft_viewer_data")
        # Otherwise use platform-specific user data directory
        elif platform.system() == "Windows":
            app_data_root = os.getenv('LOCALAPPDATA')
            app_dir = os.path.join(app_data_root, "QFTPlusViewer")
        elif platform.system() == "Darwin": # macOS
            app_data_root = os.path.expanduser('~/Library/Application Support')
            app_dir = os.path.join(app_data_root, "QFTPlusViewer")
        else: # Linux and other Unix-like
            app_data_root = os.getenv('XDG_DATA_HOME', os.path.expanduser("~/.local/share"))
            app_dir = os.path.join(app_data_root, "QFTPlusViewer")

        os.makedirs(app_dir, exist_ok=True)
        print(f"Using App Data Dir: {app_dir}") # Debug print
        return app_dir
    except Exception as e:
        print(f"Error creating/accessing app data directory: {str(e)}")
        # Fallback to current directory if primary method fails
        fallback_dir = os.path.join(os.path.abspath("."), "qft_viewer_data")
        print(f"Falling back to data directory: {fallback_dir}")
        os.makedirs(fallback_dir, exist_ok=True)
        return fallback_dir

SETTINGS_FILE = os.path.join(get_app_data_dir(), 'qft_settings.json')
DATABASE_FILE = os.path.join(get_app_data_dir(), 'qft_database.db') # Match DB name from script 2 if intended

def get_database_connection():
    """Get a connection to the database, creating it if necessary."""
    try:
        conn = sqlite3.connect(DATABASE_FILE)
        cursor = conn.cursor()
        # Create sessions table (ensure unique constraint on name)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sessions (
                session_id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_name TEXT NOT NULL UNIQUE, -- Ensure names are unique
                import_date TEXT NOT NULL,
                last_modified TEXT NOT NULL
            )
        ''')
        # Create results table (ensure foreign key cascade)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS results (
                result_id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id INTEGER,
                barcode TEXT NOT NULL,
                nil_result TEXT,
                tb1_result TEXT,
                tb2_result TEXT,
                mit_result TEXT,
                tb1_nil TEXT,
                tb2_nil TEXT,
                mit_nil TEXT,
                qft_result TEXT,
                requested_date TEXT,
                FOREIGN KEY (session_id) REFERENCES sessions (session_id) ON DELETE CASCADE -- Auto-delete results
            )
        ''')
        # Index for faster searching
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_results_barcode ON results (barcode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_results_session ON results (session_id)')

        conn.commit()
        return conn
    except sqlite3.Error as e:
        print(f"Database connection/setup error: {str(e)}")
        messagebox.showerror("Database Error", f"Failed to connect to or initialize database: {str(e)}\nDatabase path: {DATABASE_FILE}")
        return None
    except Exception as e:
        print(f"Unexpected error getting DB connection: {str(e)}")
        messagebox.showerror("Unexpected Error", f"An unexpected error occurred with the database: {str(e)}")
        return None

def load_settings():
    """Load settings from JSON file, providing defaults."""
    try:
        with open(SETTINGS_FILE, 'r') as f:
            settings = json.load(f)
        # Define defaults including WP
        defaults = {
            'pos_bg': '#FFFFE0',  # Light Yellow (like script 2 default for POS)
            'neg_bg': '#FFFFFF',  # White
            'ind_bg': '#FFFFFF',  # White
            'wp_bg':  '#FFF8DC',  # Cornsilk (light yellowish-orange) - Distinct WP BG
            'pos_text': '#e53935', # Red
            'neg_text': '#43a047', # Green
            'ind_text': '#fb8c00', # Orange
            'wp_text':  '#D2691E', # Chocolate (darker orange/brown) - Distinct WP text
            'decimal_places': 'default'
        }
        # Update loaded settings with defaults for any missing keys
        updated_settings = defaults.copy()
        updated_settings.update(settings)
        # Ensure WP keys exist after update, adding if somehow missed
        if 'wp_bg' not in updated_settings: updated_settings['wp_bg'] = defaults['wp_bg']
        if 'wp_text' not in updated_settings: updated_settings['wp_text'] = defaults['wp_text']
        return updated_settings
    except (FileNotFoundError, json.JSONDecodeError, TypeError) as e:
        print(f"Settings file not found or invalid ({e}), using defaults.")
        # Return and save default settings if file is missing or corrupt
        defaults = {
            'pos_bg': '#FFFFE0', 'neg_bg': '#FFFFFF', 'ind_bg': '#FFFFFF', 'wp_bg': '#FFF8DC',
            'pos_text': '#e53935', 'neg_text': '#43a047', 'ind_text': '#fb8c00', 'wp_text': '#D2691E',
            'decimal_places': 'default'
        }
        save_settings(defaults) # Save defaults for next time
        return defaults
    except Exception as e:
        print(f"Unexpected error loading settings: {str(e)}")
        messagebox.showerror("Settings Error", f"Could not load settings: {e}. Using defaults.")
        # Return safe defaults
        return {
            'pos_bg': '#FFFFE0', 'neg_bg': '#FFFFFF', 'ind_bg': '#FFFFFF', 'wp_bg': '#FFF8DC',
            'pos_text': '#e53935', 'neg_text': '#43a047', 'ind_text': '#fb8c00', 'wp_text': '#D2691E',
            'decimal_places': 'default'
        }

def save_settings(settings_to_save=None):
    """Save current settings to JSON file."""
    global app_settings # Use the global settings variable
    if settings_to_save is None:
        settings_to_save = app_settings # Save the current global settings if none provided

    try:
        # Ensure decimal places is saved correctly
        if 'decimal_places' not in settings_to_save:
            settings_to_save['decimal_places'] = 'default' # Add if missing

        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings_to_save, f, indent=4) # Use indent for readability
        # Update the global settings variable after saving
        app_settings = settings_to_save
        print(f"Settings saved successfully: {settings_to_save}")
    except Exception as e:
        print(f"Error saving settings: {str(e)}")
        messagebox.showerror("Settings Error", f"Could not save settings: {e}")

# Use Script 1's more robust formatting function
def format_number_with_decimals(value_str, decimals_setting):
    """Format number string based on decimal setting, preserving special cases."""
    if not isinstance(value_str, str):
        value_str = str(value_str) # Ensure it's a string

    value_str = value_str.strip()

    if not value_str or value_str == " ":
        return " " # Return empty or space as is

    if '>' in value_str or '<' in value_str:
        return value_str # Return comparison strings as is

    try:
        num = float(value_str)
        if decimals_setting == 'default':
             # Return original float string e.g., "10.5", or int string e.g. "10"
             if num == int(num):
                 return str(int(num))
             else:
                 return value_str
        else:
            try:
                decimals = int(decimals_setting)
                # Ensure we handle potential precision issues with round() before formatting
                # Or simply use f-string formatting which handles rounding
                return f"{num:.{decimals}f}" # Format to specific decimals
            except ValueError:
                 # Fallback to original if decimal setting is invalid
                 return value_str
    except (ValueError, TypeError):
        return value_str # Return original if conversion fails

# Use Script 1's comment calculation logic, but adapt if Script 2's PDF summary needs a specific variation
def calculate_comment(row_dict):
    """Calculates the comment (WP, High Nil, Low Mit) based on a dictionary of values."""
    comment = ""
    qft_result = str(row_dict.get('qft_result', '')).upper() # Ensure uppercase

    try:
        # Use .get() with default 0.0 to handle missing keys gracefully
        nil_val_str = str(row_dict.get('nil_result', '0')).replace('>', '').replace('<', '').strip()
        nil_val = float(nil_val_str) if nil_val_str else 0.0

        tb1_nil_str = str(row_dict.get('tb1_nil', '0')).replace('>', '').replace('<', '').strip()
        tb1_nil = float(tb1_nil_str) if tb1_nil_str else 0.0

        tb2_nil_str = str(row_dict.get('tb2_nil', '0')).replace('>', '').replace('<', '').strip()
        tb2_nil = float(tb2_nil_str) if tb2_nil_str else 0.0

        mit_nil_str = str(row_dict.get('mit_nil', '0')).replace('>', '').replace('<', '').strip()
        mit_nil = float(mit_nil_str) if mit_nil_str else 0.0

        if qft_result in ('POS', 'POS*'):
            if tb1_nil >= 1.0 or tb2_nil >= 1.0:
                comment = "" # Strong positive (no explicit comment needed, implied)
                # For Script 2's summary, we might need to differentiate here later
                # if tb1_nil >= 1.0 and tb2_nil >= 1.0: comment = "Strong (Both)" # Example if needed
                # elif tb1_nil >= 1.0: comment = "Strong (TB1)"
                # elif tb2_nil >= 1.0: comment = "Strong (TB2)"
            else: # Check for Weak Positive (WP)
                is_wp_tb1 = 0.35 <= tb1_nil < 1.0
                is_wp_tb2 = 0.35 <= tb2_nil < 1.0
                if is_wp_tb1 and is_wp_tb2:
                    comment = "WP (Both)"
                elif is_wp_tb1:
                    comment = "WP (TB1)"
                elif is_wp_tb2:
                    comment = "WP (TB2)"

        elif qft_result == 'IND':
            # Check Indeterminate reasons - matching Script 2's logic shown in its PDF export
            if nil_val > 8.0:
                comment = "High Nil"
            # Check Low Mitogen only if Nil is not high (Script 2's simplified Low Mit rule)
            elif mit_nil < 0.5:
                 comment = "Low Mit"
                 # Original Script 2 PDF Export Low Mit logic was more complex, let's use the simpler Mit-Nil < 0.5 first
                 # tb1_cond = tb1_nil < 0.35 and tb1_nil < (0.25 * nil_val) # Script 2's PDF calculation logic
                 # tb2_cond = tb2_nil < 0.35 and tb2_nil < (0.25 * nil_val) # Script 2's PDF calculation logic
                 # if tb1_cond and tb2_cond and mit_nil < 0.5: # Check Mit condition too
                 #     comment = "Low Mit"

    except (ValueError, TypeError) as e:
        print(f"Error calculating comment for row {row_dict.get('barcode', 'N/A')}: {e}")
        comment = "Error" # Indicate calculation error

    return comment

# --- GUI Classes (Keep Script 1's advanced versions) ---

class SplashScreen(tk.Toplevel):
    """A modern splash screen (Adapted from Script 1, visually similar to Script 2)."""
    def __init__(self, parent):
        tk.Toplevel.__init__(self, parent)
        self.title("Loading...")
        self.configure(bg='#FFFFFF')
        self.overrideredirect(True) # No window decorations
        self.attributes('-alpha', 0.0) # Start fully transparent

        width = 400 # Size from Script 2
        height = 250 # Size from Script 2
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

        # Use a themed frame for better consistency
        main_frame = ttk.Frame(self, style='Splash.TFrame', borderwidth=1, relief="solid")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1) # Pack inside window

        # Add Title and Version (like Script 2)
        ttk.Label(main_frame, text="QFT-Plus Viewer", font=('Arial', 16, 'bold'), style='Splash.TLabel').pack(pady=(20, 5))
        ttk.Label(main_frame, text="Version 2.0", font=('Arial', 10), style='Splash.TLabel').pack() # Update version

        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', length=300) # Length from Script 2
        self.progress.pack(pady=(25, 10))
        self.progress.start(15)

        self.status_label = ttk.Label(main_frame, text="Initializing...", font=('Arial', 11), style='Splash.TLabel')
        self.status_label.pack(pady=(5, 20))

        self.display_duration = 2000 # Milliseconds (2 seconds) - Adjust as needed
        self.start_time = time.time()

        self.fade_in()

    def fade_in(self):
        alpha = self.attributes('-alpha')
        if alpha < 1.0:
            alpha = min(alpha + 0.1, 1.0) # Increase alpha, capped at 1.0
            self.attributes('-alpha', alpha)
            self.after(30, self.fade_in) # Adjust speed of fade-in (milliseconds)
        else:
            # Ensure minimum display time
            elapsed = int((time.time() - self.start_time) * 1000)
            remaining_time = self.display_duration - elapsed
            if remaining_time > 0:
                self.after(remaining_time, self.close_splash)
            else:
                self.close_splash()

    def close_splash(self):
        self.progress.stop()
        self.destroy()
        if self.master:
             self.master.deiconify() # Show the main window

class CustomContextMenu(tk.Menu):
    """Custom context menu for Text widget (From Script 1)."""
    def __init__(self, parent_widget):
        super().__init__(parent_widget, tearoff=0,
                         bg='#FFFFFF', fg='#333333',
                         activebackground='#0078D4', activeforeground='white',
                         font=('Segoe UI', 9))
        self.parent_widget = parent_widget
        self.add_command(label="Copy", accelerator="Ctrl+C", command=self._copy)
        self.add_separator()
        self.add_command(label="Select All", accelerator="Ctrl+A", command=self._select_all)

    def _copy(self):
        try:
            self.parent_widget.event_generate("<<Copy>>")
        except tk.TclError:
            pass # Ignore error if no selection

    def _select_all(self):
        try:
            # Check if widget supports tagging (like tk.Text)
            if hasattr(self.parent_widget, 'tag_add'):
                self.parent_widget.tag_add(tk.SEL, "1.0", tk.END)
                self.parent_widget.mark_set(tk.INSERT, "1.0")
                self.parent_widget.see(tk.INSERT)
            # Add handling for other widget types if needed (e.g., Entry)
            elif hasattr(self.parent_widget, 'select_range'):
                 self.parent_widget.select_range(0, tk.END)
                 self.parent_widget.icursor(tk.END)

        except tk.TclError as e:
            print(f"Select all error: {e}")

    def show(self, event):
        """Post the menu at the event location."""
        try:
            # Disable/enable copy based on selection
            has_selection = False
            if isinstance(self.parent_widget, tk.Text):
                has_selection = bool(self.parent_widget.tag_ranges(tk.SEL))
            elif isinstance(self.parent_widget, ttk.Entry) or isinstance(self.parent_widget, tk.Entry):
                has_selection = bool(self.parent_widget.selection_present())

            if has_selection:
                 self.entryconfig("Copy", state="normal")
            else:
                 self.entryconfig("Copy", state="disabled")

            self.post(event.x_root, event.y_root)
        except tk.TclError as e:
            print(f"Show context menu error: {e}")

# --- Core Application Logic (Mostly from Script 1, modified Import if needed) ---

def import_data(add_mode=False):
    """Imports data from Excel or CSV files, either replacing or adding (Based on Script 1)."""
    global main_app # Access the main application instance

    if not add_mode and main_app.has_data():
        confirm = messagebox.askyesnocancel(
            "Replace Data",
            "Importing will replace the current data.\nDo you want to save the current session first?",
            icon='warning', parent=main_app.master # Added parent
        )
        if confirm is None: # Cancel
            return
        elif confirm: # Yes, save first
            if not save_session(auto_save=False): # If manual save cancelled
                 return # Don't proceed with import
        # If 'No' or save was successful, proceed

    filenames = filedialog.askopenfilenames(
        title="Select Excel or CSV file(s)" if not add_mode else "Select file(s) to Add",
        filetypes=(
            ("Supported files", "*.xlsx *.xls *.csv"),
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        )
    )

    if not filenames:
        return

    main_app.update_status("Importing data...", show_progress=True)
    imported_df_list = [] # Collect dataframes
    success_count = 0
    error_files = []

    try:
        # --- Loop through selected files ---
        for filename in filenames:
            try:
                print(f"\n--- Processing File: {os.path.basename(filename)} ---")

                if filename.lower().endswith('.csv'):
                    # Try common delimiters; adjust if needed
                    try:
                        # Use Script 2's assumption of ';' first, then try ','
                        df = pd.read_csv(filename, sep=';', dtype=str, keep_default_na=False, low_memory=False)
                        print(f"Read CSV with ';'. Shape: {df.shape}")
                        # Basic check: if only one column and more than one potential delimiter, try comma
                        if df.shape[1] <= 1 and (';' in df.columns[0] or ',' in df.columns[0]):
                             print("Warning: CSV might have wrong delimiter (read as one column). Trying ','.")
                             df = pd.read_csv(filename, sep=',', dtype=str, keep_default_na=False, low_memory=False)
                             print(f"Read CSV with ','. Shape: {df.shape}")
                    except Exception as read_err:
                        print(f"Error reading CSV {filename}: {read_err}")
                        error_files.append(f"{os.path.basename(filename)} (Read Error)")
                        continue # Skip to next file
                else: # Excel
                    try:
                         df = pd.read_excel(filename, dtype=str, keep_default_na=False)
                         print(f"Read Excel. Shape: {df.shape}")
                    except Exception as read_err:
                         print(f"Error reading Excel {filename}: {read_err}")
                         error_files.append(f"{os.path.basename(filename)} (Read Error)")
                         continue # Skip to next file

                # Check if DataFrame is empty after reading
                if df.empty:
                    print(f"Skipping {os.path.basename(filename)}: File is empty or could not be read properly.")
                    error_files.append(f"{os.path.basename(filename)} (Empty or Read Issue)")
                    continue

                # --- Data Validation/Cleaning ---
                print(f"Columns found: {list(df.columns)}")
                # Make sure required_cols matches the actual expected columns from LQS output
                # These names seem consistent between the scripts' examples
                required_cols = ['Barcode', 'RequestedDate', 'Nil_ReceivedResult', 'TB1_ReceivedResult', 'TB2_ReceivedResult', 'Mitogeno_ReceivedResult', 'DifferenceTB1_Nil', 'DifferenceTB2_Nil', 'DifferenceMitogen_Nil', 'Quantiferon_Result']
                missing_cols = [col for col in required_cols if col not in df.columns]
                if missing_cols:
                     # Try alternate spelling Mitogen vs Mitogeno
                     if 'Mitogen_ReceivedResult' in df.columns and 'Mitogeno_ReceivedResult' in missing_cols:
                          df.rename(columns={'Mitogen_ReceivedResult':'Mitogeno_ReceivedResult'}, inplace=True)
                          missing_cols.remove('Mitogeno_ReceivedResult')
                          print("Info: Renamed 'Mitogen_ReceivedResult' to 'Mitogeno_ReceivedResult'.")
                     if 'DifferenceMitogen_Nil' not in df.columns and 'DifferenceMitogeno_Nil' in df.columns:
                         df.rename(columns={'DifferenceMitogeno_Nil': 'DifferenceMitogen_Nil'}, inplace=True)
                         missing_cols.remove('DifferenceMitogen_Nil')
                         print("Info: Renamed 'DifferenceMitogeno_Nil' to 'DifferenceMitogen_Nil'.")

                     # Recheck missing after potential renames
                     missing_cols = [col for col in required_cols if col not in df.columns]
                     if missing_cols:
                         print(f"Skipping {os.path.basename(filename)}: Missing required columns: {missing_cols}")
                         error_files.append(f"{os.path.basename(filename)} (Missing Columns: {', '.join(missing_cols)})")
                         continue

                print(f"First 5 rows (before date conversion):\n{df.head().to_string()}")

                # Convert 'RequestedDate' using Script 2's format
                df['RequestedDate_original'] = df['RequestedDate'] # Keep original for debugging
                df['RequestedDate'] = pd.to_datetime(df['RequestedDate'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
                print(f"Date conversion results (first 5):\n{df[['RequestedDate_original', 'RequestedDate']].head().to_string()}")
                invalid_date_count = df['RequestedDate'].isna().sum()
                if invalid_date_count > 0:
                     print(f"Warning: {invalid_date_count} rows had invalid dates (set to NaT).")
                     # Keep rows with invalid dates for now, handle later or let user know.

                # Clean string columns (like Script 1)
                for col in df.columns:
                   if df[col].dtype == 'object': # Only process string columns
                       df[col] = df[col].fillna(" ").astype(str).str.strip()
                       df[col] = df[col].replace('', ' ')

                imported_df_list.append(df) # Add processed df to list
                success_count += 1
                print(f"Processed {os.path.basename(filename)}. Shape: {df.shape}")

            except Exception as e:
                print(f"Error processing file {filename}: {str(e)}") # Debug print
                error_files.append(f"{os.path.basename(filename)} ({type(e).__name__})")
                main_app.update_status(f"Error processing {os.path.basename(filename)}...")
                traceback.print_exc() # Print full traceback for this error

        # --- After the loop ---
        if imported_df_list:
            imported_df = pd.concat(imported_df_list, ignore_index=True)
            print(f"\n--- Final Processing ---")
            print(f"Total combined shape: {imported_df.shape}")
            print(f"Final columns: {list(imported_df.columns)}")

            # Rename columns to match the expected internal names ('results' table names in DB)
            # Use Script 1's mapping
            rename_map = {
                'Nil_ReceivedResult': 'nil_result',
                'TB1_ReceivedResult': 'tb1_result',
                'TB2_ReceivedResult': 'tb2_result',
                'Mitogeno_ReceivedResult': 'mit_result', # Matches LQS column name used in Script 2 import
                'DifferenceTB1_Nil': 'tb1_nil',
                'DifferenceTB2_Nil': 'tb2_nil',
                'DifferenceMitogen_Nil': 'mit_nil', # Matches LQS column name
                'Quantiferon_Result': 'qft_result',
                'Barcode': 'barcode', # Ensure consistent casing
                'RequestedDate': 'requested_date' # Ensure consistent casing
            }
            # Select only columns we need and rename
            cols_to_keep_and_rename = {k: v for k, v in rename_map.items() if k in imported_df.columns}
            imported_df = imported_df[list(cols_to_keep_and_rename.keys())].rename(columns=cols_to_keep_and_rename)

            print(f"Columns after rename & selection: {list(imported_df.columns)}")
            print(f"Final imported_df head:\n{imported_df.head().to_string()}")

            data_to_set_or_add = imported_df.to_dict('records') # Convert DF to list of dictionaries

            if add_mode:
                # Get existing barcodes to avoid duplicates
                existing_barcodes = set(main_app.get_all_barcodes())
                # Filter new data
                new_rows = [row for row in data_to_set_or_add if str(row.get('barcode','')) not in existing_barcodes]

                if not new_rows:
                    messagebox.showinfo("Add Data", "No new unique records found in the selected file(s).", parent=main_app.master)
                    main_app.update_status("Add data complete (no new records).", hide_progress=True)
                    return # Exit if no new data
                else:
                    print(f"Adding {len(new_rows)} new rows.")
                    main_app.add_data_rows(new_rows) # Add the new rows
                    message = f"{len(new_rows)} new unique record(s) added."
            else: # Replace mode
                print(f"Setting {len(data_to_set_or_add)} rows.")
                main_app.set_data_rows(data_to_set_or_add) # Use the list of dicts
                 # Update imported filename source tracking
                if len(filenames) == 1:
                    main_app.imported_filename_source = os.path.splitext(os.path.basename(filenames[0]))[0] # Base name like Script 2
                else:
                    main_app.imported_filename_source = f"Combined_Data_{len(filenames)}_files" # Like Script 2
                message = f"{len(data_to_set_or_add)} record(s) imported."

            # Display message after data is processed
            if error_files:
               message += "\nCould not process:\n" + "\n".join(error_files)
            messagebox.showinfo("Import Success" if not add_mode else "Add Success", message, parent=main_app.master)

            # Refresh display and apply sorting/settings
            main_app.refresh_display()
            main_app.sort_data() # Apply default sort
            main_app.update_status("Import complete.", hide_progress=True)

        else: # If imported_df_list is empty after loop
             print("\n--- No Data Imported ---")
             message = "No valid data could be imported."
             if error_files:
                 message += "\nErrors occurred in:\n" + "\n".join(error_files)
             messagebox.showerror("Import Error", message, parent=main_app.master)
             main_app.update_status("Import failed.", hide_progress=True)

    except Exception as e:
        main_app.update_status("Import failed.", hide_progress=True)
        messagebox.showerror("Import Error", f"An unexpected error occurred during import: {str(e)}", parent=main_app.master)
        traceback.print_exc()

# Function to show export options (Combine Script 1's dialog with Script 2's style if desired)
# Using Script 1's `show_export_options` as it's already set up with styles.
def show_export_options():
    """Shows a styled dialog to choose the export format."""
    global main_app
    if not main_app.has_data():
        messagebox.showwarning("Export", "No data available to export.")
        return

    export_window = tk.Toplevel(main_app.master)
    export_window.title("Export Options")
    # Adjust size for better layout with descriptions
    dialog_width = 350
    dialog_height = 500 # Increased height
    export_window.geometry(f"{dialog_width}x{dialog_height}")
    export_window.resizable(False, False)
    export_window.transient(main_app.master)
    export_window.grab_set()
    # Use the main app's background color
    export_window.configure(bg='#f0f2f5')

    # Center the dialog
    main_app.center_window(export_window, dialog_width, dialog_height)

    # Main frame using the app's background style
    main_frame = ttk.Frame(export_window, style='TFrame', padding=(20, 15))
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Title using main app's Title style (or similar)
    # Adjust font size slightly smaller for a dialog
    main_app.style.configure('DialogHeader.TLabel', font=('Open Sans', 16, 'bold'),
                             foreground='#1976D2', background='#f0f2f5')
    title_label = ttk.Label(main_frame, text="Export Options", style='DialogHeader.TLabel', anchor='center')
    title_label.pack(pady=(0, 5))

    # Subtitle using main app's Subtitle style
    subtitle_label = ttk.Label(main_frame, text="Choose your desired file format",
                               style='Subtitle.TLabel', anchor='center')
    subtitle_label.pack(pady=(0, 20))

    # Frame for the export buttons
    options_frame = ttk.Frame(main_frame, style='TFrame')
    options_frame.pack(fill=tk.X, expand=True)

    export_options = [
        {'text': "📄 PDF Report", 'desc': "Formatted report with summary", 'command': export_to_pdf},
        {'text': "📊 Excel File", 'desc': "Spreadsheet with colors/formatting", 'command': export_to_excel},
        {'text': "📝 CSV File",   'desc': "Plain text, comma-separated", 'command': export_to_csv}
    ]

    # Define a slightly smaller description style if needed
    main_app.style.configure('DialogDesc.TLabel', font=('Open Sans', 9),
                             foreground='#666666', background='#f0f2f5')

    for option in export_options:
        # Create a frame for each button + description pair
        option_sub_frame = ttk.Frame(options_frame, style='TFrame')
        option_sub_frame.pack(pady=7, fill=tk.X)

        btn = ttk.Button(
            option_sub_frame,
            text=option['text'],
            # Use lambda to ensure dialog closes *before* export starts
            command=lambda cmd=option['command']: (export_window.destroy(), cmd()),
            style='Custom.TButton', # Use main app button style
            width=20 # Adjust width if necessary
        )
        btn.pack()

        desc_label = ttk.Label(
            option_sub_frame,
            text=option['desc'],
            style='DialogDesc.TLabel',
            anchor='center'
        )
        desc_label.pack(pady=(2, 0))

    # Separator
    ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=(0, 5))

    # Cancel button using the alternative style
    cancel_button = ttk.Button(
        main_frame,
        text="Cancel",
        command=export_window.destroy,
        style='Alt.TButton', # Use alternative style
        width=20
    )
    cancel_button.pack(pady=(5, 10))


# --- Export Functions (Combine Script 1's structure with Script 2's formatting) ---

# Helper for color conversion (needed for PDF)
def hex_to_color(hex_color):
    """Converts hex color string to ReportLab Color object."""
    if not hex_color or not hex_color.startswith('#'): return reportlab_colors.black # Default
    hex_color = hex_color.lstrip('#')
    try:
        # Correctly use reportlab_colors.Color for RGB components
        r = int(hex_color[0:2], 16) / 255.0
        g = int(hex_color[2:4], 16) / 255.0
        b = int(hex_color[4:6], 16) / 255.0
        return Color(r, g, b) # Use the imported Color class
    except:
        return reportlab_colors.black # Fallback

def export_to_pdf():
    """Exports the current data view to a formatted PDF file (Script 2 Formatting)."""
    global main_app, app_settings # Need app_settings here
    if not main_app.has_data():
        messagebox.showwarning("Export PDF", "No data available to export.")
        return

    data_to_export = main_app.get_data_for_export()
    if not data_to_export:
        messagebox.showwarning("Export PDF", "No data available to export.")
        return

    default_filename = f"{main_app.imported_filename_source or 'QFT_Report'}.pdf"
    filename = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        title="Save PDF Report As",
        initialfile=default_filename
    )

    if not filename:
        return

    main_app.update_status("Exporting to PDF...", show_progress=True)

    try:
        doc = SimpleDocTemplate(filename, pagesize=landscape(letter),
                                rightMargin=30, leftMargin=30, topMargin=20, bottomMargin=20)
        styles = getSampleStyleSheet()
        story = []

        # --- Color Settings (Load from app_settings) ---
        current_colors = app_settings # Use the global settings directly

        # --- PDF Styles (Define styles including text colors) ---
        title_style = ParagraphStyle('ReportTitle', parent=styles['Heading1'], fontSize=24, alignment=1, spaceAfter=6)
        date_style = ParagraphStyle('ReportDate', parent=styles['Normal'], fontSize=12, alignment=1, spaceAfter=12)
        header_style = ParagraphStyle('TableHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, alignment=1)
        cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontSize=9, alignment=1)
        comment_style = ParagraphStyle('CommentCell', parent=styles['Normal'], fontSize=9, alignment=1, textColor=reportlab_colors.grey)

        # *** NEW: Define QFT result styles with text colors ***
        qft_pos_style = ParagraphStyle('QFT_POS', parent=cell_style, textColor=hex_to_color(current_colors['pos_text']))
        qft_neg_style = ParagraphStyle('QFT_NEG', parent=cell_style, textColor=hex_to_color(current_colors['neg_text']))
        qft_ind_style = ParagraphStyle('QFT_IND', parent=cell_style, textColor=hex_to_color(current_colors['ind_text']))

        summary_title_style = ParagraphStyle('SummaryTitle', parent=styles['h2'], fontSize=14, alignment=1, spaceBefore=20, spaceAfter=10)
        summary_header_style = ParagraphStyle('SummaryHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, alignment=1)
        summary_cell_style = ParagraphStyle('SummaryCell', parent=styles['Normal'], fontSize=10, alignment=1)
        signature_style = ParagraphStyle('Signature', parent=styles['Normal'], fontSize=10, alignment=0, spaceBefore=20)

        # *** NEW: Define summary cell styles with text colors ***
        summary_pos_style = ParagraphStyle('SummaryPOS', parent=summary_cell_style, textColor=hex_to_color(current_colors['pos_text']))
        summary_wp_style = ParagraphStyle('SummaryWP', parent=summary_cell_style, textColor=hex_to_color(current_colors['ind_text'])) # Use IND color for WP
        summary_neg_style = ParagraphStyle('SummaryNEG', parent=summary_cell_style, textColor=hex_to_color(current_colors['neg_text']))
        summary_ind_style = ParagraphStyle('SummaryIND', parent=summary_cell_style, textColor=hex_to_color(current_colors['ind_text']))


        # PDF Header (Like Script 2)
        # ... (header creation code remains the same) ...
        header_content = []
        img_width = 150 # From S2
        img_height = 90 # From S2
        # Left Logo
        if LEFT_LOGO_PATH and os.path.exists(LEFT_LOGO_PATH):
            try:
                 left_img = Image(LEFT_LOGO_PATH, width=img_width, height=img_height)
                 left_img.hAlign = 'LEFT'
                 header_content.append(left_img)
            except Exception as img_err:
                 print(f"Error loading left logo: {img_err}")
                 header_content.append(Paragraph(" ", styles['Normal'])) # Placeholder
        else:
            header_content.append(Paragraph(" ", styles['Normal']))

        # Title (Use S2's text)
        header_content.append(Paragraph("LIASION® QuantiFERON®", title_style)) # Title from S2

        # Right Logo
        if RIGHT_LOGO_PATH and os.path.exists(RIGHT_LOGO_PATH):
            try:
                right_img = Image(RIGHT_LOGO_PATH, width=img_width, height=img_height)
                right_img.hAlign = 'RIGHT'
                header_content.append(right_img)
            except Exception as img_err:
                 print(f"Error loading right logo: {img_err}")
                 header_content.append(Paragraph(" ", styles['Normal']))
        else:
            header_content.append(Paragraph(" ", styles['Normal']))

        # Header table layout from S2
        header_table = Table([header_content], colWidths=[img_width, 450, img_width]) # Widths from S2
        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))


        report_date_str = main_app.get_report_date_str()
        all_results_for_summary = data_to_export
        table_headers_s2 = ['No.', 'Barcode', 'Nil_Result', 'TB1_Result', 'TB2_Result',
                            'Mit_Result', 'TB1_Nil', 'TB2_Nil', 'Mit_Nil', 'QFT_Result', 'Comment']
        col_widths_s2 = [30, 80, 60, 70, 70, 70, 60, 60, 60, 70, 80]

        ROWS_PER_PAGE = 21
        for i in range(0, len(all_results_for_summary), ROWS_PER_PAGE):
            story.append(header_table)
            story.append(Paragraph(f"Report Date: {report_date_str}", date_style))
            story.append(Spacer(1, 10))

            page_data_rows = all_results_for_summary[i : i + ROWS_PER_PAGE]
            table_data = []
            table_data.append([Paragraph(h, header_style) for h in table_headers_s2])

            for row_num_rel, row_dict in enumerate(page_data_rows, 1):
                row_num_abs = i + row_num_rel
                nil_disp = format_number_with_decimals(row_dict.get('nil_result', ' '), app_settings['decimal_places'])
                tb1_disp = format_number_with_decimals(row_dict.get('tb1_result', ' '), app_settings['decimal_places'])
                tb2_disp = format_number_with_decimals(row_dict.get('tb2_result', ' '), app_settings['decimal_places'])
                mit_disp = format_number_with_decimals(row_dict.get('mit_result', ' '), app_settings['decimal_places'])
                tb1_nil_disp = format_number_with_decimals(row_dict.get('tb1_nil', ' '), app_settings['decimal_places'])
                tb2_nil_disp = format_number_with_decimals(row_dict.get('tb2_nil', ' '), app_settings['decimal_places'])
                mit_nil_disp = format_number_with_decimals(row_dict.get('mit_nil', ' '), app_settings['decimal_places'])
                qft_result = str(row_dict.get('qft_result', ' ')).upper()
                comment = calculate_comment(row_dict)

                # *** Select the appropriate QFT style ***
                current_qft_style = cell_style # Default
                if qft_result in ('POS', 'POS*'): current_qft_style = qft_pos_style
                elif qft_result == 'NEG': current_qft_style = qft_neg_style
                elif qft_result == 'IND': current_qft_style = qft_ind_style

                row_values = [
                    Paragraph(str(row_num_abs), cell_style),
                    Paragraph(str(row_dict.get('barcode', ' ')), cell_style),
                    Paragraph(nil_disp, cell_style), Paragraph(tb1_disp, cell_style),
                    Paragraph(tb2_disp, cell_style), Paragraph(mit_disp, cell_style),
                    Paragraph(tb1_nil_disp, cell_style), Paragraph(tb2_nil_disp, cell_style),
                    Paragraph(mit_nil_disp, cell_style),
                    Paragraph(qft_result, current_qft_style), # *** Use the selected style ***
                    Paragraph(comment, comment_style)
                ]
                table_data.append(row_values)

            data_table = Table(table_data, colWidths=col_widths_s2, repeatRows=1)

            # --- PDF Table Styling (Apply BG color, remove TEXTCOLOR) ---
            table_style_commands = [
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('GRID', (0, 0), (-1, -1), 1, reportlab_colors.black),
                ('LINEBELOW', (0, 0), (-1, 0), 2, reportlab_colors.black),
                # Header backgrounds
                ('BACKGROUND', (2, 0), (2, 0), reportlab_colors.Color(0.85, 0.85, 0.85)),
                ('BACKGROUND', (3, 0), (3, 0), reportlab_colors.Color(0.7, 0.9, 0.7)),
                ('BACKGROUND', (4, 0), (4, 0), reportlab_colors.Color(1.0, 0.95, 0.7)),
                ('BACKGROUND', (5, 0), (5, 0), reportlab_colors.Color(0.85, 0.7, 0.9)),
                ('BACKGROUND', (9, 0), (9, 0), reportlab_colors.Color(0.529, 0.808, 0.922)),
                # Cell padding
                ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ]

            # Apply row-specific background colors ONLY
            for r_idx, row_dict in enumerate(page_data_rows, 1):
                qft_result = str(row_dict.get('qft_result', '')).upper()
                row_bg_color = None

                if qft_result in ('POS', 'POS*'):
                    row_bg_color = hex_to_color(current_colors['pos_bg'])
                elif qft_result == 'NEG':
                    row_bg_color = hex_to_color(current_colors['neg_bg'])
                elif qft_result == 'IND':
                    row_bg_color = hex_to_color(current_colors['ind_bg'])

                if row_bg_color:
                    table_style_commands.append(('BACKGROUND', (0, r_idx), (-1, r_idx), row_bg_color))
                # *** REMOVED TEXTCOLOR command here - handled by ParagraphStyle ***

            data_table.setStyle(TableStyle(table_style_commands))
            story.append(data_table)

            if i + ROWS_PER_PAGE < len(all_results_for_summary):
                story.append(PageBreak())


        # --- PDF Summary Page (Apply colored Paragraph styles) ---
        story.append(PageBreak())
        summary_style_s2 = ParagraphStyle('SummaryStyle', parent=styles['Normal'], fontSize=12, spaceAfter=5, alignment=1)
        story.append(Paragraph("Summary Report", summary_title_style))
        story.append(Paragraph(f"Summary Date: {report_date_str}", summary_style_s2))
        story.append(Spacer(1, 20))

        # Calculate summary stats (remains the same)
        total_samples = len(all_results_for_summary)
        positive_results_count = 0 # Total strong positives
        tb1_strong_only = 0
        tb2_strong_only = 0
        both_tb_strong = 0
        wp_total_count = 0
        wp_tb1_only = 0
        wp_tb2_only = 0
        wp_both = 0
        negative_results_count = 0
        indeterminate_results_count = 0
        high_nil_count = 0
        low_mit_count = 0

        for res in all_results_for_summary:
            qft = str(res.get('qft_result', '')).upper()
            comment = calculate_comment(res)

            if qft in ('POS', 'POS*'):
                is_wp = 'WP' in comment
                if is_wp:
                    wp_total_count += 1
                    if 'Both' in comment: wp_both += 1
                    elif 'TB1' in comment: wp_tb1_only += 1
                    elif 'TB2' in comment: wp_tb2_only += 1
                else: # Strong Positive Check
                     try:
                          tb1_nil_str = str(res.get('tb1_nil', '0')).replace('>', '').replace('<', '').strip()
                          tb1_nil = float(tb1_nil_str) if tb1_nil_str and tb1_nil_str != ' ' else 0.0
                          tb2_nil_str = str(res.get('tb2_nil', '0')).replace('>', '').replace('<', '').strip()
                          tb2_nil = float(tb2_nil_str) if tb2_nil_str and tb2_nil_str != ' ' else 0.0

                          is_strong = tb1_nil >= 1.0 or tb2_nil >= 1.0
                          if is_strong:
                              positive_results_count += 1 # Count strong positives
                              if tb1_nil >= 1.0 and tb2_nil >= 1.0: both_tb_strong += 1
                              elif tb1_nil >= 1.0: tb1_strong_only += 1
                              elif tb2_nil >= 1.0: tb2_strong_only += 1
                     except ValueError:
                          pass # Ignore if values aren't numeric

            elif qft == 'NEG':
                negative_results_count += 1
            elif qft == 'IND':
                indeterminate_results_count += 1
                if 'High Nil' in comment: high_nil_count += 1
                elif 'Low Mit' in comment: low_mit_count += 1
        
        summary_pos_style = ParagraphStyle('SummaryPOS', parent=summary_cell_style, textColor=hex_to_color(current_colors['pos_text']))
        summary_wp_style = ParagraphStyle('SummaryWP', parent=summary_cell_style, textColor=hex_to_color(current_colors['wp_text'])) # Use WP text color
        summary_neg_style = ParagraphStyle('SummaryNEG', parent=summary_cell_style, textColor=hex_to_color(current_colors['neg_text']))
        summary_ind_style = ParagraphStyle('SummaryIND', parent=summary_cell_style, textColor=hex_to_color(current_colors['ind_text']))
        
        # Use helper P function with specific colored styles for data row
        def P_header(text): return Paragraph(str(text), summary_header_style)
        def P_data(text): return Paragraph(str(text), summary_cell_style)
        def P_pos(text): return Paragraph(str(text), summary_pos_style)
        def P_wp(text): return Paragraph(str(text), summary_wp_style) # Use defined style
        def P_neg(text): return Paragraph(str(text), summary_neg_style)
        def P_ind(text): return Paragraph(str(text), summary_ind_style)

        summary_data_s2 = [
            [P_header('Total Samples'), P_header('POS Results'), '', '', '', P_header('WP Results'), '', '', '', P_header('NEG Results'), P_header('IND Results'), '', ''],
            ['', P_header('Total'), P_header('TB1'), P_header('TB2'), P_header('Both'), P_header('Total'), P_header('TB1'), P_header('TB2'), P_header('Both'), '', P_header('Total'), P_header('Nil'), P_header('Mit')],
            # Use colored styles for data cells
            [P_data(total_samples),
             P_pos(positive_results_count), P_pos(tb1_strong_only), P_pos(tb2_strong_only), P_pos(both_tb_strong),
             P_wp(wp_total_count), P_wp(wp_tb1_only), P_wp(wp_tb2_only), P_wp(wp_both), # Use P_wp helper
             P_neg(negative_results_count),
             P_ind(indeterminate_results_count), P_ind(high_nil_count), P_ind(low_mit_count)],
            [Paragraph('', styles['Normal']), '', '', '', '', '', '', '', '', '', '', '', ''], # Empty row
            [Paragraph('Name/Signature:', signature_style), Paragraph('_____________________', signature_style), '', '', '', '', '', '', '', '', '', '', '']
        ]


        col_width_s2_summary_main = 90
        col_width_s2_summary_sub = 50
        summary_col_widths = [
            col_width_s2_summary_main, col_width_s2_summary_sub, col_width_s2_summary_sub,
            col_width_s2_summary_sub, col_width_s2_summary_sub, col_width_s2_summary_sub,
            col_width_s2_summary_sub, col_width_s2_summary_sub, col_width_s2_summary_sub,
            col_width_s2_summary_main, col_width_s2_summary_sub, col_width_s2_summary_sub,
            col_width_s2_summary_sub
        ]

        summary_table = Table(summary_data_s2, colWidths=summary_col_widths)
        summary_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12), ('FONTSIZE', (1, 1), (-1, 1), 10),
            ('FONTSIZE', (0, 2), (-1, 2), 10),
            # Spans
            ('SPAN', (1, 0), (4, 0)), ('SPAN', (5, 0), (8, 0)), ('SPAN', (10, 0), (12, 0)),
            ('SPAN', (0, 0), (0, 1)), ('SPAN', (9, 0), (9, 1)),
            # Backgrounds using settings (WP headers use WP BG)
            ('BACKGROUND', (0, 0), (0, 1), reportlab_colors.lightgrey),
            ('BACKGROUND', (1, 0), (4, 1), hex_to_color(current_colors['pos_bg'])),
            ('BACKGROUND', (5, 0), (8, 1), hex_to_color(current_colors['wp_bg'])), # WP headers use WP BG
            ('BACKGROUND', (9, 0), (9, 1), hex_to_color(current_colors['neg_bg'])),
            ('BACKGROUND', (10, 0), (12, 1), hex_to_color(current_colors['ind_bg'])),
            # Grid
            ('GRID', (0, 0), (-1, 2), 0.25, reportlab_colors.grey),
            # Text colors for data row are handled by Paragraph styles now
            # Signature line
            ('ALIGN', (0, 4), (1, 4), 'LEFT'), ('VALIGN', (0, 4), (1, 4), 'MIDDLE'),
            ('SPAN', (1, 4), (4, 4)),
        ]))

        story.append(summary_table)

        # Build PDF
        doc.build(story)
        main_app.update_status("PDF export complete.", hide_progress=True)
        messagebox.showinfo("Export PDF", "PDF report exported successfully!")

        if messagebox.askyesno("Open File", "Do you want to open the exported PDF file?", parent=main_app.master):
            try:
                if platform.system() == 'Darwin': os.system(f'open "{filename}"')
                elif platform.system() == 'Windows': os.startfile(filename)
                else: os.system(f'xdg-open "{filename}"')
            except Exception as open_err:
                 messagebox.showwarning("Open File Error", f"Could not automatically open the file:\n{open_err}", parent=main_app.master)

    except Exception as e:
        main_app.update_status("PDF export failed.", hide_progress=True)
        messagebox.showerror("Export PDF Error", f"Failed to export PDF: {str(e)}")
        traceback.print_exc()

def export_to_excel():
    """Exports the current data view to a formatted Excel file (Script 2 Formatting)."""
    global main_app, app_settings
    if not main_app.has_data():
        messagebox.showwarning("Export Excel", "No data available to export.")
        return

    data_to_export = main_app.get_data_for_export()
    if not data_to_export:
        messagebox.showwarning("Export Excel", "No data available to export.")
        return

    default_filename = f"{main_app.imported_filename_source or 'QFT_Report'}.xlsx"
    filename = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save Excel Report As",
        initialfile=default_filename
    )

    if not filename:
        return

    main_app.update_status("Exporting to Excel...", show_progress=True)

    try:
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet("QFT Results")

        # --- Excel Formats (Add WP format) ---
        current_colors = app_settings
        header_format = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#D9D9D9', 'border': 1})
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        date_format_s2 = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm:ss', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        comment_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'italic': True, 'font_color': '#555555'})

        # Formats for QFT result cell using settings (like S2 approach)
        pos_format = workbook.add_format({
            'font_color': current_colors.get('pos_text', '#e53935'),
            'bg_color': current_colors.get('pos_bg', '#FFFFF0'), # Includes background from settings
            'align': 'center', 'valign': 'vcenter', 'border': 1})
        neg_format = workbook.add_format({
            'font_color': current_colors.get('neg_text', '#43a047'),
            'bg_color': current_colors.get('neg_bg', '#FFFFFF'), # Includes background from settings
            'align': 'center', 'valign': 'vcenter', 'border': 1})
        ind_format = workbook.add_format({
            'font_color': current_colors.get('ind_text', '#fb8c00'),
            'bg_color': current_colors.get('ind_bg', '#FFFFFF'), # Includes background from settings
            'align': 'center', 'valign': 'vcenter', 'border': 1})
        wp_format = workbook.add_format({
            'font_color': current_colors.get('wp_text', '#D2691E'),
            'bg_color': current_colors.get('wp_bg', '#FFF8DC'),
            'align': 'center', 'valign': 'vcenter', 'border': 1})

        # --- Write Headers (Match Script 2's Excel Export) ---
        headers_s2 = ['Barcode', 'Nil_Result', 'TB1_Result', 'TB2_Result', 'Mit_Result',
                      'TB1_Nil', 'TB2_Nil', 'Mit_Nil', 'QFT_Result', 'Comments', 'Requested Date']
        worksheet.write_row(0, 0, headers_s2, header_format)

        # --- Write Data ---
        decimals = app_settings['decimal_places']
        for r_idx, row_dict in enumerate(data_to_export, 1):
            col = 0
            barcode = row_dict.get('barcode', '')
            worksheet.write(r_idx, col, barcode, cell_format); col += 1

            # Numerical values (Nil to Mit-Nil)
            num_keys = ['nil_result', 'tb1_result', 'tb2_result', 'mit_result', 'tb1_nil', 'tb2_nil', 'mit_nil']
            for key in num_keys:
                val_str = format_number_with_decimals(row_dict.get(key, ''), decimals)
                try:
                    if '>' in val_str or '<' in val_str or val_str.strip() == "":
                         worksheet.write_string(r_idx, col, val_str, cell_format)
                    else:
                         worksheet.write_number(r_idx, col, float(val_str), cell_format)
                except ValueError:
                     worksheet.write_string(r_idx, col, val_str, cell_format)
                col += 1

            # QFT Result (apply specific format including WP)
            qft_result = str(row_dict.get('qft_result', '')).upper()
            comment = calculate_comment(row_dict) # Calculate comment
            is_wp = "WP" in comment             # Check if WP

            qft_format_to_use = cell_format # Default
            if is_wp: qft_format_to_use = wp_format          # WP format first
            elif qft_result in ('POS', 'POS*'): qft_format_to_use = pos_format
            elif qft_result == 'NEG': qft_format_to_use = neg_format
            elif qft_result == 'IND': qft_format_to_use = ind_format

            # Apply the determined format to this cell
            worksheet.write(r_idx, col, qft_result, qft_format_to_use); col += 1

            # Comment
            worksheet.write(r_idx, col, comment, comment_format); col += 1 # Use pre-calculated comment

            # Requested Date
            req_date = row_dict.get('requested_date')
            if isinstance(req_date, datetime.datetime):
                worksheet.write_datetime(r_idx, col, req_date, date_format_s2)
            else:
                 worksheet.write(r_idx, col, str(req_date) if req_date else '', cell_format)
            col += 1

        # --- Adjust Column Widths (Match S2's Excel Export) ---
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:I', 10) # B to I covers Nil_Result to QFT_Result
        worksheet.set_column('J:J', 15)
        worksheet.set_column('K:K', 20)
        worksheet.freeze_panes(1, 0)

        workbook.close()
        main_app.update_status("Excel export complete.", hide_progress=True)
        messagebox.showinfo("Export Excel", "Excel file exported successfully!")

        if messagebox.askyesno("Open File", "Do you want to open the exported Excel file?", parent=main_app.master):
             try:
                 if platform.system() == 'Darwin': os.system(f'open "{filename}"')
                 elif platform.system() == 'Windows': os.startfile(filename)
                 else: os.system(f'xdg-open "{filename}"')
             except Exception as open_err:
                  messagebox.showwarning("Open File Error", f"Could not automatically open the file:\n{open_err}", parent=main_app.master)

    except Exception as e:
        main_app.update_status("Excel export failed.", hide_progress=True)
        messagebox.showerror("Export Excel Error", f"Failed to export Excel: {str(e)}")
        traceback.print_exc()

def export_to_csv():
    """Exports the current data view to a CSV file (Script 2 Formatting)."""
    global main_app, app_settings
    if not main_app.has_data():
        messagebox.showwarning("Export CSV", "No data available to export.")
        return

    data_to_export = main_app.get_data_for_export()
    if not data_to_export:
        messagebox.showwarning("Export CSV", "No data available to export.")
        return

    # Use Script 2's filename logic
    default_filename = f"{main_app.imported_filename_source or 'QFT_Report'}.csv"
    filename = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        title="Save CSV Report As",
        initialfile=default_filename
    )

    if not filename:
        return

    main_app.update_status("Exporting to CSV...", show_progress=True)

    try:
        # Match headers from Script 2's CSV export
        headers_s2 = ['Barcode', 'Nil_Result', 'TB1_Result', 'TB2_Result', 'Mit_Result',
                      'TB1_Nil', 'TB2_Nil', 'Mit_Nil', 'QFT_Result', 'Comments', 'Requested Date']
        decimals = app_settings['decimal_places']

        # Use standard comma delimiter unless ';' is strongly preferred
        # Script 2 implicitly used comma with default writer
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile) # Default comma delimiter
            writer.writerow(headers_s2)

            for row_dict in data_to_export:
                # Format data row according to S2 headers/order
                row_values = [
                    row_dict.get('barcode', ''),
                    format_number_with_decimals(row_dict.get('nil_result', ''), decimals),
                    format_number_with_decimals(row_dict.get('tb1_result', ''), decimals),
                    format_number_with_decimals(row_dict.get('tb2_result', ''), decimals),
                    format_number_with_decimals(row_dict.get('mit_result', ''), decimals),
                    format_number_with_decimals(row_dict.get('tb1_nil', ''), decimals),
                    format_number_with_decimals(row_dict.get('tb2_nil', ''), decimals),
                    format_number_with_decimals(row_dict.get('mit_nil', ''), decimals),
                    row_dict.get('qft_result', ''),
                    calculate_comment(row_dict), # Include comment column like S2
                ]
                # Format date for CSV (YYYY-MM-DD HH:MM:SS is a good standard)
                req_date = row_dict.get('requested_date')
                date_str = req_date.strftime('%Y-%m-%d %H:%M:%S') if isinstance(req_date, datetime.datetime) else str(req_date or '')
                row_values.append(date_str)

                writer.writerow(row_values)

        main_app.update_status("CSV export complete.", hide_progress=True)
        messagebox.showinfo("Export CSV", "CSV file exported successfully!")

        # Ask to open file
        if messagebox.askyesno("Open File", "Do you want to open the exported CSV file?", parent=main_app.master):
             try:
                 if platform.system() == 'Darwin': os.system(f'open "{filename}"')
                 elif platform.system() == 'Windows': os.startfile(filename)
                 else: os.system(f'xdg-open "{filename}"')
             except Exception as open_err:
                  messagebox.showwarning("Open File Error", f"Could not automatically open the file:\n{open_err}", parent=main_app.master)

    except Exception as e:
        main_app.update_status("CSV export failed.", hide_progress=True)
        messagebox.showerror("Export CSV Error", f"Failed to export CSV: {str(e)}")
        traceback.print_exc()


# --- Session Management (Keep Script 1's implementation) ---
# Includes save_session, manage_sessions, load_selected_session, etc.

def save_session(auto_save=False, pre_clear=False, session_name_in=None):
    """Saves the current data to the database (Based on Script 1, allows name input)."""
    global main_app
    if not main_app.has_data():
        if not auto_save and not pre_clear: # Only warn on manual save
            messagebox.showwarning("Save Session", "No data to save.")
        return False # Indicate save did not happen

    current_time_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    session_name = session_name_in # Use provided name if available

    if not session_name: # If no name provided (e.g., manual save or auto)
        if auto_save:
            session_name = f"AutoSave_{current_time_str}"
        elif pre_clear:
            session_name = f"PreClearBackup_{current_time_str}"
        else:
            # Manual save, prompt for name
            suggested_name = f"Session_{current_time_str}"
            # Use imported filename as base suggestion (like Script 2)
            if main_app.imported_filename_source and not main_app.imported_filename_source.startswith("session"):
                 suggested_name = f"{main_app.imported_filename_source}_{main_app.get_report_date_str()}"
            elif main_app.imported_filename_source.startswith("session"):
                 suggested_name = f"{main_app.imported_filename_source.replace('session ','')}_Modified_{current_time_str}"


            session_name = simpledialog.askstring(
                "Save Session",
                "Enter a unique name for this session",
                initialvalue=suggested_name,
                parent=main_app.master
            )
            if not session_name:
                return False # User cancelled manual save

    main_app.update_status("Saving session...", show_progress=True)
    conn = None # Ensure conn is defined for finally block
    try:
        conn = get_database_connection()
        if not conn:
            main_app.update_status("Save failed (DB connection).", hide_progress=True)
            return False

        cursor = conn.cursor()
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Check if session name already exists
        cursor.execute("SELECT session_id FROM sessions WHERE session_name = ?", (session_name,))
        existing = cursor.fetchone()

        session_id = None
        if existing:
            session_id = existing[0]
            if auto_save or pre_clear: # Silently overwrite/ignore autosaves/backups if name clashes (rare)
                 print(f"Info: AutoSave/PreClear name '{session_name}' exists. Overwriting.")
                 # Delete existing results for this session before inserting new ones
                 cursor.execute("DELETE FROM results WHERE session_id = ?", (session_id,))
                 # Update the last_modified timestamp for the session
                 cursor.execute("UPDATE sessions SET last_modified = ? WHERE session_id = ?", (timestamp, session_id))
            else: # Manual save
                 confirm_overwrite = messagebox.askyesno(
                      "Overwrite Session?",
                      f"Session name '{session_name}' already exists.\nDo you want to overwrite it with the current data?",
                      icon='warning', parent=main_app.master
                 )
                 if confirm_overwrite:
                      print(f"Info: Overwriting existing session '{session_name}'.")
                      # Delete existing results for this session before inserting new ones
                      cursor.execute("DELETE FROM results WHERE session_id = ?", (session_id,))
                      # Update the last_modified timestamp for the session
                      cursor.execute("UPDATE sessions SET last_modified = ? WHERE session_id = ?", (timestamp, session_id))
                 else:
                      main_app.update_status("Save cancelled (name exists).", hide_progress=True)
                      conn.close()
                      return False # Indicate manual save failed/cancelled overwrite

        else: # Session name is new, insert it
            cursor.execute('''
                INSERT INTO sessions (session_name, import_date, last_modified)
                VALUES (?, ?, ?)
            ''', (session_name, timestamp, timestamp))
            session_id = cursor.lastrowid

        # Get data and insert results
        data_to_save = main_app.get_data_for_export() # Get current data state
        results_to_insert = []
        for row_dict in data_to_save:
            req_date = row_dict.get('requested_date')
            # Use consistent DB date format
            date_str = req_date.strftime('%Y-%m-%d %H:%M:%S') if isinstance(req_date, datetime.datetime) else str(req_date or '')

            results_to_insert.append((
                session_id,
                str(row_dict.get('barcode', '')),
                str(row_dict.get('nil_result', '')),
                str(row_dict.get('tb1_result', '')),
                str(row_dict.get('tb2_result', '')),
                str(row_dict.get('mit_result', '')),
                str(row_dict.get('tb1_nil', '')),
                str(row_dict.get('tb2_nil', '')),
                str(row_dict.get('mit_nil', '')),
                str(row_dict.get('qft_result', '')),
                date_str
            ))

        cursor.executemany('''
            INSERT INTO results (
                session_id, barcode, nil_result, tb1_result, tb2_result, mit_result,
                tb1_nil, tb2_nil, mit_nil, qft_result, requested_date
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', results_to_insert)

        conn.commit()

        if not auto_save and not pre_clear: # Show success only for manual saves
            messagebox.showinfo("Save Session", f"Session '{session_name}' saved successfully!", parent=main_app.master)

        main_app.update_status(f"Session '{session_name}' saved.", hide_progress=True)
        # Update the source tracker if manually saved
        if not auto_save and not pre_clear:
            main_app.imported_filename_source = f"session {session_name}"
        return True # Indicate success

    except sqlite3.Error as db_err:
        if conn: conn.rollback()
        messagebox.showerror("Database Error", f"Failed to save session '{session_name}': {db_err}", parent=main_app.master)
        main_app.update_status("Save failed (DB error).", hide_progress=True)
        traceback.print_exc()
        return False
    except Exception as e:
        if conn: conn.rollback()
        messagebox.showerror("Save Error", f"An unexpected error occurred while saving session '{session_name}': {e}", parent=main_app.master)
        main_app.update_status("Save failed (unexpected error).", hide_progress=True)
        traceback.print_exc()
        return False
    finally:
        if conn:
            conn.close()


def manage_sessions():
    """Opens the session management window (Keep Script 1's implementation)."""
    global main_app
    conn = None
    try:
        conn = get_database_connection()
        if not conn: return

        cursor = conn.cursor()
        # Use Script 1's more detailed query
        cursor.execute('''
            SELECT
                s.session_id, s.session_name, s.import_date, s.last_modified,
                COUNT(r.result_id) as total_samples,
                SUM(CASE WHEN r.qft_result = 'POS' OR r.qft_result = 'POS*' THEN 1 ELSE 0 END) as pos,
                SUM(CASE WHEN r.qft_result = 'NEG' THEN 1 ELSE 0 END) as neg,
                SUM(CASE WHEN r.qft_result = 'IND' THEN 1 ELSE 0 END) as ind,
                MIN(r.requested_date) as earliest_date,
                MAX(r.requested_date) as latest_date
            FROM sessions s
            LEFT JOIN results r ON s.session_id = r.session_id
            GROUP BY s.session_id, s.session_name, s.import_date, s.last_modified
            ORDER BY s.last_modified DESC
        ''')
        sessions_data = cursor.fetchall()

        if not sessions_data:
            messagebox.showinfo("Manage Sessions", "No saved sessions found.", parent=main_app.master)
            conn.close() # Close connection if no sessions
            return

        # --- Session Management Window (Use Script 1's layout) ---
        session_window = tk.Toplevel(main_app.master)
        session_window.title("Manage Saved Sessions")
        session_window.geometry("950x600")
        session_window.transient(main_app.master)
        session_window.grab_set()
        session_window.configure(bg=app_settings.get('dialog_bg', '#F5F5F5'))
        main_app.center_window(session_window, 950, 600)

        top_frame = ttk.Frame(session_window, style='Dialog.TFrame')
        top_frame.pack(pady=10, padx=10, fill='x')
        ttk.Label(top_frame, text="Saved Sessions", style='DialogTitle.TLabel').pack(side='left')

        tree_frame = ttk.Frame(session_window, style='Dialog.TFrame')
        tree_frame.pack(pady=5, padx=10, fill='both', expand=True)

        # --- Treeview Setup (Script 1) ---
        cols = ('name', 'modified', 'samples', 'pos', 'neg', 'ind', 'date_range')
        tree = ttk.Treeview(tree_frame, columns=cols, show='headings', selectmode='browse') # browse = single selection

        col_widths = {'name': 250, 'modified': 140, 'samples': 80, 'pos': 50, 'neg': 50, 'ind': 50, 'date_range': 180}
        col_align = {'samples': 'center', 'pos': 'center', 'neg': 'center', 'ind': 'center'}

        for col in cols:
            tree.heading(col, text=col.replace('_', ' ').title())
            tree.column(col, width=col_widths.get(col, 100), anchor=col_align.get(col, 'w'))

        # Treeview scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side='right', fill='y')
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side='bottom', fill='x')
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side='left', fill='both', expand=True)

        # Populate Treeview
        for row in sessions_data:
            (session_id, name, imp_date, mod_date, total, pos, neg, ind, early_dt, late_dt) = row
            total, pos, neg, ind = (total or 0, pos or 0, neg or 0, ind or 0) # Handle None

            date_range_str = "N/A"
            try:
                # Use the correct DB date format '%Y-%m-%d %H:%M:%S'
                if early_dt and late_dt:
                    e_dt = datetime.datetime.strptime(early_dt, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                    l_dt = datetime.datetime.strptime(late_dt, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                    date_range_str = f"{e_dt} to {l_dt}" if e_dt != l_dt else e_dt
                elif early_dt: # Handle cases where only one date might exist
                     date_range_str = datetime.datetime.strptime(early_dt, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                elif late_dt:
                     date_range_str = datetime.datetime.strptime(late_dt, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
            except (ValueError, TypeError): pass # Ignore parsing errors

            tree.insert('', 'end', values=(name, mod_date, total, pos, neg, ind, date_range_str), tags=(str(session_id),))

        # --- Action Buttons (Script 1) ---
        button_frame = ttk.Frame(session_window, style='Dialog.TFrame')
        button_frame.pack(pady=10, padx=10, fill='x')

        load_button = ttk.Button(button_frame, text="Load Selected", style='Dialog.TButton', state='disabled',
                                 command=lambda: load_selected_session(tree, session_window, conn))
        load_button.pack(side='left', padx=5)

        delete_button = ttk.Button(button_frame, text="Delete Selected", style='DialogAlt.TButton', state='disabled',
                                   command=lambda: delete_selected_session(tree, session_window, conn))
        delete_button.pack(side='left', padx=5)

        rename_button = ttk.Button(button_frame, text="Rename Selected", style='Dialog.TButton', state='disabled',
                                   command=lambda: rename_selected_session(tree, session_window, conn))
        rename_button.pack(side='left', padx=5)

        cancel_button = ttk.Button(button_frame, text="Close", style='Dialog.TButton',
                                 command=lambda: (conn.close(), session_window.destroy()))
        cancel_button.pack(side='right', padx=5)

        # Enable buttons on selection
        def on_select(event):
            if tree.selection():
                load_button['state'] = 'normal'
                delete_button['state'] = 'normal'
                rename_button['state'] = 'normal'
            else:
                load_button['state'] = 'disabled'
                delete_button['state'] = 'disabled'
                rename_button['state'] = 'disabled'

        tree.bind('<<TreeviewSelect>>', on_select)
        tree.bind('<Double-1>', lambda e: load_selected_session(tree, session_window, conn) if tree.selection() else None)

    except sqlite3.Error as db_err:
        messagebox.showerror("Database Error", f"Failed to retrieve sessions: {db_err}", parent=main_app.master)
        if conn: conn.close() # Close connection on error
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open session manager: {e}", parent=main_app.master)
        if conn: conn.close() # Close connection on error


def load_selected_session(tree, session_window, db_conn_passed_in=None):
    """Loads the data from the selected session into the main app (Keep Script 1)."""
    global main_app
    selection = tree.selection()
    if not selection: return

    item = tree.item(selection[0])
    session_id = int(item['tags'][0])
    session_name = item['values'][0]

    if main_app.has_data():
         # Use Script 1's save confirmation logic
         confirm = messagebox.askyesnocancel(
             "Load Session",
             f"Loading '{session_name}' will replace the current data.\nDo you want to save the current session first?",
             icon='warning', parent=session_window
         )
         if confirm is None: return
         elif confirm:
             if not save_session(auto_save=False): return

    main_app.update_status(f"Loading session '{session_name}'...", show_progress=True)
    conn = None
    try:
        conn = db_conn_passed_in or get_database_connection()
        if not conn:
            main_app.update_status("Load failed (DB connection).", hide_progress=True)
            return

        cursor = conn.cursor()
        # Select all necessary columns from 'results' table
        cursor.execute('''
            SELECT barcode, nil_result, tb1_result, tb2_result, mit_result,
                   tb1_nil, tb2_nil, mit_nil, qft_result, requested_date
            FROM results
            WHERE session_id = ?
        ''', (session_id,))
        results_data = cursor.fetchall()

        loaded_rows = []
        # Use correct indices based on the SELECT query
        col_names = ['barcode', 'nil_result', 'tb1_result', 'tb2_result', 'mit_result',
                     'tb1_nil', 'tb2_nil', 'mit_nil', 'qft_result', 'requested_date']
        for row in results_data:
            row_dict = dict(zip(col_names, row)) # Create dict directly

            # Convert date string back to datetime object using DB format
            date_str = row_dict.get('requested_date')
            if date_str and isinstance(date_str, str):
                try:
                    row_dict['requested_date'] = datetime.datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                except (ValueError, TypeError):
                    print(f"Warning: Could not parse date '{date_str}' for barcode {row_dict['barcode']}. Setting to None.")
                    row_dict['requested_date'] = None
            elif not isinstance(date_str, datetime.datetime):
                 row_dict['requested_date'] = None # Handle other non-datetime cases

            loaded_rows.append(row_dict)

        # Update main application data
        main_app.set_data_rows(loaded_rows)
        main_app.imported_filename_source = f"session {session_name}" # Track source
        main_app.refresh_display() # Refresh the view
        main_app.sort_data() # Apply default sort (or last used sort?)

        session_window.destroy() # Close session manager
        main_app.update_status(f"Session '{session_name}' loaded.", hide_progress=True)
        messagebox.showinfo("Load Session", f"Session '{session_name}' loaded successfully!", parent=main_app.master)

    except sqlite3.Error as db_err:
        messagebox.showerror("Database Error", f"Failed to load session data: {db_err}", parent=session_window)
        main_app.update_status("Load failed (DB error).", hide_progress=True)
    except Exception as e:
        messagebox.showerror("Load Error", f"An unexpected error occurred loading session {e}", parent=session_window)
        main_app.update_status("Load failed (unexpected error).", hide_progress=True)
        traceback.print_exc()
    finally:
        if conn and not db_conn_passed_in:
            conn.close()

# Keep delete_selected_session from Script 1
def delete_selected_session(tree, session_window, db_conn):
    """Deletes the selected session from the database after confirmation."""
    selection = tree.selection()
    if not selection: return

    # --- Get session info (no change here) ---
    item = tree.item(selection[0])
    session_id = int(item['tags'][0])
    session_name = item['values'][0]

    confirm = messagebox.askyesno(
        "Confirm Delete",
        f"Are you sure you want to permanently delete the session:\n'{session_name}'?\n\nThis action cannot be undone.",
        icon='warning', parent=session_window
    )

    if confirm:
        try:
            cursor = db_conn.cursor()
            # Delete session (results deleted automatically due to ON DELETE CASCADE)
            cursor.execute('DELETE FROM sessions WHERE session_id = ?', (session_id,))
            db_conn.commit()

            # Remove from treeview
            tree.delete(selection[0])
            messagebox.showinfo("Delete Session", f"Session '{session_name}' deleted successfully.", parent=session_window)

            # Close window if the tree becomes empty
            if not tree.get_children():
                 session_window.destroy()

        except sqlite3.Error as db_err:
            db_conn.rollback()
            messagebox.showerror("Database Error", f"Failed to delete session: {db_err}", parent=session_window)
        except Exception as e:
            db_conn.rollback()
            # Print traceback for unexpected errors during development/debugging
            traceback.print_exc()
            messagebox.showerror("Delete Error", f"An unexpected error occurred: {e}", parent=session_window)
# Keep rename_selected_session from Script 1
def rename_selected_session(tree, session_window, db_conn):
    """Renames the selected session."""
    selection = tree.selection()
    if not selection: return

    item = tree.item(selection[0])
    session_id = int(item['tags'][0])
    old_name = item['values'][0]

    new_name = simpledialog.askstring(
        "Rename Session",
        f"Enter a new unique name for session '{old_name}':",
        initialvalue=old_name,
        parent=session_window
    )

    if not new_name or new_name == old_name:
        return # Cancelled or no change

    try:
        cursor = db_conn.cursor()
        # Check if new name already exists (using UNIQUE constraint)
        cursor.execute("UPDATE sessions SET session_name = ?, last_modified = ? WHERE session_id = ?",
                       (new_name, datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), session_id))
        db_conn.commit()

        # Update treeview
        tree.item(selection[0], values=(new_name,) + item['values'][1:]) # Update name in values tuple
        messagebox.showinfo("Rename Session", f"Session renamed to '{new_name}'.", parent=session_window)

    except sqlite3.IntegrityError: # Catch unique constraint violation
        db_conn.rollback()
        messagebox.showerror("Rename Error", f"Session name '{new_name}' already exists. Please choose a unique name.", parent=session_window)
    except sqlite3.Error as db_err:
        db_conn.rollback()
        messagebox.showerror("Database Error", f"Failed to rename session {db_err}", parent=session_window)
    except Exception as e:
        db_conn.rollback()
        messagebox.showerror("Rename Error", f"An unexpected error occurred: {e}", parent=session_window)

# --- Global Search (Keep Script 1's implementation, including exports) ---
# show_global_search, export_global_search_to_excel, export_global_search_to_csv

def show_global_search():
    """Opens the global search window."""
    global main_app

    search_window = tk.Toplevel(main_app.master)
    search_window.title("Global Sample Search")
    search_window.geometry("1000x700")
    search_window.transient(main_app.master)
    search_window.grab_set()
    search_window.configure(bg=app_settings.get('dialog_bg', '#F5F5F5'))
    main_app.center_window(search_window, 1000, 700)

    # --- Search Input Area ---
    search_frame = ttk.Frame(search_window, style='Dialog.TFrame')
    search_frame.pack(pady=10, padx=10, fill='x')

    ttk.Label(search_frame, text="Search Barcode:", style='Dialog.TLabel').pack(side='left', padx=(0, 5))
    search_var = tk.StringVar()
    search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30, style='Dialog.TEntry')
    search_entry.pack(side='left', padx=5)
    search_entry.focus_set()
    search_context_menu = CustomContextMenu(search_entry)
    search_entry.bind("<Button-3>", search_context_menu.show)

    search_results_data = [] # To store results for export
    tree = None # Define tree here so clear_search_results can access it

    # --- Function Definitions within show_global_search ---
    def perform_search(event=None):
        # ... (perform_search code remains the same as before) ...
        nonlocal search_results_data, tree # Make tree accessible
        term = search_var.get().strip()
        if not term:
            messagebox.showwarning("Search", "Please enter a barcode (or part of it) to search.", parent=search_window)
            return

        for i in tree.get_children(): tree.delete(i)
        search_results_data = []
        export_button['state'] = 'disabled'
        status_label.config(text="Searching...")
        search_window.update_idletasks()

        conn = None
        try:
            conn = get_database_connection()
            if not conn: return

            cursor = conn.cursor()
            cursor.execute('''
                SELECT
                    s.session_name, r.barcode, r.nil_result, r.tb1_result, r.tb2_result, r.mit_result,
                    r.tb1_nil, r.tb2_nil, r.mit_nil, r.qft_result, r.requested_date
                FROM results r
                JOIN sessions s ON r.session_id = s.session_id
                WHERE r.barcode LIKE ?
                ORDER BY s.last_modified DESC, r.requested_date DESC
            ''', (f'%{term}%',))
            results = cursor.fetchall()

            if not results:
                status_label.config(text=f"No results found for '{term}'.")
                return

            decimals = app_settings['decimal_places']
            temp_data_list = []
            for row in results:
                (s_name, bc, nil, tb1, tb2, mit, t1n, t2n, mitn, qft, req_dt_str) = row
                row_dict = {
                    'session_name': s_name, 'barcode': bc, 'nil_result': nil, 'tb1_result': tb1,
                    'tb2_result': tb2, 'mit_result': mit, 'tb1_nil': t1n, 'tb2_nil': t2n,
                    'mit_nil': mitn, 'qft_result': qft
                }
                comment = calculate_comment(row_dict)
                row_dict['comment'] = comment
                row_dict['requested_date_str'] = req_dt_str
                try:
                    row_dict['requested_date_obj'] = datetime.datetime.strptime(req_dt_str, '%Y-%m-%d %H:%M:%S') if req_dt_str else None
                except (ValueError, TypeError):
                    row_dict['requested_date_obj'] = None
                temp_data_list.append(row_dict)

            search_results_data = temp_data_list
            for index, row_dict in enumerate(search_results_data):
                 display_row = (
                     row_dict['session_name'], row_dict['barcode'],
                     format_number_with_decimals(row_dict['nil_result'], decimals),
                     format_number_with_decimals(row_dict['tb1_result'], decimals),
                     format_number_with_decimals(row_dict['tb2_result'], decimals),
                     format_number_with_decimals(row_dict['mit_result'], decimals),
                     format_number_with_decimals(row_dict['tb1_nil'], decimals),
                     format_number_with_decimals(row_dict['tb2_nil'], decimals),
                     format_number_with_decimals(row_dict['mit_nil'], decimals),
                     row_dict['qft_result'], row_dict['comment'], row_dict['requested_date_str']
                 )
                 tree.insert('', 'end', values=display_row, tags=(str(index),))

            status_label.config(text=f"{len(results)} result(s) found for '{term}'.")
            if results:
                 export_button['state'] = 'normal' # Ensure button state is controlled here
        except sqlite3.Error as db_err:
            status_label.config(text="Database error occurred.")
            messagebox.showerror("Database Error", f"Search failed: {db_err}", parent=search_window)
        except Exception as e:
            status_label.config(text="An unexpected error occurred.")
            messagebox.showerror("Search Error", f"Search failed: {e}", parent=search_window)
            traceback.print_exc()
        finally:
            if conn: conn.close()

    # *** ADDED: Function to clear search ***
    def clear_search_results():
        nonlocal search_results_data, tree, export_button # Need to modify these
        search_var.set("") # Clear search box
        for i in tree.get_children(): # Clear treeview
            tree.delete(i)
        search_results_data = [] # Clear data list
        status_label.config(text="Enter barcode and press Search.") # Reset status
        export_button['state'] = 'disabled' # Disable export button
        search_entry.focus_set() # Set focus back to entry


    search_button = ttk.Button(search_frame, text="Search", style='Dialog.TButton', command=perform_search)
    search_button.pack(side='left', padx=5)
    search_entry.bind('<Return>', perform_search)

    # *** ADDED: Clear Button ***
    clear_button_search = ttk.Button(search_frame, text="Clear", style='DialogAlt.TButton', command=clear_search_results)
    clear_button_search.pack(side='left', padx=(0, 5)) # Place it next to search


    # --- Results Area ---
    results_frame = ttk.Frame(search_window, style='Dialog.TFrame')
    results_frame.pack(pady=5, padx=10, fill='both', expand=True)

    cols = ('session', 'barcode', 'nil', 'tb1', 'tb2', 'mit', 'tb1n', 'tb2n', 'mitn', 'qft', 'comment', 'req_date')
    # Assign tree to the variable defined earlier
    tree = ttk.Treeview(results_frame, columns=cols, show='headings', selectmode='extended')

    # Ctrl+A binding...
    def select_all_items(event=None):
        tree.selection_add(tree.get_children())
        return "break"
    tree.bind('<Control-a>', select_all_items)
    tree.bind('<Control-A>', select_all_items)
    tree.bind('<Command-a>', select_all_items)
    tree.bind('<Command-A>', select_all_items)

    # Column setup...
    col_widths = {'session': 180, 'barcode': 100, 'nil': 50, 'tb1': 50, 'tb2': 50, 'mit': 50,
                  'tb1n': 50, 'tb2n': 50, 'mitn': 50, 'qft': 60, 'comment': 80, 'req_date': 140}
    col_headers = {'tb1n':'TB1-Nil','tb2n':'TB2-Nil','mitn':'Mit-Nil','req_date':'Requested Date'}
    col_align = {c: 'center' for c in cols if c not in ['session', 'comment', 'req_date']}
    col_align['session'] = 'w'; col_align['comment'] = 'w'; col_align['req_date'] = 'center'

    for col in cols:
        tree.heading(col, text=col_headers.get(col, col.replace('_', ' ').title()))
        tree.column(col, width=col_widths.get(col, 80), anchor=col_align.get(col, 'w'), stretch=tk.NO)

    # Scrollbars...
    vsb = ttk.Scrollbar(results_frame, orient="vertical", command=tree.yview)
    vsb.pack(side='right', fill='y')
    hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=tree.xview)
    hsb.pack(side='bottom', fill='x')
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.pack(side='left', fill='both', expand=True)

    # --- Status and Export ---
    status_export_frame = ttk.Frame(search_window, style='Dialog.TFrame')
    status_export_frame.pack(pady=10, padx=10, fill='x')

    status_label = ttk.Label(status_export_frame, text="Enter barcode and press Search.", style='DialogStatus.TLabel')
    status_label.pack(side='left', fill='x', expand=True)

    # --- Export Dialog Definition (remains the same) ---
    def export_search_results_dialog():
        # ... (code for the styled export dialog as provided previously) ...
        nonlocal search_results_data, tree # Need tree access
        if not search_results_data:
             messagebox.showinfo("Export", "No search results to export.", parent=search_window)
             return

        selected_item_ids = tree.selection()
        num_selected = len(selected_item_ids)
        num_total = len(search_results_data)

        export_title = "Export Search Results"
        export_message = f"Choose export format"
        if num_selected > 0:
             export_message += f" for {num_selected} selected item(s):"
        else:
             export_message += f" for all {num_total} displayed items:"

        export_choice_win = tk.Toplevel(search_window)
        export_choice_win.title(export_title)
        dialog_width = 350
        dialog_height = 410
        export_choice_win.geometry(f"{dialog_width}x{dialog_height}")
        export_choice_win.resizable(False, False)
        export_choice_win.transient(search_window)
        export_choice_win.grab_set()
        export_choice_win.configure(bg='#f0f2f5')
        main_app.center_window(export_choice_win, dialog_width, dialog_height)

        main_frame_export = ttk.Frame(export_choice_win, style='TFrame', padding=(20, 15))
        main_frame_export.pack(fill=tk.BOTH, expand=True)

        try:
             main_app.style.configure('DialogHeader.TLabel', font=('Open Sans', 16, 'bold'), foreground='#1976D2', background='#f0f2f5')
        except tk.TclError: pass
        title_label_export = ttk.Label(main_frame_export, text=export_title, style='DialogHeader.TLabel', anchor='center')
        title_label_export.pack(pady=(0, 5))

        subtitle_label_export = ttk.Label(main_frame_export, text=export_message, style='Subtitle.TLabel', anchor='center', wraplength=dialog_width-40)
        subtitle_label_export.pack(pady=(0, 20))

        options_frame_export = ttk.Frame(main_frame_export, style='TFrame')
        options_frame_export.pack(fill=tk.X, expand=True)

        def do_export(fmt):
            data_to_export_list = []
            if selected_item_ids:
                try:
                    for item_id in selected_item_ids:
                        original_index = int(tree.item(item_id, 'tags')[0])
                        data_to_export_list.append(search_results_data[original_index])
                except (IndexError, ValueError) as e:
                    messagebox.showerror("Export Error", f"Error retrieving selected data: {e}", parent=export_choice_win)
                    export_choice_win.destroy()
                    return
            else:
                data_to_export_list = search_results_data

            if not data_to_export_list:
                 messagebox.showwarning("Export", "No data selected or available for export.", parent=export_choice_win)
                 export_choice_win.destroy()
                 return

            export_choice_win.destroy()
            if fmt == 'excel': export_global_search_to_excel(data_to_export_list)
            elif fmt == 'csv': export_global_search_to_csv(data_to_export_list)
            elif fmt == 'pdf': export_global_search_to_pdf(data_to_export_list)

        export_options_list = [
            {'text': "📄 PDF Report", 'desc': "Formatted PDF document", 'command': lambda: do_export('pdf')},
            {'text': "📊 Excel File", 'desc': "Spreadsheet with formatting", 'command': lambda: do_export('excel')},
            {'text': "📝 CSV File",   'desc': "Plain text, comma-separated", 'command': lambda: do_export('csv')}
        ]

        try:
            main_app.style.configure('DialogDesc.TLabel', font=('Open Sans', 9), foreground='#666666', background='#f0f2f5')
        except tk.TclError: pass

        for option in export_options_list:
            option_sub_frame = ttk.Frame(options_frame_export, style='TFrame')
            option_sub_frame.pack(pady=7, fill=tk.X)
            btn = ttk.Button(option_sub_frame, text=option['text'], command=option['command'], style='Custom.TButton', width=20)
            btn.pack()
            desc_label = ttk.Label(option_sub_frame, text=option['desc'], style='DialogDesc.TLabel', anchor='center')
            desc_label.pack(pady=(2, 0))

        ttk.Separator(main_frame_export, orient='horizontal').pack(fill='x', pady=(15, 10))
        cancel_button = ttk.Button(main_frame_export, text="Cancel", command=export_choice_win.destroy, style='Alt.TButton', width=20)
        cancel_button.pack(pady=(5, 10))

    # --- Define and Place the SINGLE Export Button ---
    # Ensure this is the ONLY place the export button for the status_export_frame is defined and packed.
    export_button = ttk.Button(status_export_frame, text="Export Results...", style='Dialog.TButton', state='disabled',
                               command=export_search_results_dialog)
    export_button.pack(side='right') # Packed on the right side of the status frame
def export_global_search_to_pdf(results_list):
    """Exports global search results list (of dicts) to PDF."""
    if not results_list: return

    default_filename = f"QFT_GlobalSearch_{datetime.datetime.now().strftime('%Y%m%d')}.pdf"
    filename = filedialog.asksaveasfilename(
        defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
        title="Save Global Search Results As PDF", initialfile=default_filename)
    if not filename: return

    main_app.update_status("Exporting search results to PDF...", show_progress=True)

    try:
        doc = SimpleDocTemplate(filename, pagesize=landscape(letter),
                                rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        styles = getSampleStyleSheet()
        story = []

        # --- PDF Styles (Similar to main PDF export, maybe smaller font) ---
        current_colors = app_settings
        title_style = ParagraphStyle('SearchTitle', parent=styles['h1'], fontSize=18, alignment=1, spaceAfter=12)
        date_style = ParagraphStyle('SearchDate', parent=styles['Normal'], fontSize=10, alignment=1, spaceAfter=12)
        header_style = ParagraphStyle('SearchHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, alignment=1)
        cell_style = ParagraphStyle('SearchCell', parent=styles['Normal'], fontSize=8, alignment=1)
        comment_style = ParagraphStyle('SearchComment', parent=cell_style, textColor=reportlab_colors.grey)
        # Styles with colors for QFT cells
        qft_pos_style = ParagraphStyle('SearchQFT_POS', parent=cell_style, textColor=hex_to_color(current_colors['pos_text']))
        qft_wp_style = ParagraphStyle('SearchQFT_WP', parent=cell_style, textColor=hex_to_color(current_colors['wp_text']))
        qft_neg_style = ParagraphStyle('SearchQFT_NEG', parent=cell_style, textColor=hex_to_color(current_colors['neg_text']))
        qft_ind_style = ParagraphStyle('SearchQFT_IND', parent=cell_style, textColor=hex_to_color(current_colors['ind_text']))


        # --- PDF Header ---
        story.append(Paragraph("Global Search Results", title_style))
        story.append(Paragraph(f"Exported: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", date_style))
        story.append(Spacer(1, 15))

        # --- PDF Data Table ---
        # Define headers including Session Name
        headers = ['Session', 'Barcode', 'Nil', 'TB1', 'TB2', 'Mit', 'TB1-Nil', 'TB2-Nil', 'Mit-Nil', 'QFT Result', 'Comment', 'Req. Date']
        table_data = [[Paragraph(h, header_style) for h in headers]]
        decimals = app_settings['decimal_places']

        for row_dict in results_list:
            # Determine QFT style
            qft_result = str(row_dict.get('qft_result', ' ')).upper()
            comment = row_dict.get('comment', '') # Comment already calculated
            is_wp = "WP" in comment

            current_qft_style = cell_style
            if is_wp: current_qft_style = qft_wp_style
            elif qft_result in ('POS', 'POS*'): current_qft_style = qft_pos_style
            elif qft_result == 'NEG': current_qft_style = qft_neg_style
            elif qft_result == 'IND': current_qft_style = qft_ind_style

            row_values = [
                Paragraph(str(row_dict.get('session_name', '')), cell_style),
                Paragraph(str(row_dict.get('barcode', '')), cell_style),
                Paragraph(format_number_with_decimals(row_dict.get('nil_result', ''), decimals), cell_style),
                Paragraph(format_number_with_decimals(row_dict.get('tb1_result', ''), decimals), cell_style),
                Paragraph(format_number_with_decimals(row_dict.get('tb2_result', ''), decimals), cell_style),
                Paragraph(format_number_with_decimals(row_dict.get('mit_result', ''), decimals), cell_style),
                Paragraph(format_number_with_decimals(row_dict.get('tb1_nil', ''), decimals), cell_style),
                Paragraph(format_number_with_decimals(row_dict.get('tb2_nil', ''), decimals), cell_style),
                Paragraph(format_number_with_decimals(row_dict.get('mit_nil', ''), decimals), cell_style),
                Paragraph(qft_result, current_qft_style), # Use colored style
                Paragraph(comment, comment_style),
                Paragraph(str(row_dict.get('requested_date_str', '')), cell_style), # Use date string
            ]
            table_data.append(row_values)

        # Define column widths (adjust as needed)
        col_widths = [100, 70, 45, 45, 45, 45, 50, 50, 50, 60, 65, 80] # Added session, adjusted req date

        data_table = Table(table_data, colWidths=col_widths, repeatRows=1)

        # --- PDF Table Styling ---
        table_style_commands = [
            ('GRID', (0, 0), (-1, -1), 0.5, reportlab_colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('BACKGROUND', (0, 0), (-1, 0), reportlab_colors.lightgrey), # Header BG
            ('TEXTCOLOR', (0, 0), (-1, 0), reportlab_colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            # Add specific header background colors if desired (like main export)
            ('BACKGROUND', (2, 0), (2, 0), reportlab_colors.Color(0.85, 0.85, 0.85)), # Nil
            ('BACKGROUND', (3, 0), (3, 0), reportlab_colors.Color(0.7, 0.9, 0.7)),   # TB1
            ('BACKGROUND', (4, 0), (4, 0), reportlab_colors.Color(1.0, 0.95, 0.7)),  # TB2
            ('BACKGROUND', (5, 0), (5, 0), reportlab_colors.Color(0.85, 0.7, 0.9)),  # Mit
            ('BACKGROUND', (9, 0), (9, 0), reportlab_colors.Color(0.529, 0.808, 0.922)),# QFT
        ]

        # Apply row-specific background colors (text color handled by Paragraph)
        for r_idx, row_dict in enumerate(results_list, 1):
            qft_result = str(row_dict.get('qft_result', '')).upper()
            comment = row_dict.get('comment', '')
            is_wp = "WP" in comment
            row_bg_color = None

            if is_wp: row_bg_color = hex_to_color(current_colors['wp_bg'])
            elif qft_result in ('POS', 'POS*'): row_bg_color = hex_to_color(current_colors['pos_bg'])
            elif qft_result == 'NEG': row_bg_color = hex_to_color(current_colors['neg_bg'])
            elif qft_result == 'IND': row_bg_color = hex_to_color(current_colors['ind_bg'])

            if row_bg_color:
                table_style_commands.append(('BACKGROUND', (0, r_idx), (-1, r_idx), row_bg_color))

        data_table.setStyle(TableStyle(table_style_commands))
        story.append(data_table)

        # --- Build PDF ---
        doc.build(story)
        main_app.update_status("Search results PDF export complete.", hide_progress=True)
        messagebox.showinfo("Export PDF", "Search results PDF exported successfully!", parent=main_app.master) # Specify parent

        # Ask to open file
        if messagebox.askyesno("Open File", "Do you want to open the exported PDF file?", parent=main_app.master):
            try:
                if platform.system() == 'Darwin': os.system(f'open "{filename}"')
                elif platform.system() == 'Windows': os.startfile(filename)
                else: os.system(f'xdg-open "{filename}"')
            except Exception as open_err:
                 messagebox.showwarning("Open File Error", f"Could not automatically open the file:\n{open_err}", parent=main_app.master)

    except Exception as e:
        main_app.update_status("Search results PDF export failed.", hide_progress=True)
        messagebox.showerror("Export PDF Error", f"Failed to export search results PDF: {str(e)}", parent=main_app.master)
        traceback.print_exc()

def export_global_search_to_excel(results_list):
    """Exports global search results list (of dicts) to Excel (Keep Script 1)."""
    if not results_list: return

    default_filename = f"QFT_GlobalSearch_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
    filename = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
        title="Save Global Search Results As", initialfile=default_filename)
    if not filename: return

    main_app.update_status("Exporting search results to Excel...", show_progress=True)
    try:
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet("Search Results")

        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'border': 1})
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'align': 'center', 'valign': 'vcenter', 'border': 1}) # Match DB format
        comment_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'italic': True, 'font_color': '#555555'})


        headers = ['Session', 'Barcode', 'Nil', 'TB1', 'TB2', 'Mit', 'TB1-Nil', 'TB2-Nil', 'Mit-Nil', 'QFT Result', 'Comment', 'Requested Date']
        worksheet.write_row(0, 0, headers, header_format)

        decimals = app_settings['decimal_places']
        for r_idx, row_dict in enumerate(results_list, 1):
            col = 0
            worksheet.write(r_idx, col, row_dict.get('session_name',''), cell_format); col+=1
            worksheet.write(r_idx, col, row_dict.get('barcode',''), cell_format); col+=1

            num_keys = ['nil_result', 'tb1_result', 'tb2_result', 'mit_result', 'tb1_nil', 'tb2_nil', 'mit_nil']
            for key in num_keys:
                val_str = format_number_with_decimals(row_dict.get(key, ''), decimals)
                # Try number conversion
                try:
                    if '>' in val_str or '<' in val_str or val_str.strip() == "":
                         worksheet.write_string(r_idx, col, val_str, cell_format)
                    else:
                         worksheet.write_number(r_idx, col, float(val_str), cell_format)
                except ValueError:
                     worksheet.write_string(r_idx, col, val_str, cell_format)
                col += 1

            worksheet.write(r_idx, col, row_dict.get('qft_result',''), cell_format); col+=1
            worksheet.write(r_idx, col, row_dict.get('comment',''), comment_format); col+=1
            # Write date using the parsed object if available, otherwise the string
            req_date_obj = row_dict.get('requested_date_obj')
            if isinstance(req_date_obj, datetime.datetime):
                worksheet.write_datetime(r_idx, col, req_date_obj, date_format)
            else:
                worksheet.write(r_idx, col, row_dict.get('requested_date_str',''), cell_format) # Use the original string
            col+=1

        # Adjust widths (adjust as needed)
        worksheet.set_column('A:A', 25) # Session
        worksheet.set_column('B:B', 12) # Barcode
        worksheet.set_column('C:I', 10) # Numeric
        worksheet.set_column('J:J', 12) # QFT Result
        worksheet.set_column('K:K', 15) # Comment
        worksheet.set_column('L:L', 18) # Date
        worksheet.freeze_panes(1, 0)

        workbook.close()
        main_app.update_status("Search results exported.", hide_progress=True)
        messagebox.showinfo("Export Successful", "Search results exported to Excel.", parent=main_app.master)
        # Offer to open
        if messagebox.askyesno("Open File", "Open the exported Excel file?", parent=main_app.master):
             try:
                 if platform.system() == 'Darwin': os.system(f'open "{filename}"')
                 elif platform.system() == 'Windows': os.startfile(filename)
                 else: os.system(f'xdg-open "{filename}"')
             except Exception as open_err:
                  messagebox.showwarning("Open File Error", f"Could not automatically open the file:\n{open_err}", parent=main_app.master)

    except Exception as e:
        main_app.update_status("Export failed.", hide_progress=True)
        messagebox.showerror("Export Error", f"Failed to export search results to Excel: {e}", parent=main_app.master)
        traceback.print_exc()

def export_global_search_to_csv(results_list):
    """Exports global search results list (of dicts) to CSV (Keep Script 1)."""
    if not results_list: return

    default_filename = f"QFT_GlobalSearch_{datetime.datetime.now().strftime('%Y%m%d')}.csv"
    filename = filedialog.asksaveasfilename(
        defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
        title="Save Global Search Results As", initialfile=default_filename)
    if not filename: return

    main_app.update_status("Exporting search results to CSV...", show_progress=True)
    try:
        headers = ['Session', 'Barcode', 'Nil_Result', 'TB1_Result', 'TB2_Result', 'Mit_Result',
                   'TB1_Nil', 'TB2_Nil', 'Mit_Nil', 'QFT_Result', 'Comment', 'Requested_Date']
        decimals = app_settings['decimal_places']

        # Use semicolon for consistency with main export if desired, or comma
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile, delimiter=',') # Use comma like main CSV export
            writer.writerow(headers)
            for row_dict in results_list:
                 row_values = [
                     row_dict.get('session_name',''), row_dict.get('barcode',''),
                     format_number_with_decimals(row_dict.get('nil_result', ''), decimals),
                     format_number_with_decimals(row_dict.get('tb1_result', ''), decimals),
                     format_number_with_decimals(row_dict.get('tb2_result', ''), decimals),
                     format_number_with_decimals(row_dict.get('mit_result', ''), decimals),
                     format_number_with_decimals(row_dict.get('tb1_nil', ''), decimals),
                     format_number_with_decimals(row_dict.get('tb2_nil', ''), decimals),
                     format_number_with_decimals(row_dict.get('mit_nil', ''), decimals),
                     row_dict.get('qft_result',''), row_dict.get('comment',''),
                     row_dict.get('requested_date_str','') # Use the string date from DB
                 ]
                 writer.writerow(row_values)

        main_app.update_status("Search results exported.", hide_progress=True)
        messagebox.showinfo("Export Successful", "Search results exported to CSV.", parent=main_app.master)
        if messagebox.askyesno("Open File", "Open the exported CSV file?", parent=main_app.master):
             try:
                 if platform.system() == 'Darwin': os.system(f'open "{filename}"')
                 elif platform.system() == 'Windows': os.startfile(filename)
                 else: os.system(f'xdg-open "{filename}"')
             except Exception as open_err:
                  messagebox.showwarning("Open File Error", f"Could not automatically open the file:\n{open_err}", parent=main_app.master)

    except Exception as e:
        main_app.update_status("Export failed.", hide_progress=True)
        messagebox.showerror("Export Error", f"Failed to export search results to CSV: {e}", parent=main_app.master)
        traceback.print_exc()

# --- Manual Order (Keep Script 1's implementation) ---
def show_manual_order_dialog():
    """Opens a dialog to manually reorder the currently displayed data (Keep Script 1)."""
    global main_app
    print("--- Attempting to open Manual Order dialog ---")
    if not main_app.has_data():
        print("Manual Order: No data found.")
        messagebox.showinfo("Manual Order", "No data loaded to reorder.", parent=main_app.master)
        return
    print("Manual Order: Data found, creating window...")

    try:
        order_window = tk.Toplevel(main_app.master)
        order_window.title("Manually Reorder Results")
        order_window.geometry("700x600")
        order_window.transient(main_app.master)
        order_window.grab_set()
        order_window.configure(bg=app_settings.get('dialog_bg', '#F5F5F5'))
        main_app.center_window(order_window, 700, 600)
        print("Manual Order: Window created successfully.")
    except Exception as e:
         print(f"!!! ERROR creating Manual Order window: {e}")
         traceback.print_exc()
         messagebox.showerror("Error", f"Could not open Manual Order window:\n{e}", parent=main_app.master)
         return # Exit if window fails


    # Frame for listbox and scrollbar
    list_frame = ttk.Frame(order_window, style='Dialog.TFrame')
    list_frame.pack(padx=10, pady=(10, 5), fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(list_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    order_listbox = tk.Listbox(
        list_frame,
        selectmode=tk.SINGLE,
        yscrollcommand=scrollbar.set,
        font=('Consolas', 10), # Monospaced font for alignment
        activestyle='dotbox', # Visual feedback on selection
        exportselection=False # Prevent selection loss on focus change
    )
    order_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=order_listbox.yview)

    # Populate listbox with current data (using internal data store)
    current_ordered_data = main_app.current_data # Get the list of dicts

    # Use a stable unique identifier if possible (barcode assumed unique here)
    listbox_data_map = {} # Map listbox index to original data dictionary

    for i, row_dict in enumerate(current_ordered_data):
        bc = str(row_dict.get('barcode', 'N/A'))
        qft = str(row_dict.get('qft_result', ''))
        req_dt = row_dict.get('requested_date')
        dt_str = req_dt.strftime('%Y-%m-%d') if isinstance(req_dt, datetime.datetime) else 'No Date'
        # Pad barcode for better alignment
        display_text = f"{bc:<15} | {qft:<5} | {dt_str}"
        order_listbox.insert(tk.END, display_text)
        # Store the actual dictionary with the listbox index
        listbox_data_map[i] = row_dict

    def move_item(direction):
        selected_idx_tuple = order_listbox.curselection()
        if not selected_idx_tuple: return
        current_lb_idx = selected_idx_tuple[0]

        if direction == 'up' and current_lb_idx > 0:
            target_lb_idx = current_lb_idx - 1
        elif direction == 'down' and current_lb_idx < order_listbox.size() - 1:
            target_lb_idx = current_lb_idx + 1
        else:
            return # Cannot move further

        # Swap listbox items (text)
        current_text = order_listbox.get(current_lb_idx)
        target_text = order_listbox.get(target_lb_idx)
        order_listbox.delete(current_lb_idx)
        order_listbox.insert(target_lb_idx, current_text)

        # Swap the associated data dictionaries in our map
        temp_data = listbox_data_map[current_lb_idx]
        listbox_data_map[current_lb_idx] = listbox_data_map[target_lb_idx]
        listbox_data_map[target_lb_idx] = temp_data

        # Update selection and view
        order_listbox.selection_clear(0, tk.END)
        order_listbox.selection_set(target_lb_idx)
        order_listbox.activate(target_lb_idx)
        order_listbox.see(target_lb_idx) # Ensure visible


    def apply_manual_order():
            print("\n--- Inside apply_manual_order ---")
            try:
                # Get the final order of data dictionaries from the map
                new_data_order = [listbox_data_map[i] for i in range(order_listbox.size())]

                if len(new_data_order) != len(main_app.current_data):
                    print(f"!!! ERROR: Length mismatch! Original: {len(main_app.current_data)}, New: {len(new_data_order)}")
                    messagebox.showerror("Order Error", "Failed to apply order due to item mismatch.", parent=order_window)
                    return

                print("Applying new order to main_app.current_data")
                main_app.current_data = new_data_order # Update the main data store

                # Update sorting state to reflect manual order
                main_app.current_sort_column = None
                main_app.current_sort_direction = None
                main_app.sort_dropdown.set("Manual Order") # Update dropdown display
                print("Sort dropdown set to 'Manual Order'.")

                print("Calling main_app.refresh_display()")
                main_app.refresh_display() # Refresh the display with new order
                print("refresh_display completed.")

                order_window.destroy()
                main_app.update_status("Manual order applied.")
                print("--- apply_manual_order finished successfully ---")

            except Exception as e:
                print(f"!!! ERROR in apply_manual_order: {e}")
                traceback.print_exc()
                messagebox.showerror("Error Applying Order", f"An unexpected error occurred: {e}", parent=order_window)

    # Buttons (Keep Script 1's)
    button_frame = ttk.Frame(order_window, style='Dialog.TFrame')
    button_frame.pack(pady=5, padx=10, fill='x')

    up_button = ttk.Button(button_frame, text="Move Up ↑", style='Dialog.TButton', command=lambda: move_item('up'))
    up_button.pack(side=tk.LEFT, padx=5)

    down_button = ttk.Button(button_frame, text="Move Down ↓", style='Dialog.TButton', command=lambda: move_item('down'))
    down_button.pack(side=tk.LEFT, padx=5)

    apply_button = ttk.Button(button_frame, text="Apply This Order", style='DialogHighlight.TButton', command=apply_manual_order)
    apply_button.pack(side=tk.RIGHT, padx=5)

    cancel_button = ttk.Button(button_frame, text="Cancel", style='DialogAlt.TButton', command=order_window.destroy)
    cancel_button.pack(side=tk.RIGHT, padx=5)

    # Bind arrow keys
    order_window.bind('<Up>', lambda e: move_item('up'))
    order_window.bind('<Down>', lambda e: move_item('down'))


# --- Settings/Color Customization (Keep Script 1's implementation) ---
def customize_appearance():
    """Opens the appearance customization window (Adds WP)."""
    global main_app, app_settings

    custom_window = tk.Toplevel(main_app.master)
    custom_window.title("Customize Appearance")
    custom_window.geometry("450x430")
    custom_window.transient(main_app.master)
    custom_window.grab_set()
    custom_window.configure(bg=app_settings.get('dialog_bg', '#F5F5F5'))
    main_app.center_window(custom_window, 450, 430)

    # --- Variables (remain the same) ---
    pos_bg_var = tk.StringVar(value=app_settings.get('pos_bg','#FFFFE0'))
    neg_bg_var = tk.StringVar(value=app_settings.get('neg_bg','#FFFFFF'))
    ind_bg_var = tk.StringVar(value=app_settings.get('ind_bg','#FFFFFF'))
    wp_bg_var = tk.StringVar(value=app_settings.get('wp_bg','#FFF8DC'))
    pos_text_var = tk.StringVar(value=app_settings.get('pos_text','#e53935'))
    neg_text_var = tk.StringVar(value=app_settings.get('neg_text','#43a047'))
    ind_text_var = tk.StringVar(value=app_settings.get('ind_text','#fb8c00'))
    wp_text_var = tk.StringVar(value=app_settings.get('wp_text','#D2691E'))
    decimal_var = tk.StringVar(value=str(app_settings.get('decimal_places','default')))

    # --- Frames (remain the same) ---
    main_options_frame = ttk.Frame(custom_window, style='Dialog.TFrame')
    main_options_frame.pack(pady=10, padx=15, fill='both', expand=True)
    ttk.Label(main_options_frame, text="Result Highlighting", style='DialogSubtitle.TLabel').pack(anchor='w', pady=(0, 5))
    preview_widgets = {}

    # --- Helper to create color pickers (MODIFIED LAYOUT) ---
    def create_color_picker(parent, text, bg_var, text_var, preview_id):
        # Main frame for this row
        row_frame = ttk.Frame(parent, style='Dialog.TFrame', padding=(0, 2)) # Add vertical padding between rows
        row_frame.pack(fill='x')

        # Configure columns for Label, BG controls, Text controls
        row_frame.grid_columnconfigure(0, weight=1, minsize=130) # Label column - allow expanding, give min size
        row_frame.grid_columnconfigure(1, weight=0) # BG Controls column
        row_frame.grid_columnconfigure(2, weight=0) # Text Controls column

        # --- Label ---
        label = ttk.Label(row_frame, text=f"{text}:", style='Dialog.TLabel')
        label.grid(row=0, column=0, sticky='w', padx=(0, 15)) # Pad right of label

        # --- BG Controls Sub-Frame ---
        bg_controls_frame = ttk.Frame(row_frame, style='Dialog.TFrame')
        bg_controls_frame.grid(row=0, column=1, sticky='w') # Place in column 1

        bg_preview = tk.Label(bg_controls_frame, width=2, height=1, bg=bg_var.get(), relief='solid', borderwidth=1)
        bg_preview.pack(side='left', padx=(0, 2)) # Pack tightly left

        bg_button = ttk.Button(bg_controls_frame, text="BG Color", width=8, style='SmallDialog.TButton',
                              command=lambda v=bg_var, p=bg_preview: choose_color(v, p))
        bg_button.pack(side='left') # Pack next to preview

        # --- Text Controls Sub-Frame ---
        text_controls_frame = ttk.Frame(row_frame, style='Dialog.TFrame')
        text_controls_frame.grid(row=0, column=2, sticky='w', padx=(10, 0)) # Place in column 2, pad left

        text_preview = tk.Label(text_controls_frame, text="Text", width=4, height=1, fg=text_var.get(), bg='white', relief='solid', borderwidth=1)
        text_preview.pack(side='left', padx=(0, 2)) # Pack tightly left

        text_button = ttk.Button(text_controls_frame, text="Text Color", width=9, style='SmallDialog.TButton',
                                command=lambda v=text_var, p=text_preview: choose_color(v, p, is_text=True))
        text_button.pack(side='left') # Pack next to preview

        # Store preview labels for live update (no change)
        preview_widgets[preview_id] = {'bg': bg_preview, 'text': text_preview}
        bg_var.trace_add('write', lambda name, index, mode, p=bg_preview, v=bg_var: p.config(bg=v.get()))
        text_var.trace_add('write', lambda name, index, mode, p=text_preview, v=text_var: p.config(fg=v.get()))

    def choose_color(color_var, preview_label, is_text=False):
        # ... (choose_color function remains the same) ...
        title = "Choose Text Color" if is_text else "Choose Background Color"
        initial_color = color_var.get()
        result = colorchooser.askcolor(color=initial_color, title=title, parent=custom_window)
        new_color_hex = result[1] # Get the hex string
        if new_color_hex: # If a color was chosen
            color_var.set(new_color_hex)

    # --- Create pickers including WP (call remains the same) ---
    create_color_picker(main_options_frame, "Positive (POS)", pos_bg_var, pos_text_var, 'pos')
    create_color_picker(main_options_frame, "Weak Positive (WP)", wp_bg_var, wp_text_var, 'wp')
    create_color_picker(main_options_frame, "Negative (NEG)", neg_bg_var, neg_text_var, 'neg')
    create_color_picker(main_options_frame, "Indeterminate (IND)", ind_bg_var, ind_text_var, 'ind')

    # --- Separator, Decimal Places, Action Buttons (remain the same) ---
    ttk.Separator(main_options_frame, orient='horizontal').pack(fill='x', pady=10)
    ttk.Label(main_options_frame, text="Numeric Display", style='DialogSubtitle.TLabel').pack(anchor='w', pady=(5, 5))
    decimal_frame = ttk.Frame(main_options_frame, style='Dialog.TFrame', padding=5)
    decimal_frame.pack(fill='x', pady=2)
    ttk.Label(decimal_frame, text="Decimal Places:", width=15, style='Dialog.TLabel').pack(side='left', padx=(0,10))
    decimal_combo = ttk.Combobox(decimal_frame, textvariable=decimal_var,
                                 values=['default', '0', '1', '2', '3'],
                                 state='readonly', width=10, style='Dialog.TCombobox')
    decimal_combo.pack(side='left')

    button_frame = ttk.Frame(custom_window, style='Dialog.TFrame')
    button_frame.pack(side='bottom', pady=10, padx=15, fill='x')
    # ... (apply_settings, reset_to_defaults functions remain the same) ...
    def apply_settings():
        global app_settings
        new_settings = {
            'pos_bg': pos_bg_var.get(), 'neg_bg': neg_bg_var.get(), 'ind_bg': ind_bg_var.get(), 'wp_bg': wp_bg_var.get(),
            'pos_text': pos_text_var.get(), 'neg_text': neg_text_var.get(), 'ind_text': ind_text_var.get(), 'wp_text': wp_text_var.get(),
            'decimal_places': decimal_var.get()
        }
        for key, val in new_settings.items():
             if ('bg' in key or 'text' in key) and not (val.startswith('#') and len(val) == 7):
                  try:
                       custom_window.winfo_rgb(val)
                  except tk.TclError:
                       messagebox.showerror("Invalid Color", f"Invalid color format for {key}: '{val}'.\nPlease use #RRGGBB format.", parent=custom_window)
                       return
        app_settings = new_settings
        save_settings(app_settings)
        main_app.apply_styles()
        main_app.refresh_display()
        custom_window.destroy()
        main_app.update_status("Appearance settings applied.")

    def reset_to_defaults():
        defaults = {
            'pos_bg': '#FFFFE0', 'neg_bg': '#FFFFFF', 'ind_bg': '#FFFFFF', 'wp_bg': '#FFF8DC',
            'pos_text': '#e53935', 'neg_text': '#43a047', 'ind_text': '#fb8c00', 'wp_text': '#D2691E',
            'decimal_places': 'default'
        }
        pos_bg_var.set(defaults['pos_bg'])
        neg_bg_var.set(defaults['neg_bg'])
        ind_bg_var.set(defaults['ind_bg'])
        wp_bg_var.set(defaults['wp_bg'])
        pos_text_var.set(defaults['pos_text'])
        neg_text_var.set(defaults['neg_text'])
        ind_text_var.set(defaults['ind_text'])
        wp_text_var.set(defaults['wp_text'])
        decimal_var.set(defaults['decimal_places'])

    reset_button = ttk.Button(button_frame, text="Reset Defaults", style='DialogAlt.TButton', command=reset_to_defaults)
    reset_button.pack(side='left', padx=5)
    cancel_button = ttk.Button(button_frame, text="Cancel", style='DialogAlt.TButton', command=custom_window.destroy)
    cancel_button.pack(side='right', padx=5)
    apply_button = ttk.Button(button_frame, text="Apply Settings", style='DialogHighlight.TButton', command=apply_settings)
    apply_button.pack(side='right', padx=5)

# ... (Rest of the script remains the same) ...
# --- Main Application Class ---

class QFTApp:
    """Main application class (Using Script 1's structure, adapted GUI)."""
    def __init__(self, master):
        self.master = master
        self.master.title("QFT-Plus Data Viewer v2.1") # Updated version
        self.master.geometry("1000x700") # Initial size from S2
        self.master.minsize(1000, 700) # Min size from S2
        self.master.configure(background='#f0f2f5') # Background from S2

        # Define column structure (used for Text widget formatting) - Adjust widths to match S2 look
        # Widths below are approximate character counts for Consolas font
        self.headers_info = [
             ("Barcode", 'header_base', 12),    # S2: 12
             ("Nil_Result", 'header_nil', 12),  # S2: 12
             ("TB1_Result", 'header_tb1', 12),  # S2: 12
             ("TB2_Result", 'header_tb2', 12),  # S2: 12
             ("Mit_Result", 'header_mit', 12),  # S2: 12
             ("TB1_Nil", 'header_tb1', 12),     # S2: 12
             ("TB2_Nil", 'header_tb2', 12),     # S2: 12
             ("Mit_Nil", 'header_mit', 12),     # S2: 12
             ("QFT_Result", 'header_qft', 8),   # S2: 8
             ("Comment", 'header_comment', 15), # Added for comment display
             ("Request Date", 'header_base', 19) # Added for date display
         ]
        self.total_width = sum(w for _, _, w in self.headers_info) + len(self.headers_info) # Approx total width

        # --- Global App State ---
        self.current_data = [] # List of dictionaries, canonical data store
        self.imported_filename_source = "" # Track origin (filename or session)
        self.current_sort_column = 'requested_date' # Default sort (matches S2 initial sort)
        self.current_sort_direction = 'asc'         # Default sort (matches S2 initial sort)

        # Load settings early
        global app_settings
        app_settings = load_settings()

        # Try setting the app icon
        try:
            if APP_ICON_PATH and os.path.exists(APP_ICON_PATH):
                self.master.iconbitmap(APP_ICON_PATH)
        except Exception as icon_err:
            print(f"Could not set application icon: {icon_err}")

        self.setup_styles()
        self.create_menu()       # Creates the menu bar
        self.create_widgets()    # Creates the main GUI layout (adapted from S2)
        self.apply_styles()      # Apply loaded settings/styles to widgets

        # Center the window on first launch
        self.center_window(self.master, 1000, 700) # Use S2 dimensions

        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Initialize status bar
        self.update_status("Ready. Import data or load a session.")


    def center_window(self, window_instance, width=None, height=None):
        """Centers a window (main or Toplevel) on the screen."""
        window_instance.update_idletasks() # Ensure dimensions are calculated
        w = width or window_instance.winfo_width()
        h = height or window_instance.winfo_height()
        sw = window_instance.winfo_screenwidth()
        sh = window_instance.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        window_instance.geometry(f'{w}x{h}+{x}+{y}')

    def setup_styles(self):
        """Configure ttk styles (Based on S2 styles + S1 structure)."""
        self.style = ttk.Style()
        self.style.theme_use('clam') # Consistent theme

        # --- Color Palette (Use S2's general feel) ---
        primary_color = "#2196F3" # S2 Blue
        secondary_color = "#f0f2f5" # S2 Background
        text_color = "#333333"
        alt_button_bg = "#DCDCDC" # Light grey for less prominent buttons
        alt_button_fg = "#333333"
        dialog_bg = '#F5F5F5'
        status_fg = '#666666' # S2 Subtitle color

        # --- Base Styles ---
        self.style.configure('.', font=('Open Sans', 10), background=secondary_color, foreground=text_color) # S2 Font
        self.style.configure('TFrame', background=secondary_color)
        self.style.configure('TLabel', background=secondary_color, foreground=text_color, padding=5)
        # Button style from S2
        self.style.configure('Custom.TButton', font=('Open Sans', 11, 'bold'), padding=(25, 12),
                             background=primary_color, foreground='white')
        self.style.map('Custom.TButton',
                       background=[('active', '#1976D2'), ('disabled', '#BDBDBD')], # Darker blue on active
                       foreground=[('disabled', '#757575')])
        # Alt button style (for less prominent buttons like Clear/Cancel)
        self.style.configure('Alt.TButton', font=('Open Sans', 11, 'bold'), padding=(25, 12),
                             background=alt_button_bg, foreground=alt_button_fg)
        self.style.map('Alt.TButton', background=[('active', '#C8C8C8')])

        # --- Specific Styles from S2 GUI Structure ---
        self.style.configure('Title.TLabel', font=('Open Sans', 32, 'bold'), foreground='#1976D2', background=secondary_color)
        self.style.configure('Subtitle.TLabel', font=('Open Sans', 10), foreground=status_fg, background=secondary_color) # Smaller subtitle font

        # Sort/Search Controls Styles from S2
        self.style.configure('Sort.TLabel', font=('Open Sans', 10), background=secondary_color, foreground=status_fg)
        self.style.configure('Sort.TCombobox', padding=5)
        self.style.map('Sort.TCombobox', fieldbackground=[('readonly', 'white')])
        self.style.configure('Search.TEntry', padding=5, fieldbackground='white')
        self.style.configure('Search.TButton', font=('Open Sans', 9), padding=(5, 3)) # Smaller search buttons

        # Status Bar Style
        self.style.configure('Status.TFrame', background=secondary_color) # Match main background
        self.style.configure('Status.TLabel', background=secondary_color, foreground=status_fg, font=('Open Sans', 9))

        # Dialog styles (Keep S1's setup for consistency in dialogs)
        self.style.configure('Dialog.TFrame', background=dialog_bg)
        self.style.configure('Dialog.TLabel', background=dialog_bg, foreground=text_color)
        self.style.configure('DialogTitle.TLabel', font=('Segoe UI', 14, 'bold'), background=dialog_bg, foreground='#1976D2') # Use S2 blue
        self.style.configure('DialogSubtitle.TLabel', font=('Segoe UI', 10, 'italic'), background=dialog_bg, foreground=status_fg)
        self.style.configure('DialogStatus.TLabel', font=('Segoe UI', 9), background=dialog_bg, foreground=status_fg)
        self.style.configure('Dialog.TButton', font=('Segoe UI', 10), padding=(10, 6))
        self.style.configure('SmallDialog.TButton', font=('Segoe UI', 9), padding=(5, 3))
        self.style.configure('DialogAlt.TButton', font=('Segoe UI', 10), padding=(10, 6), background=alt_button_bg, foreground=alt_button_fg)
        self.style.map('DialogAlt.TButton', background=[('active', '#C8C8C8')])
        self.style.configure('DialogHighlight.TButton', font=('Segoe UI', 10, 'bold'), padding=(10, 6), background='#28A745', foreground='white')
        self.style.map('DialogHighlight.TButton', background=[('active', '#218838')])
        self.style.configure('Dialog.TEntry', padding=5, fieldbackground='white')
        self.style.configure('Dialog.TCombobox', padding=5)
        self.style.map('Dialog.TCombobox', fieldbackground=[('readonly', 'white')])

        # Splash screen styles
        self.style.configure('Splash.TFrame', background='#FFFFFF')
        self.style.configure('Splash.TLabel', background='#FFFFFF', foreground=text_color)

        # Treeview styling (Keep S1's for Session/Global Search)
        self.style.configure("Treeview", rowheight=25, font=('Segoe UI', 9), fieldbackground='white')
        self.style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), padding=5)
        self.style.map("Treeview", background=[('selected', '#0078D4')], foreground=[('selected', 'white')])

    def apply_styles(self):
        """Apply loaded/current settings to widget styles where needed."""
        # Primarily for Text widget tags based on settings
        self.configure_text_tags()

    def configure_text_tags(self):
        """Configure tags for the results Text widget based on settings."""
        if not hasattr(self, 'results_text'): return # Widget not created yet

        # Header tags (fixed colors like S2)
        self.results_text.tag_configure('header_base', font=('Consolas', 11, 'bold'), background='#f0f2f5')
        self.results_text.tag_configure('header_nil', font=('Consolas', 11, 'bold'), background='#D9D9D9')
        self.results_text.tag_configure('header_tb1', font=('Consolas', 11, 'bold'), background='#B3E6B3')
        self.results_text.tag_configure('header_tb2', font=('Consolas', 11, 'bold'), background='#FFF2B3')
        self.results_text.tag_configure('header_mit', font=('Consolas', 11, 'bold'), background='#D9B3E6')
        self.results_text.tag_configure('header_qft', font=('Consolas', 11, 'bold'), background='#87CEEB')
        self.results_text.tag_configure('header_comment', font=('Consolas', 11, 'bold'), background='#E0E0E0')

        # Data row tags (using settings) - Background applied to whole row
        self.results_text.tag_configure('pos_row', background=app_settings.get('pos_bg','#FFFFE0'))
        self.results_text.tag_configure('neg_row', background=app_settings.get('neg_bg','#FFFFFF'))
        self.results_text.tag_configure('ind_row', background=app_settings.get('ind_bg','#FFFFFF'))
        self.results_text.tag_configure('wp_row', background=app_settings.get('wp_bg','#FFF8DC')) # Added WP row BG

        # Specific cell tags (primarily for QFT result text color)
        self.results_text.tag_configure('qft_pos', foreground=app_settings.get('pos_text','#e53935'))
        self.results_text.tag_configure('qft_neg', foreground=app_settings.get('neg_text','#43a047'))
        self.results_text.tag_configure('qft_ind', foreground=app_settings.get('ind_text','#fb8c00'))
        self.results_text.tag_configure('qft_wp', foreground=app_settings.get('wp_text','#D2691E')) # Added WP text color
        # Comment tag
        self.results_text.tag_configure('comment', font=('Consolas', 10, 'italic'), foreground='#666666')
        # Search highlight tag
        self.results_text.tag_configure('search_highlight', background='yellow', foreground='black')

    def create_menu(self):
        """Create the main application menu bar (Adapted from S2 for structure)."""
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        # --- File Menu (Structure from S2) ---
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)

        # Import/Export submenu (like S2)
        import_export_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Import/Export", menu=import_export_menu)
        # Commands point to S1 functions
        import_export_menu.add_command(label="Import Data...", accelerator="Ctrl+O", command=lambda: import_data(add_mode=False))
        import_export_menu.add_command(label="Add Data...", command=lambda: import_data(add_mode=True))
        import_export_menu.add_separator()
        import_export_menu.add_command(label="Export Options...", accelerator="Ctrl+E", command=show_export_options) # Use combined options dialog

        # Session Management submenu (like S2)
        session_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Session", menu=session_menu)
        # Commands point to S1 functions
        session_menu.add_command(label="Save Session", accelerator="Ctrl+S", command=lambda: save_session(auto_save=False))
        session_menu.add_command(label="Manage Sessions...", accelerator="Ctrl+L", command=manage_sessions) # Renamed from Load Session

        file_menu.add_separator()
        file_menu.add_command(label="Clear Current Data", command=self.clear_data) # Use S1 clear function
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing) # Use S1 closing function

        # --- Edit Menu (Keep S1's advanced options + S2 structure) ---
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Copy", accelerator="Ctrl+C", command=lambda: self.master.focus_get().event_generate('<<Copy>>'))
        edit_menu.add_command(label="Select All", accelerator="Ctrl+A", command=self._handle_select_all) # Helper needed
        edit_menu.add_separator()
        edit_menu.add_command(label="Manual Reorder...", accelerator="Ctrl+M", command=show_manual_order_dialog)

        # --- Search Menu (Like S2) ---
        search_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Search", menu=search_menu)
        search_menu.add_command(label="Global Sample Search...", accelerator="Ctrl+F", command=show_global_search)

        # --- View Menu (Combine S1 and S2 items) ---
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        # view_menu.add_command(label="Export Options", command=show_export_options) # Already in File menu
        view_menu.add_command(label="Customize Appearance...", command=customize_appearance) # Keep S1's customization
        view_menu.add_separator()
        # Decimal Places submenu (like S2)
        self.decimal_menu = tk.Menu(view_menu, tearoff=0) # Store reference to update later
        view_menu.add_cascade(label="Decimal Places", menu=self.decimal_menu)
        self.update_decimal_menu() # Populate the menu


        # --- Help Menu (Like S2) ---
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about) # Use S1's detailed about

        # Bind keyboard shortcuts (Keep S1's)
        self.master.bind_all("<Control-o>", lambda e: import_data(add_mode=False))
        self.master.bind_all("<Control-e>", lambda e: show_export_options())
        self.master.bind_all("<Control-s>", lambda e: save_session(auto_save=False))
        self.master.bind_all("<Control-l>", lambda e: manage_sessions())
        self.master.bind_all("<Control-m>", lambda e: show_manual_order_dialog())
        self.master.bind_all("<Control-f>", lambda e: show_global_search())
        # Ctrl+A handled by _handle_select_all

    def _handle_select_all(self, event=None):
        """Helper to call SelectAll on the currently focused widget."""
        try:
            widget = self.master.focus_get()
            if isinstance(widget, tk.Text):
                widget.tag_add(tk.SEL, "1.0", tk.END)
                widget.mark_set(tk.INSERT, "1.0")
                widget.see(tk.INSERT)
            elif isinstance(widget, ttk.Entry) or isinstance(widget, tk.Entry):
                 widget.select_range(0, tk.END)
                 widget.icursor(tk.END)
        except Exception as e:
            print(f"Select All error: {e}")

    def update_decimal_menu(self):
        """Updates the View > Decimal Places menu."""
        self.decimal_menu.delete(0, tk.END) # Clear existing items

        # Use the global app_settings for the variable
        current_setting = app_settings.get('decimal_places', 'default')

        options = {'Default': 'default', '0 Decimals': '0', '1 Decimal': '1', '2 Decimals': '2', '3 Decimals': '3'}

        # Create a Tkinter variable linked to the setting if not already done elsewhere
        if not hasattr(self, 'decimal_var_menu'):
             self.decimal_var_menu = tk.StringVar(value=str(current_setting))

        for label, value in options.items():
            self.decimal_menu.add_radiobutton(
                label=label,
                variable=self.decimal_var_menu, # Use the class variable
                value=str(value),
                command=lambda v=value: self.set_decimal_places(v),
                selectcolor=self.style.lookup('TCheckbutton', 'indicatorcolor') # Match theme indicator
            )
        # Ensure the menu variable reflects the actual current setting
        self.decimal_var_menu.set(str(current_setting))

    def set_decimal_places(self, value):
        """Callback function when a decimal place option is selected from the menu."""
        global app_settings
        print(f"Setting decimal places to: {value}")
        app_settings['decimal_places'] = str(value)
        self.decimal_var_menu.set(str(value)) # Update menu variable state
        save_settings(app_settings)
        self.refresh_display()
        self.update_status(f"Decimal places set to '{value}'.")


    def create_widgets(self):
        """Create all the widgets for the main application window (Layout from Script 2)."""
        # Configure root window grid (No change)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_rowconfigure(4, weight=1) # Make results area expand

        # Main Frame (like S2)
        main_frame = ttk.Frame(self.master, style='TFrame')
        main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

        # Configure main_frame grid for centering
        main_frame.grid_columnconfigure(0, weight=1) # Left spacer column
        main_frame.grid_columnconfigure(1, weight=0) # Central column for content (auto-sizes)
        main_frame.grid_columnconfigure(2, weight=1) # Right spacer column

        # Title and Subtitle (Place in the central column 1)
        title_label = ttk.Label(main_frame, text="QFT-Plus Data Viewer", style='Title.TLabel', anchor='center')
        title_label.grid(row=0, column=1, pady=(0, 5), sticky='ew')
        subtitle_label = ttk.Label(main_frame, text="Import, View, and Export QuantiFERON Data", style='Subtitle.TLabel', anchor='center')
        subtitle_label.grid(row=1, column=1, pady=(0, 25), sticky='ew')

        # Button Frame (Place in the central column 1)
        # This frame will shrink/grow only to fit the buttons packed inside it.
        # The main_frame's grid (cols 0 and 2) will handle the centering.
        button_frame = ttk.Frame(main_frame, style='TFrame')
        button_frame.grid(row=2, column=1, pady=15) # Grid in central column 1

        # --- Create Buttons (remain the same) ---
        import_button = ttk.Button(button_frame, text="📂 Import Data", command=lambda: import_data(add_mode=False), style='Custom.TButton')
        add_button = ttk.Button(button_frame, text="➕ Add Data", command=lambda: import_data(add_mode=True), style='Custom.TButton')
        export_button = ttk.Button(button_frame, text="📄 Export Options", command=show_export_options, style='Custom.TButton')
        session_button = ttk.Button(button_frame, text="💾 Sessions", command=manage_sessions, style='Custom.TButton')
        clear_button = ttk.Button(button_frame, text="🗑️ Clear Data", command=self.clear_data, style='Alt.TButton')

        # --- REVERTED: Use pack side-by-side within button_frame ---
        import_button.pack(side='left', padx=5)
        add_button.pack(side='left', padx=5)
        export_button.pack(side='left', padx=5)
        session_button.pack(side='left', padx=5)
        clear_button.pack(side='left', padx=5) # Use pack like the others

        # --- Controls Frame (Search and Sort) - Remains the same ---
        controls_frame = ttk.Frame(main_frame, style='TFrame')
        controls_frame.grid(row=3, column=1, pady=(20, 10), sticky='ew') # Grid in column 1
        controls_frame.grid_columnconfigure(1, weight=1)
        # Search Controls (remain the same internal packing)
        search_subframe = ttk.Frame(controls_frame, style='TFrame')
        search_subframe.grid(row=0, column=0, sticky='w')
        ttk.Label(search_subframe, text="Search Sample ID:", style='Sort.TLabel').pack(side='left', padx=(0, 8))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_subframe, textvariable=self.search_var, width=20, style='Search.TEntry')
        search_entry.pack(side='left', padx=(0, 8))
        search_entry.bind('<Return>', self.perform_search)
        search_context_menu = CustomContextMenu(search_entry)
        search_entry.bind("<Button-3>", search_context_menu.show)
        search_entry.bind('<Control-v>', self._handle_paste)
        search_entry.bind('<Command-v>', self._handle_paste)
        search_btn = ttk.Button(search_subframe, text="🔍 Search", style='Search.TButton', command=self.perform_search)
        search_btn.pack(side='left', padx=(0, 5))
        clear_search_btn = ttk.Button(search_subframe, text="❌ Clear", style='Search.TButton', command=self.clear_search)
        clear_search_btn.pack(side='left')
        # Sort Controls (remain the same internal packing)
        sort_subframe = ttk.Frame(controls_frame, style='TFrame')
        sort_subframe.grid(row=0, column=2, sticky='e')
        ttk.Label(sort_subframe, text="Sort Results by:", style='Sort.TLabel').pack(side='left', padx=(0, 8))
        self.sort_var = tk.StringVar()
        self.sort_options = {
             "Date (Oldest First)": ('requested_date', 'asc'),
             "Date (Newest First)": ('requested_date', 'desc'),
             "Sample ID (A-Z)": ('barcode', 'asc'),
             "Sample ID (Z-A)": ('barcode', 'desc'),
             "QFT Result (A-Z)": ('qft_result', 'asc'),
             "QFT Result (Z-A)": ('qft_result', 'desc'),
             "Manual Order": (None, None)
         }
        self.sort_dropdown = ttk.Combobox(sort_subframe, textvariable=self.sort_var,
                                          values=list(self.sort_options.keys()),
                                          width=20, state="readonly", style='Sort.TCombobox')
        self.sort_dropdown.set("Date (Oldest First)")
        self.sort_dropdown.pack(side='left')
        self.sort_dropdown.bind('<<ComboboxSelected>>', self.sort_data)

        # --- Results Text Area and Scrollbars (Remain the same) ---
        self.results_text = tk.Text(
            main_frame, width=self.total_width, height=25, font=('Consolas', 11),
            background='white', foreground='#333333', padx=10, pady=10,
            selectbackground='#90CAF9', selectforeground='black',
            state='disabled', wrap='none'
        )
        self.results_text.grid(row=4, column=1, sticky="nsew", pady=(5, 5)) # Place in central col 1
        vsb = ttk.Scrollbar(main_frame, orient="vertical", command=self.results_text.yview)
        vsb.grid(row=4, column=2, sticky="ns") # Place in right spacer col 2
        self.results_text.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=self.results_text.xview)
        hsb.grid(row=5, column=1, sticky="ew") # Place below central col 1
        self.results_text.configure(xscrollcommand=hsb.set)
        self.results_context_menu = CustomContextMenu(self.results_text)
        self.results_text.bind("<Button-3>", self.results_context_menu.show)

        # --- Status Bar (Remains the same) ---
        self.status_frame = ttk.Frame(main_frame, style='Status.TFrame')
        self.status_frame.grid(row=6, column=1, sticky="ew", pady=(10, 0)) # Place in central col 1
        self.status_frame.columnconfigure(1, weight=1)
        self.status_label = ttk.Label(self.status_frame, text="Initializing...", style='Status.TLabel', anchor='w')
        self.status_label.grid(row=0, column=0, sticky='w', padx=5)
        self.progress_bar = ttk.Progressbar(self.status_frame, mode='indeterminate', length=150)
        footer_text = "© 2024 Created by Hosam Al Mokashefy and Aly Sherif (AST Team)"
        self.version_label = ttk.Label(self.status_frame, text=footer_text, style='Status.TLabel', anchor='e')
        self.version_label.grid(row=0, column=2, sticky='e', padx=5)

        # --- Configure main_frame row/column weights ---
        main_frame.grid_rowconfigure(4, weight=1) # Allow results text area to expand vertically
    # --- Data Handling Methods (Keep Script 1's) ---

    def set_data_rows(self, list_of_dicts):
        """Clears existing data and sets new data (Keep Script 1)."""
        print(f"--- Inside set_data_rows ---")
        print(f"Received {len(list_of_dicts)} records.")
        if list_of_dicts:
             print(f"Sample received record (raw): {list_of_dicts[0]}")

        processed_rows = []
        required_keys = ['barcode', 'nil_result', 'tb1_result', 'tb2_result', 'mit_result', 'tb1_nil', 'tb2_nil', 'mit_nil', 'qft_result', 'requested_date']

        for i, row in enumerate(list_of_dicts):
            processed_row = {}
            # Ensure essential keys exist, default to empty string or None
            for key in required_keys:
                processed_row[key] = str(row.get(key, '')) # Convert all to string initially

            # Handle potential case difference from import rename mapping
            processed_row['barcode'] = str(row.get('barcode', row.get('Barcode', '')))
            processed_row['requested_date'] = row.get('requested_date', row.get('RequestedDate')) # Keep original type for now

            # Process the date specifically
            date_val = processed_row['requested_date']
            parsed_date = None
            if isinstance(date_val, datetime.datetime):
                 parsed_date = date_val # Already correct type
            elif isinstance(date_val, str):
                try:
                    # Attempt parsing from expected DB format first if loading session
                    parsed_date = datetime.datetime.strptime(date_val, '%Y-%m-%d %H:%M:%S')
                except (ValueError, TypeError):
                    try:
                         # Attempt parsing from import format '%d/%m/%Y %H:%M:%S'
                         parsed_date = datetime.datetime.strptime(date_val, '%d/%m/%Y %H:%M:%S')
                    except (ValueError, TypeError):
                         print(f"Warning: Could not parse date string '{date_val}' for barcode {processed_row['barcode']}. Setting to None.")
                         parsed_date = None # Failed parsing string
            elif pd.isna(date_val): # Handle Pandas NaT
                 parsed_date = None
            else:
                 # If it's None, or some other non-string/non-datetime type
                 parsed_date = None

            processed_row['requested_date'] = parsed_date # Assign the parsed datetime object or None

            processed_rows.append(processed_row)
            # if i < 5: # Print first few processed rows
            #      print(f"Processed row {i}: {processed_row}")

        self.current_data = processed_rows # Assign the processed list back
        print(f"Stored {len(self.current_data)} records in self.current_data.")
        if self.current_data:
             print(f"First stored record (processed): {self.current_data[0]}")
        # self.update_data_info() # Placeholder call - not implemented

    def add_data_rows(self, list_of_dicts):
        """Adds new rows, ensuring date conversion (Keep Script 1)."""
        processed_rows = []
        for row in list_of_dicts:
             # Process date similar to set_data_rows
             date_val = row.get('requested_date', row.get('RequestedDate'))
             parsed_date = None
             if isinstance(date_val, datetime.datetime):
                 parsed_date = date_val
             elif isinstance(date_val, str):
                  try:
                       parsed_date = datetime.datetime.strptime(date_val, '%Y-%m-%d %H:%M:%S')
                  except (ValueError, TypeError):
                     try:
                         parsed_date = datetime.datetime.strptime(date_val, '%d/%m/%Y %H:%M:%S')
                     except (ValueError, TypeError):
                          parsed_date = None
             elif pd.isna(date_val):
                 parsed_date = None
             else:
                 parsed_date = None
             row['requested_date'] = parsed_date
             # Convert other fields to string for consistency
             for key in row:
                 if key != 'requested_date':
                     row[key] = str(row.get(key,''))
             processed_rows.append(row)

        self.current_data.extend(processed_rows)
        # self.update_data_info() # Placeholder call

    def has_data(self):
        """Checks if there is data loaded (Keep Script 1)."""
        return bool(self.current_data)

    def get_all_barcodes(self):
        """Returns a list of all barcodes currently loaded (Keep Script 1)."""
        return [str(row.get('barcode', '')) for row in self.current_data]

    # def update_data_info(self):
    #      """Placeholder for updating info like date ranges if needed."""
    #      pass # Not implemented

    def get_report_date_str(self):
         """Gets the latest date from the data or current date for reports (Keep Script 1)."""
         latest_date = None
         for row in self.current_data:
              dt = row.get('requested_date')
              if isinstance(dt, datetime.datetime):
                   if latest_date is None or dt > latest_date:
                        latest_date = dt
         return latest_date.strftime('%Y-%m-%d') if latest_date else datetime.datetime.now().strftime('%Y-%m-%d')


    def clear_data(self):
        """Clears all data from the application after confirmation (Keep Script 1)."""
        if not self.has_data():
            messagebox.showinfo("Clear Data", "There is no data loaded to clear.", parent=self.master)
            return

        confirm = messagebox.askyesnocancel(
            "Confirm Clear",
            "This will remove all currently loaded data.\nDo you want to save the current session first?",
            icon='warning', parent=self.master
        )
        if confirm is None: # Cancel
            return
        elif confirm: # Yes, save first
            # Use the current session name if available for the save prompt
            current_session_name = None
            if self.imported_filename_source.startswith("Session "):
                 current_session_name = self.imported_filename_source.replace("Session ", "")

            if not save_session(auto_save=False, session_name_in=current_session_name): # Pass current name if overwriting
                 return # Don't proceed with clear if save was cancelled

        # Proceed with clearing (after potential save)
        self.current_data = []
        self.imported_filename_source = ""
        self.refresh_display() # Update the empty view
        self.search_var.set("") # Clear search box
        self.sort_dropdown.set("Date (Oldest First)") # Reset sort to S2 default
        self.update_status("Data cleared.")
        messagebox.showinfo("Clear Data", "Current data has been cleared.", parent=self.master)

    # --- Display and Formatting ---

    def refresh_display(self):
        """Populates the results Text widget with current_data (S2 Style)."""
        print("--- Refreshing Display ---")
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)

        if not self.current_data:
            self.results_text.insert('1.0', "No data loaded. Use File > Import/Export > Import or File > Session > Manage Sessions.")
            self.results_text.config(state='disabled')
            print("No data to display.")
            return

        # --- Headers (Like S2) ---
        header_line = ""
        tag_positions = [] # Store (tag, start_char, end_char)
        current_pos = 0
        # Use S2 header texts from self.headers_info
        header_texts = [h[0] for h in self.headers_info]
        # Use 'header_base' for Request Date
        self.headers_info[-1] = (self.headers_info[-1][0], 'header_base', self.headers_info[-1][2])

        for i, (text, tag, width) in enumerate(self.headers_info):
             # Pad text to fit width, add space separator
             padded_text = f"{text:<{width}}"
             header_line += padded_text + (" " if i < len(self.headers_info) - 1 else "") # Add space between columns
             tag_positions.append((tag, current_pos, current_pos + width))
             current_pos += width + 1 # Add 1 for the space separator

        self.results_text.insert('1.0', header_line + "\n") # Insert header line first
        # Apply specific background tags to header segments
        for tag, start, end in tag_positions:
            if tag != 'header_base': # Only apply specific background if not base
                 self.results_text.tag_add(tag, f"1.{start}", f"1.{end}")

        # Separator line (Match width, accounting for spaces)
        separator = "-" * current_pos
        self.results_text.insert('end', separator + "\n")
        print(f"Displayed Headers. Total width: {current_pos}")

        # --- Data Rows ---
        decimals = app_settings['decimal_places']
        print(f"Using decimal places: {decimals}")
        row_count = 0
        for row_dict in self.current_data:
            row_count += 1
            line_start_index = self.results_text.index('end-1c linestart') # Get index before inserting

            # *** Calculate comment ONCE per row ***
            comment = calculate_comment(row_dict)
            is_wp = "WP" in comment
            qft_result_final = str(row_dict.get('qft_result', ' ')).upper()

            # Format data for display based on headers_info
            data_line = ""
            tags_to_apply = [] # List of (tag, start_index, end_index) for this line
            current_pos_data = 0

            # Map keys from row_dict to the order in headers_info
            key_map = {
                 "Barcode": "barcode", "Nil_Result": "nil_result", "TB1_Result": "tb1_result",
                 "TB2_Result": "tb2_result", "Mit_Result": "mit_result", "TB1_Nil": "tb1_nil",
                 "TB2_Nil": "tb2_nil", "Mit_Nil": "mit_nil", "QFT_Result": "qft_result",
                 "Comment": None, # Calculated
                 "Request Date": "requested_date"
             }

            for i, (header_text, _, width) in enumerate(self.headers_info):
                value_str = " " # Default to space
                key = key_map.get(header_text)

                if key: # Standard key lookup
                     raw_value = row_dict.get(key, " ")
                     if key == 'requested_date':
                         value_str = raw_value.strftime('%Y-%m-%d %H:%M:%S') if isinstance(raw_value, datetime.datetime) else 'No Date'
                     elif key == 'qft_result':
                          value_str = str(raw_value).upper()
                     elif key == 'barcode':
                          value_str = str(raw_value)
                     else: # Numeric fields
                         value_str = format_number_with_decimals(str(raw_value), decimals)
                elif header_text == "Comment": # Calculate comment
                     value_str = calculate_comment(row_dict)

                # Pad/truncate value to fit the column width
                padded_value = f"{str(value_str):<{width}}"[:width] # Truncate if too long
                data_line += padded_value + (" " if i < len(self.headers_info) - 1 else "") # Add space separator

                # --- Determine Cell/Row Tags ---
                start_idx_rel_char = current_pos_data
                end_idx_rel_char = start_idx_rel_char + width
                tag_start = f"{line_start_index}+{start_idx_rel_char}c"
                tag_end = f"{line_start_index}+{end_idx_rel_char}c"

                # Specific tags for QFT result cell (text color only)
                if header_text == "QFT_Result":
                     qft_cell_tag = None
                     if is_wp: qft_cell_tag = 'qft_wp' # WP first
                     elif qft_result_final in ('POS', 'POS*'): qft_cell_tag = 'qft_pos'
                     elif qft_result_final == 'NEG': qft_cell_tag = 'qft_neg'
                     elif qft_result_final == 'IND': qft_cell_tag = 'qft_ind'
                     if qft_cell_tag:
                         tags_to_apply.append((qft_cell_tag, tag_start, tag_end))

                # Tag for comment cell
                if header_text == "Comment":
                      tags_to_apply.append(('comment', tag_start, tag_end))

                current_pos_data += width + 1 # Add 1 for space

            # Insert the full data line
            self.results_text.insert('end', data_line + "\n")
            line_end_index = self.results_text.index('end-1c lineend') # End of inserted line

            # Apply row background color based on QFT result
            row_tag = None
            if is_wp: row_tag = 'wp_row' # WP first
            elif qft_result_final in ('POS', 'POS*'): row_tag = 'pos_row'
            elif qft_result_final == 'NEG': row_tag = 'neg_row'
            elif qft_result_final == 'IND': row_tag = 'ind_row'

            if row_tag:
                # Apply background tag to the entire line
                self.results_text.tag_add(row_tag, line_start_index, line_end_index)

            # Apply specific cell tags collected earlier (text colors, comment style)
            for tag, start, end in tags_to_apply:
                try:
                    self.results_text.tag_add(tag, start, end)
                except tk.TclError as tag_error:
                     print(f"Error applying tag '{tag}' from {start} to {end}: {tag_error}")


        print(f"Displayed {row_count} data rows.")
        # Apply search highlights if search term exists
        self.apply_search_highlight()

        self.results_text.config(state='disabled')
        # Scroll to top after refresh
        self.results_text.yview_moveto(0)
        print("--- Refresh Display Complete ---")


    def sort_data(self, event=None):
        """Sorts the internal data (self.current_data) and refreshes display (Keep Script 1)."""
        sort_key_display = self.sort_var.get()
        if not sort_key_display or not self.has_data():
            print(f"Sort skipped: Key='{sort_key_display}', HasData={self.has_data()}")
            return # No sort criteria or no data

        if sort_key_display == "Manual Order":
             # If user selects Manual Order, don't re-sort, just keep current order
             self.current_sort_column = None
             self.current_sort_direction = None
             print("Manual order selected. No automatic sorting applied.")
             # Optionally refresh display if needed, but data order shouldn't change
             # self.refresh_display()
             return

        sort_column, direction = self.sort_options.get(sort_key_display, (None, None))

        if not sort_column:
            print(f"Invalid sort key: {sort_key_display}")
            return

        print(f"Sorting by: {sort_column}, Direction: {direction}")
        self.current_sort_column = sort_column
        self.current_sort_direction = direction
        reverse_order = (direction == 'desc')

        # Sorting logic
        try:
            # Use a lambda function for sorting, handling potential None or errors
            def sort_func(row_dict):
                value = row_dict.get(sort_column)

                # Handle None or missing values - place them consistently
                if value is None or (isinstance(value, str) and value.strip() == ''):
                    # Place Nones/empty last when ascending, first when descending
                    return (float('inf') if not reverse_order else float('-inf'))

                if sort_column == 'requested_date':
                    # Should be datetime objects, direct comparison works
                    if isinstance(value, datetime.datetime):
                        return value
                    else: # Should not happen if data processing is correct
                         return (datetime.datetime.max if not reverse_order else datetime.datetime.min)

                elif sort_column == 'barcode':
                     # Basic string sort, consider natsort library for true natural sort if needed
                     return str(value)

                elif sort_column == 'qft_result':
                    # Define custom order for QFT results
                    order = {'POS*': 0, 'POS': 1, 'IND': 2, 'NEG': 3}
                    return order.get(str(value).upper(), 99) # Place unknowns last

                else: # Assume numeric sort for other columns
                    try:
                         # Handle comparison symbols if sorting numeric columns
                         val_str = str(value).replace('>', '').replace('<', '').strip()
                         if not val_str or val_str == " ":
                              return (float('inf') if not reverse_order else float('-inf'))
                         return float(val_str)
                    except (ValueError, TypeError):
                         # If conversion fails, treat as string or place based on order
                         print(f"Warning: Non-numeric value '{value}' found in supposedly numeric column '{sort_column}' during sort.")
                         # Fallback: treat non-numeric as very large/small based on direction
                         return (float('inf') if not reverse_order else float('-inf'))


            self.current_data.sort(key=sort_func, reverse=reverse_order)
            self.refresh_display() # Update view with sorted data
            self.update_status(f"Data sorted by {sort_key_display}.")

        except Exception as e:
            messagebox.showerror("Sort Error", f"Could not sort data: {e}", parent=self.master)
            traceback.print_exc()

    # --- Search (Keep Script 1's logic) ---
    def perform_search(self, event=None):
        """Highlights rows matching the search term in the barcode."""
        search_term = self.search_var.get().strip().lower()
        found_count = self.apply_search_highlight() # Apply based on current term, get count
        if search_term:
            if found_count > 0:
                # Find the first match and scroll to it
                first_match_index = self.find_first_search_match(search_term)
                if first_match_index:
                    self.results_text.see(first_match_index)
                self.update_status(f"Found {found_count} match(es) for '{search_term}'.")
            else:
                self.update_status(f"No matches found for '{search_term}'.")
        else:
            self.update_status("Search cleared.")


    def clear_search(self):
        """Clears the search term and removes highlights."""
        self.search_var.set("")
        self.apply_search_highlight() # Removes highlights when term is empty
        self.update_status("Search cleared.")

    def apply_search_highlight(self):
        """Applies or removes search highlights based on self.search_var. Returns count."""
        search_term = self.search_var.get().strip().lower()
        match_count = 0
        self.results_text.config(state='normal')
        self.results_text.tag_remove('search_highlight', '1.0', tk.END) # Clear previous

        if search_term and self.current_data:
            # Iterate through lines in the Text widget more carefully
            # Header lines = 2 (Header + Separator)
            start_line_index = 3 # Data starts on line 3 (1-based index)
            end_line_index = int(self.results_text.index('end-1c').split('.')[0])

            # Determine the character range for the barcode column
            barcode_col_index = -1
            barcode_start_char = 0
            barcode_end_char = 0
            current_char = 0
            for i, (text, _, width) in enumerate(self.headers_info):
                if text == "Barcode":
                    barcode_col_index = i
                    barcode_start_char = current_char
                    barcode_end_char = current_char + width
                    break
                current_char += width + 1 # Account for space separator

            if barcode_col_index != -1:
                for i in range(start_line_index, end_line_index + 1):
                    line_content = self.results_text.get(f"{i}.0", f"{i}.end")
                    # Extract barcode using character indices
                    barcode_in_line = line_content[barcode_start_char:barcode_end_char].strip().lower()

                    if search_term in barcode_in_line:
                        self.results_text.tag_add('search_highlight', f"{i}.0", f"{i}.end")
                        match_count += 1

        self.results_text.config(state='disabled')
        return match_count # Return the number of matches found

    def find_first_search_match(self, search_term):
        """Finds the index ('line.char') of the first search match (Keep Script 1)."""
        if not search_term or not self.current_data:
            return None

        search_term = search_term.lower()
        # Header lines = 2 (Header + Separator)
        start_line_index = 3 # Data starts on line 3 (1-based index)
        end_line_index = int(self.results_text.index('end-1c').split('.')[0])

        barcode_col_index = -1
        barcode_start_char = 0
        barcode_end_char = 0
        current_char = 0
        for i, (text, _, width) in enumerate(self.headers_info):
            if text == "Barcode":
                 barcode_col_index = i
                 barcode_start_char = current_char
                 barcode_end_char = current_char + width
                 break
            current_char += width + 1

        if barcode_col_index != -1:
            for i in range(start_line_index, end_line_index + 1):
                line_content = self.results_text.get(f"{i}.0", f"{i}.end")
                barcode_in_line = line_content[barcode_start_char:barcode_end_char].strip().lower()
                if search_term in barcode_in_line:
                    return f"{i}.0" # Return start index of the line
        return None


    # --- Status Bar & User Feedback (Keep Script 1's) ---
    def update_status(self, message, show_progress=False, hide_progress=False):
        """Updates the status bar message and optionally shows/hides progress bar."""
        self.status_label.config(text=message)
        if show_progress:
            self.progress_bar.grid(row=0, column=1, sticky='ew', padx=(10, 10)) # Place between status and version
            self.progress_bar.start(10)
        elif hide_progress or not show_progress: # Ensure hide if not explicitly shown
            self.progress_bar.stop()
            self.progress_bar.grid_remove() # Hide progress bar
        self.master.update_idletasks() # Force UI update

    # --- Utility & Closing (Keep Script 1's) ---
    def get_data_for_export(self):
        """Returns the current data, potentially sorted or filtered."""
        # Currently returns the data as sorted/ordered in self.current_data
        return self.current_data

    def _handle_paste(self, event=None):
        """Handles pasting into the search entry, cleaning input."""
        try:
            clipboard_content = self.master.clipboard_get()
            cleaned_content = clipboard_content.strip() # Remove leading/trailing whitespace
            self.search_var.set(cleaned_content) # Set the cleaned content
            self.perform_search() # Optionally trigger search immediately after paste
        except tk.TclError:
            pass # Ignore if clipboard is empty or invalid
        return "break" # Prevent default paste behavior

    def show_about(self):
        """Displays the About dialog (Keep Script 1's detailed version)."""
        messagebox.showinfo(
            "About QFT-Plus Viewer",
            "QFT-Plus Data Viewer v2.1\n\n" # Update version
            "This application helps import, visualize, format, and export\n"
            "results generated by the Diasorin LIAISON® Quantiferon Software.\n\n"
            "Features:\n"
            " - Import from Excel (.xlsx, .xls) and CSV (.csv)\n"
            " - Customizable result highlighting and decimal places\n"
            " - Flexible sorting options and manual reordering\n"
            " - Export to formatted PDF, Excel, and CSV (Script 2 Style)\n"
            " - Session management (Save/Load/Manage/Rename)\n"
            " - Global search across all saved sessions\n\n"
            "Disclaimer: This viewer is for visualization and formatting purposes only.\n"
            "It does not perform any calculations or interpretations of the results.\n"
            "All calculations originate from the source LQS software.\n\n"
            "Developed by: Hosam Al Mokashefy & Aly Sherif (AST Team)\n"
            "© 2024",
            parent=self.master
        )

    def on_closing(self):
        """Handles the window close event (Keep Script 1's save confirmation)."""
        if self.has_data():
            confirm = messagebox.askyesnocancel(
                "Exit Confirmation",
                "Do you want to save the current session before exiting?",
                icon='question', parent=self.master
            )
            if confirm is None: # Cancel
                return
            elif confirm: # Yes, save
                 # Use the current session name if available for the save prompt
                 current_session_name = None
                 if self.imported_filename_source.startswith("session "):
                      current_session_name = self.imported_filename_source.replace("session ", "")
                 if not save_session(auto_save=False, session_name_in=current_session_name):
                     # User cancelled the save dialog, so don't exit
                     return
            # If 'No' or save successful, proceed to exit
        # else: No data, just exit cleanly

        # Add any cleanup tasks here (e.g., closing files, threads)
        print("Exiting application...")
        try:
            self.master.destroy()
        except tk.TclError:
             pass # Ignore errors if window already destroyed
        sys.exit(0) # Force exit

# --- Main Execution (Keep Script 1's structure) ---

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw() # Hide the main window initially

    # Create database and directories if they don't exist
    try:
        get_app_data_dir() # Ensure directory exists
        db_conn_test = get_database_connection()
        if db_conn_test:
            db_conn_test.close()
        else:
             # Show error and exit if DB connection failed critically
             messagebox.showerror("Fatal Error", "Could not initialize the application database. Exiting.")
             sys.exit(1)
    except Exception as init_err:
        messagebox.showerror("Fatal Error", f"Application initialization failed: {init_err}. Exiting.")
        traceback.print_exc()
        sys.exit(1)


    # Show Splash Screen
    splash = SplashScreen(root)
    root.update_idletasks()

    # Initialize the main application after splash setup
    main_app = QFTApp(root)

    # Splash screen's close_splash method will call root.deiconify()

    # Ensure the main window is shown even if splash fails (add failsafe)
    # Check if splash window still exists before trying to schedule deiconify based on it
    try:
         if splash.winfo_exists():
             root.after(splash.display_duration + 500, lambda: root.deiconify() if root.state() == 'withdrawn' else None)
         else: # If splash closed early or failed, show main window sooner
             root.deiconify()
    except tk.TclError: # Handle cases where splash might be destroyed unexpectedly
        if root.state() == 'withdrawn':
             root.deiconify()


    root.mainloop()