# gui.py
import tkinter as tk
from tkinter import filedialog, ttk
import os
import sys

# --- Helper function to get resource path for PyInstaller compatibility ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Helper function to center windows ---
def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

# --- NEW: A base class to handle common dialog features ---
class BaseDialog(tk.Toplevel):
    def __init__(self, parent, title, icon_path="app_icon.ico"):
        super().__init__(parent)
        self.title(title)
        self.result = None # Use None to indicate the window was closed

        try:
            self.iconbitmap(resource_path(icon_path))
        except (tk.TclError, FileNotFoundError):
            pass # Ignore if icon not found

        self.resizable(False, False)
        self.grab_set()
        # --- FIX: Handle the window close button ("X") ---
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self.result = None # Set a specific value for closing
        self.destroy()

# --- Custom Dialog for Info/Error Messages ---
class CustomDialog(BaseDialog):
    def __init__(self, parent, title, message, dialog_type="info", icon_path="app_icon.ico"):
        super().__init__(parent, title, icon_path)
        self.title(title)

        try:
            self.iconbitmap(icon_path)
        except (tk.TclError, FileNotFoundError):
            print(f"Warning: Icon file '{icon_path}' not found.")

        # self.geometry("380x150")
        self.resizable(False, False)
        self.grab_set()

        # --- Layout ---
        main_frame = tk.Frame(self)
        main_frame.pack(padx=20, pady=15, expand=True, fill="both")

        text_frame = tk.Frame(main_frame)
        text_frame.pack(side="left", fill="both", expand=True)

        title_font_color = "green" if dialog_type == "success" else "red"
        title_label = tk.Label(text_frame, text=title.upper(), font=("DejaVuSans", 10, "bold"), fg=title_font_color)
        title_label.pack(anchor="w")
        
        message_label = tk.Label(text_frame, text=message, wraplength=300, justify="left")
        message_label.pack(anchor="w", pady=(5,0))

        button_frame = tk.Frame(self)
        button_frame.pack(pady=(0, 15))

        ok_button = ttk.Button(button_frame, text="OK", style="Accent.TButton", command=self.destroy)
        ok_button.pack()
        ok_button.focus_set()
        print("Binding Enter key to OK button")
        
        self.bind('<Return>', lambda event: ok_button.invoke())
        center_window(self)
        self.wait_window(self)

# --- Custom Dialog for Yes/No Questions ---
class CustomYesNoDialog(BaseDialog):
    def __init__(self, parent, title, message, icon_path="app_icon.ico"):
        super().__init__(parent, title, icon_path)
        self.title(title)
        self.result = False # Default to No

        try:
            self.iconbitmap(resource_path(icon_path))
        except tk.TclError:
            print(f"Warning: Icon file '{icon_path}' not found.")
        
        # self.geometry("380x150")
        self.resizable(False, False)
        self.grab_set()

        # ... (Layout is similar to CustomDialog but with two buttons) ...
        main_frame = tk.Frame(self); main_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        text_frame = tk.Frame(main_frame)
        text_frame.pack(side="left")

        title_label = tk.Label(text_frame, text=title, font=("DejaVuSans", 12, "bold"))
        title_label.pack(anchor="w")
        
        message_label = tk.Label(text_frame, text=message, wraplength=300, justify="left")
        message_label.pack(anchor="w", pady=(5,0))

        button_frame = tk.Frame(self); button_frame.pack(pady=(10, 15))
        
        def on_yes():
            self.result = True
            self.destroy()
        def on_no():
            self.result = False
            self.destroy()

        yes_button = ttk.Button(button_frame, text="Yes", style="Accent.TButton", command=on_yes)
        yes_button.pack(side="left", padx=10)
        yes_button.focus_set()
        
        no_button = ttk.Button(button_frame, text="No", command=on_no)
        no_button.pack(side="left", padx=10)

        def on_enter_key(event):
            # Find which widget has focus
            focused_widget = self.focus_get()
            # If the focused widget is a button, "click" it
            if isinstance(focused_widget, ttk.Button):
                focused_widget.invoke()

        self.bind('<Return>', on_enter_key)
        center_window(self)
        self.wait_window(self)

def get_input_file(root):
    """Opens a file dialog to select an Excel file and returns its path."""
    
    filepath = filedialog.askopenfilename(
        parent=root,
        title="Please select the monthly employee Excel file",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    return filepath

def show_success(root, message):
    CustomDialog(root, title="Success", message=message, dialog_type="success")

def show_error(root, message):
    CustomDialog(root, title="Error", message=message, dialog_type="error")

class CustomLoginDialog(BaseDialog):
    def __init__(self, parent, title="Manager Login", icon_path="app_icon.ico"):
        super().__init__(parent, title, icon_path)
        self.geometry("300x160")
        
        def on_submit():
            email = email_entry.get()
            password = password_entry.get()
            if email and password:
                self.result = {'email': email, 'password': password}
                self.destroy()
        
        # ... (Layout of labels, entries, and button is the same) ...
        tk.Label(self, text="Your Email Address:").pack(pady=(10, 0))
        email_entry = tk.Entry(self, width=40); email_entry.pack(); email_entry.focus_set()
        tk.Label(self, text="Your Email Password/App Password:").pack(pady=(10, 0))
        password_entry = tk.Entry(self, show="*", width=40); password_entry.pack()
        submit_button = ttk.Button(self, text="Login & Continue", command=on_submit); submit_button.pack(pady=15)
        
        self.bind('<Return>', lambda event: on_submit())
        center_window(self)
        self.wait_window(self)

def ask_to_use_saved_credentials(root, email):
    dialog = CustomYesNoDialog(root, title="Automatic Login", message=f"The application has found saved credentials for:\n\nEmail : {email}\n\nContinue with this account?")
    print(f"User chose to use saved credentials: {dialog.result}")
    return dialog.result

def ask_to_continue_with_defaults(root, message):
    dialog = CustomYesNoDialog(root, title="Incomplete Data Found", message=message)
    return dialog.result

def get_credentials(root):
    dialog = CustomLoginDialog(root)
    return dialog.result