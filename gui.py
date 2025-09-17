# gui.py
import tkinter as tk
from tkinter import filedialog, messagebox

def get_input_file():
    """Opens a file dialog to select an Excel file and returns its path."""
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename(
        title="Please select the monthly employee Excel file",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    return filepath

def show_success(message):
    """Displays a success message pop-up."""
    messagebox.showinfo("Success", message)

def show_error(message):
    """Displays an error message pop-up."""
    messagebox.showerror("Error", message)

def get_credentials():
    """
    Opens a dialog to securely get the sender's email and password.
    Returns a dictionary with credentials or None if closed.
    """
    cred_window = tk.Toplevel()
    cred_window.title("Manager Login")
    cred_window.geometry("300x150")
    
    # Simple way to store credentials from the dialog
    credentials = {}

    def on_submit():
        credentials['email'] = email_entry.get()
        credentials['password'] = password_entry.get()
        cred_window.destroy()

    tk.Label(cred_window, text="Your Email Address:").pack(pady=(10, 0))
    email_entry = tk.Entry(cred_window, width=40)
    email_entry.pack()

    tk.Label(cred_window, text="Your Email Password/App Password:").pack(pady=(10, 0))
    password_entry = tk.Entry(cred_window, show="*", width=40)
    password_entry.pack()

    submit_button = tk.Button(cred_window, text="Login & Continue", command=on_submit)
    submit_button.pack(pady=10)

    # This makes the main script wait until the credential window is closed
    cred_window.wait_window()
    
    return credentials if 'email' in credentials and credentials['email'] else None

def ask_to_use_saved_credentials(email):
    """
    Shows a 'Yes/No' dialog asking if the user wants to proceed with the
    found credentials or log in again.
    Returns True for 'Yes', False for 'No'.
    """
    title = "Automatic Login"
    message = (
        f"The application has found saved credentials for:\n\n"
        f"Email: {email}\n\n"
        "Would you like to continue with this account?"
    )
    # askyesno returns True for Yes, False for No
    return messagebox.askyesno(title, message)