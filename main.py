# main.py
import tkinter as tk
import pandas as pd
import keyring
import gui
import data_handler
import calculations
import pdf_generator
import email_sender

SERVICE_NAME = "PayslipApp"

def run_payslip_process(root):
    """The main process for the payslip generator application."""
    try:
        pdf_generator.register_fonts()
    except Exception as e:
        gui.show_error(root,f"Could not load font files. Please ensure the 'fonts' folder is correct.\n\nError: {e}")
        return

    # --- Credential Handling ---
    email = keyring.get_password(SERVICE_NAME, "user_email")
    password = keyring.get_password(SERVICE_NAME, "user_password")
    
    is_new_login = False
    old_credentials_existed = (email and password)

    if email and password:
        # --- FIX: Explicitly check the result for True, False, or None ---
        saved_cred_choice = gui.ask_to_use_saved_credentials(root, email)
        
        if saved_cred_choice is False: # User clicked "No"
            is_new_login = True
        elif saved_cred_choice is None: # User clicked "X"
            print("Login cancelled by user. Exiting.")
            return
        # If True, we do nothing and proceed with saved creds
    else:
        is_new_login = True

    if is_new_login:
        credentials_input = gui.get_credentials(root)
        if not credentials_input: # Handles both empty dict and None from close button
            print("Login cancelled. Exiting.")
            return
        email = credentials_input['email']
        password = credentials_input['password']

    # --- File Selection and Validation ---
    input_file = gui.get_input_file(root)
    if not input_file:
        print("No file selected.")
        return

    employee_df = data_handler.load_employee_data(input_file)
    if employee_df is None:
        gui.show_error(root,"Failed to load or read the Excel file. Please check the file format.")
        return

    # 1. Check for missing essential columns
    essential_fields = ['Employee_Email', 'Gross_Salary', 'Income_Tax']
    missing_essential = [field for field in essential_fields if field not in employee_df.columns]
    if missing_essential:
        gui.show_error(root,f"Process stopped. Essential columns are missing: {', '.join(missing_essential)}")
        return

    # 2. Check for missing optional columns and fill with defaults
    # ... (code for handling optional columns) ...

    # 3. Pre-scan for empty/invalid rows in essential fields
    invalid_rows = employee_df[essential_fields].isnull().any(axis=1) | (employee_df['Gross_Salary'] == 0)
    invalid_employee_df = employee_df[invalid_rows]
    if not invalid_employee_df.empty:
        names = "\n - ".join(invalid_employee_df['Employee_Name'].tolist())
        message = f"The following employees have missing essential data:\n\n - {names}\n\nContinue with only the valid employees?"
        if not gui.ask_to_continue_with_defaults(root, message):
            gui.show_error(root,"Process stopped by user. Input the correct data and try again.")
            return

    valid_employee_df = employee_df.drop(invalid_employee_df.index)
    if valid_employee_df.empty:
        gui.show_error(root,"No valid employee records found to process.")
        return

    server = None
    try:
        server = email_sender.connect_to_server(email, password)
        if not server:
            gui.show_error(root,"Could not connect to the email server. Please check credentials and network.")
            return

        if is_new_login:
            print("Login successful. Saving new credentials...")
            if old_credentials_existed:
                keyring.delete_password(SERVICE_NAME, "user_email")
                keyring.delete_password(SERVICE_NAME, "user_password")
            keyring.set_password(SERVICE_NAME, "user_email", email)
            keyring.set_password(SERVICE_NAME, "user_password", password)
            print("Credentials saved.")

        output_dir = None
        for index, employee_row in valid_employee_df.iterrows():
            salary_details = calculations.calculate_salary(employee_row)
            pdf_path, output_dir = pdf_generator.create_payslip(employee_row, salary_details)
            if pdf_path:
                recipient_email = employee_row.get('Employee_Email')
                if recipient_email and recipient_email != 'N/A':
                    email_sender.send_single_email(server, email, recipient_email, employee_row['Employee_Name'], employee_row['Period'], pdf_path)
        
        if output_dir:
            gui.show_success(root,f"All valid payslips generated successfully!\n\nCheck the '{output_dir.resolve()}' folder.")
        else:
            gui.show_success(root,"Process complete!")
    
    except Exception as e:
        gui.show_error(root,f"An unexpected error occurred: {e}")
    
    finally:
        if server:
            email_sender.close_connection(server)

if __name__ == "__main__":
    app_root = tk.Tk()
    app_root.withdraw()
    run_payslip_process(app_root)
    app_root.destroy()
    # run_payslip_process()
    print("Application finished.")