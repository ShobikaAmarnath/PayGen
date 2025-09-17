# main.py
import gui
import data_handler
import calculations
import pdf_generator
import email_sender
import keyring

SERVICE_NAME = "PayslipApp"

def run_payslip_process():
    """The main process for the payslip generator application."""
    try:
        pdf_generator.register_fonts()
    except Exception as e:
        gui.show_error(f"Could not load font files. Please ensure the 'fonts' folder is correct.\n\nError: {e}")
        return
    
    # Try to get the saved email and password from the keyring
    email = keyring.get_password(SERVICE_NAME, "user_email")
    password = keyring.get_password(SERVICE_NAME, "user_password")

    use_saved_credentials = False
    if email and password:
        # If credentials are found, ask the user if they want to use them.
        if gui.ask_to_use_saved_credentials(email):
            use_saved_credentials = True
            print(f"User confirmed. Proceeding with saved credentials for {email}.")
        else:
            # User wants to log in with a different account.
            print("User chose to log in with a different account.")
    
    # If we are NOT using saved credentials, show the login window.
    if not use_saved_credentials:
        credentials_input = gui.get_credentials()
        if not credentials_input:
            print("Login cancelled. Exiting.")
            return
        email = credentials_input['email']
        password = credentials_input['password']
        # Set a flag to save these new credentials after a successful send.
        save_credentials_on_success = True
    else:
        save_credentials_on_success = False

    input_file = gui.get_input_file()
    if not input_file:
        print("No file selected. Exiting.")
        return

    employee_df = data_handler.load_employee_data(input_file)
    if employee_df is None:
        gui.show_error("Failed to load or read the Excel file. Please check the file format.")
        return
    
    excel_columns = employee_df.columns.tolist()

    # 1. Check for ESSENTIAL fields. Stop if any are missing.
    essential_fields = ['Employee_Email', 'Gross_Salary', 'Income_Tax']
    missing_essential = [field for field in essential_fields if field not in excel_columns]
    
    if missing_essential:
        missing_fields_str = "\n - ".join(missing_essential)
        gui.show_error(
            "Process stopped. The following essential columns are missing from the Excel file:\n\n"
            f" - {missing_fields_str}"
        )
        return # Exit the application

    # 2. Check for OPTIONAL fields. Ask to continue if any are missing.
    optional_fields = {
        'Location': 'N/A', 'Date_of_Joining': 'N/A', 'Days_Worked': 'N/A',
        'Bank_Name': 'N/A', 'Bank_Account_No': 'N/A', 'PAN_Number': 'N/A',
        'PF_Account_Number': 'N/A', 'ESI_Number': 'N/A', 'UAN_Number': 'N/A',
        'LOP_Days': 'N/A'
    }
    missing_optional = [field for field in optional_fields if field not in excel_columns]

    if missing_optional:
        missing_fields_str = "\n - ".join(missing_optional)
        message = (
            f"The following columns were not found in the Excel file:\n - {missing_fields_str}\n\n"
            "Do you want to continue by using default placeholder values for them?"
        )
        if not gui.messagebox.askyesno("Optional Data Missing", message):
            gui.show_error("Process stopped by user.")
            return

        # If user clicks 'Yes', add the missing columns with default values
        for field in missing_optional:
            employee_df[field] = optional_fields[field]

    # At this point, all essential fields are present, and optional fields are filled with defaults if missing.
    # Find rows that have an empty value in ANY of the essential columns
    invalid_rows = employee_df[
        essential_fields
    ].isnull().any(axis=1) | (employee_df['Gross_Salary'] == 0)

    invalid_employee_df = employee_df[invalid_rows]

    if not invalid_employee_df.empty:
        # Get the names of employees with missing data
        invalid_employee_names = invalid_employee_df['Employee_Name'].tolist()
        names_to_display = "\n - ".join(invalid_employee_names)
        
        message = (
            "The following employees have missing or invalid data in essential "
            "fields (Email, Gross Salary, or Income Tax):\n\n"
            f" - {names_to_display}\n\n"
            "Do you want to continue and process payslips only for the valid employees?"
        )
        
        if not gui.messagebox.askyesno("Incomplete Data Found", message):
            gui.show_error("Process stopped by user. Please correct the Excel file.")
            return

    # --- Create a new DataFrame containing only the valid rows ---
    valid_employee_df = employee_df.drop(invalid_employee_df.index)

    if valid_employee_df.empty:
        gui.show_error("No valid employee records found to process in the selected file.")
        return

    server = None
    try:
        output_dir = None
        # Connect and log in
        server = email_sender.connect_to_server(email, password)
        if not server:
            gui.show_error("Could not connect to the email server. Please check credentials and network.")
            return
        
        for index, employee_row in valid_employee_df.iterrows():            
            salary_details = calculations.calculate_salary(employee_row)
            pdf_path, output_dir = pdf_generator.create_payslip(employee_row, salary_details)
            if not pdf_path:
                print(f"Skipping employee {employee_row['Employee_Name']} due to PDF generation failure.")
                continue

            if save_credentials_on_success:
                # Delete any old credentials before setting new ones
                keyring.delete_password(SERVICE_NAME, "user_email")
                keyring.delete_password(SERVICE_NAME, "user_password")
                # Save the new credentials
                keyring.set_password(SERVICE_NAME, "user_email", email)
                keyring.set_password(SERVICE_NAME, "user_password", password)
                print("New credentials saved successfully for future use.")
                save_credentials_on_success = False # Important: Only save once!

            recipient_email = employee_row.get('Employee_Email')
            print(f"Preparing to send email to: {recipient_email}")
            
            if recipient_email and recipient_email != 'N/A':                
                email_sender.send_single_email(
                    server=server,
                    sender_email=email,
                    recipient_email=recipient_email,
                    employee_name=employee_row['Employee_Name'],
                    period=employee_row['Period'],
                    pdf_path=pdf_path
                )
            else:
                print(f"Skipping email for {employee_row['Employee_Name']} due to missing email address.")
        
        if output_dir:
            gui.show_success(f"All payslips have been generated successfully!\n\nCheck the '{output_dir.resolve()}' folder.")
        else:
            gui.show_success("Process complete!")
    
    except Exception as e:
        gui.show_error(f"An unexpected error occurred during PDF generation:\n\n{e}")

    finally:
        # --- Ensure the connection is always closed at the end ---
        if server:
            email_sender.close_connection(server)

if __name__ == "__main__":
    run_payslip_process()