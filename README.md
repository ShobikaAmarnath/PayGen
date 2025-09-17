# Automated PDF Payslip Generator & Emailer

A Python desktop application that automates the entire payroll process: it reads employee data from an Excel file, calculates salary components, generates professional PDF payslips, and securely emails them to each employee.


## Features

- **GUI Interface**: A simple and user-friendly graphical interface for selecting files and logging in.
- **Dynamic Excel Processing**: Intelligently reads a monthly Excel file, with flexible column name mapping.
- **Automated Calculations**: Automatically calculates all earnings (Basic, HRA, Special Allowance) and deductions (EPF, Prof. Tax) from a single "Gross Salary" input.
- **PDF Generation**: Creates clean, professional, and consistently formatted PDF payslips for each employee.
- **Secure Emailing**:
    - Sends payslips directly to employees' email addresses.
    - Securely saves the sender's login credentials using the operating system's native credential manager (`keyring`).
    - Auto-detects the email provider (Gmail/Microsoft 365) to use the correct SMTP server.
- **Robust Validation**:
    - Checks for missing essential columns and data rows before processing.
    - Notifies the user of any issues and allows them to decide whether to proceed.
- **Packaged Application**: Can be easily bundled into a standalone `.exe` file for distribution using PyInstaller.

## Project Structure

The project is organized into several modules for clarity and maintainability:

- `main.py`: The main entry point and orchestrator for the application.
- `gui.py`: Manages all graphical user interface elements (file dialogs, login windows, messages).
- `data_handler.py`: Handles reading and validating the input Excel file.
- `calculations.py`: Contains all the business logic for salary calculations.
- `pdf_generator.py`: Responsible for creating the PDF documents using ReportLab.
- `email_sender.py`: Manages connecting to SMTP servers and sending emails.
- `config.py`: Stores static configuration like company details and constants.
- `column_config.py`: Maps flexible Excel column names to internal code names.

## Setup and Installation

1.  **Clone the repository:**
    ```bash
    git clone <your-repo-url>
    cd EMPPAYSLIP
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    # For Windows
    python -m venv .venv
    .venv\Scripts\activate
    ```

3.  **Install the required libraries:**
    ```bash
    pip install -r requirements.txt
    ```

## How to Use

1.  **Prepare your Excel file** with the employee and monthly salary data.
2.  **Run the application** from the command line:
    ```bash
    python main.py
    ```
3.  **Follow the on-screen prompts**:
    - Log in with your email credentials (only required the first time).
    - Select the prepared Excel file.
    - Confirm any data validation messages.

The application will then generate the payslips in a new folder (e.g., `Payslips_Apr_2025`) and email them.

## Building the Executable

To create a standalone `.exe` file for easy distribution:

1.  Ensure PyInstaller is installed (`pip install pyinstaller`).
2.  Run the following command from the project's root directory:
    ```bash
    pyinstaller --onefile --windowed --noconsole --add-data "company_logo.png:." --add-data "fonts:fonts" main.py
    ```
3.  The final application will be located in the `dist` folder.