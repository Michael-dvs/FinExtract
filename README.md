# FinExtract

**FinExtract** is a robust desktop application designed to automate the extraction of financial transaction data from PDF bank statements (*Rekening Koran*) into structured, formatted Excel files. Built with Python and a modern GUI using CustomTkinter, it simplifies the process of converting unstructured PDF data into usable spreadsheets for accounting and analysis.

## 🚀 Features

-   **Multi-Bank Support**: Specialized extraction logic for major Indonesian banks:
    -   **BNI** (Bank Negara Indonesia)
    -   **Mandiri** (Regular Statements)
    -   **Livin' by Mandiri** (Mobile App Statements)
    -   **OCBC**
    -   **BRI** (Bank Rakyat Indonesia)
-   **Modern GUI**: A clean, user-friendly interface built with `customtkinter`, featuring:
    -   Dark and Light mode support.
    -   Customizable color themes (Default Blue, Special Edition Pink).
    -   Real-time process logging.
-   **Smart Extraction**:
    -   Handles multi-line transaction descriptions.
    -   Intelligent column mapping (Debit/Credit detection).
    -   Automatic date formatting.
-   **Security**:
    -   Detects encrypted/password-protected PDFs.
    -   Securely prompts the user for passwords only when necessary.


## 📋 Prerequisites

Ensure you have **Python 3.8+** installed on your system. You will need the following Python libraries:

-   `customtkinter` (UI Framework)
-   `pdfplumber` (PDF Extraction)
-   `pandas` (Data Manipulation)
-   `pypdf` (PDF Decryption/Manipulation)
-   `xlsxwriter` (Excel Formatting)

## 🛠️ Installation

1.  **Clone the repository** or download the source code to your local machine.
    ```bash
    git clone https://github.com/Michael-dvs/FinExtract.git
    cd FinExtract
    ```

2.  **Install dependencies**:
    It is recommended to use a virtual environment.
    ```bash
    pip install customtkinter pdfplumber pandas pypdf xlsxwriter
    ```

## 💻 Usage

1.  **Run the Application**:
    Execute the main script to launch the GUI.
    ```bash
    python main.py
    ```

2.  **Select Input Files**:
    -   Click **"Browse..."** next to "Input PDF File(s)" to select one or multiple PDF bank statements.

3.  **Select Output Folder**:
    -   Choose the destination folder where the resulting Excel files will be saved.

4.  **Choose Bank**:
    -   Click the button corresponding to the bank statement you are processing (e.g., **BNI**, **Mandiri**, **OCBC**, etc.) from the sidebar.

5.  **Process**:
    -   The application will process the files.
    -   **Password Protected Files**: If a PDF is encrypted, a popup dialog will ask for the password.
    -   Check the **"Log Proses"** panel for status updates.

6.  **Finish**:
    -   Once complete, you can use the **"Buka File Output"** or **"Buka Folder Output"** buttons to view your data.

## Configuration

The application creates a configuration file named `FinExtract_Settings.json` in your user home directory (e.g., `C:\Users\Name\` or `/Users/Name/`).

-   **Themes**: Stores your preferred appearance mode (System/Dark/Light) and color theme.
-   **Bank Configs**: Advanced users can modify this JSON to customize Excel column visibility, labels, and widths for specific banks (handled via `utils.py`).

## 📂 Project Structure

-   `main.py`: The entry point of the application. Contains the GUI logic and thread management.
-   `utils.py`: Utility functions for loading configurations and saving styled Excel files.
-   `BNI.py`: Extraction logic specific to BNI statements.
-   `Mandiri.py`: Extraction logic specific to Mandiri statements.
-   `Livin.py`: Extraction logic for Livin' by Mandiri app exports.
-   `OCBC.py`: Extraction logic specific to OCBC statements.
-   `BRI.py`: Extraction logic specific to BRI statements.

## ⚠️ Disclaimer

This tool is intended for personal productivity and data organization purposes.
-   **Accuracy**: While every effort has been made to ensure accurate extraction, PDF formats can change. Always verify the output Excel against the original PDF.
-   **Security**: Passwords entered in the application are used solely in memory for decryption during the runtime session and are never stored or transmitted.

## 👤 Author

**Michael Aristyo R.**
*FinExtract v1.0.8*

---

> **Note**: If you encounter "Permission Error" when saving files, ensure the target Excel file is not currently open in another program.