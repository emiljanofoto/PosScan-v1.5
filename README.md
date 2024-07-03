PosSCAN v1.5

Description
PosSCAN is a Tkinter-based desktop application for scanning and managing records. The application allows users to scan and validate entries, save records, and view logs. Originally designed to use Excel for data storage, the application has been updated to support Microsoft SQL Server for data management.

Features
Scan and validate entries with predefined formats.
Automatically look up and fill additional fields based on scanned data.
Display the last 20 records in a table.
Validate entries to ensure uniqueness.
Provide real-time data and time updates.
Open secondary forms for additional functionalities.
Save records to and retrieve records from a Microsoft SQL Server database.
Prerequisites
Python 3.6 or later
Microsoft SQL Server
Required Python packages: tkinter, openpyxl, Pillow, pyodbc
Installation
Clone the repository:

git clone https://github.com/yourusername/posscan.git
cd posscan

Install required packages:
pip install -r requirements.txt

Configure SQL Server:

Set up a Microsoft SQL Server instance.
Create a database and a table named people with appropriate columns matching the fields used in the application (e.g., Pozicion, Sasi, Harness, PO, Adresa, Etiketa, DataOra).
Update database connection settings:

Modify the DATABASE_CONFIG in your main script (main.py) to match your SQL Server settings.
Usage
Run the application:
python main.py

Interact with the GUI:

Etiketa Field: Enter the etiketa value in the format ----\-----\----. Validation occurs when the field loses focus.
Adresa Field: Enter the address value starting with * or Adresa. Validation occurs when the field loses focus.
Save Records: Click the "RUAJ TE DHENAT" button to save valid records to the database.
Open Log Page: Click the "LOG PAGE" button to open the log page (secondary form).
Open Secondary Form: Click the "HAP SKANIMET" button to open the secondary form.
Modifications for SQL Server Integration
The original application used Excel for data storage. The following changes were made to integrate with Microsoft SQL Server:

Install pyodbc package:
pip install pyodbc

Additional Information
For further details or troubleshooting, refer to the Tkinter documentation, pyodbc documentation, and Microsoft SQL Server documentation.

Contact
For any questions or support, please contact emiljanofoto@gmail.com

the software is under contstruction!!!






# PosScan
