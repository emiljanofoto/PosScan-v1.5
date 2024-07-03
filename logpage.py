import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import openpyxl
import subprocess


def load_data():
    path = "/Users/emiliofoto/Documents/PosScan/logs.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    data = {str(row[0]): row[1] for row in sheet.iter_rows(min_row=2, values_only=True)}
    return data

# Function to handle login
def login(event=None):
    user_id = id_entry.get()
    if user_id in data:
        name_entry.config(state='normal')
        name_entry.delete(0, tk.END)
        name_entry.insert(0, data[user_id])
        name_entry.config(state='readonly')
        show_message("Hyrja me sukses!", "green")
        root.after(2000, open_main_form)
    else:
        show_message("ID nuk ekziston!", "red")
        name_entry.config(state='normal')
        name_entry.delete(0, tk.END)
        name_entry.config(state='readonly')

# Function to open the main form
def open_main_form():
    subprocess.Popen(['python', 'main.py'])
    root.destroy()  # Close the login window

# Function to open the users form
def open_users_form():
    # Window for password entry
    password_window = tk.Toplevel(root)
    password_window.title("Verifikim")
    password_window.geometry("300x150")

    # Password label and entry
    ttk.Label(password_window, text="Password:").pack(pady=10)
    password_entry = ttk.Entry(password_window, show="*")
    password_entry.pack(pady=5)

    # Function to check the password
    def check_password():
        if password_entry.get() == "12345":
            subprocess.Popen(['python', 'users.py'])
            root.destroy()  # Close the login window
            password_window.destroy()  # Close the password window
        else:
            error_label.config(text="Password i gabuar", foreground="red")

    # Submit button
    submit_button = ttk.Button(password_window, text="Hyr", command=check_password)
    submit_button.pack(pady=10)

    # Error message label
    error_label = ttk.Label(password_window, text="")
    error_label.pack(pady=5)

# Function to show a message
def show_message(message, color):
    message_label.config(text=message, bg=color, fg="white")

# Main application window
root = tk.Tk()
root.title("Hyrja ne Aplikacion")
root.geometry("400x450")

# Load data from Excel
data = load_data()

# Add button to open users.py and close current form
users_button = ttk.Button(root, text="Shto Perdorues", command=open_users_form)
users_button.pack(pady=20)

# Load and display the logo
logo_path = "/Users/emiliofoto/Documents/PosScan/logo.png"
logo_image = Image.open(logo_path)
logo_photo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(root, image=logo_photo)
logo_label.pack(pady=5)

# ID entry
ttk.Label(root, text="ID:").pack(pady=10)
id_entry = ttk.Entry(root, font=("Helvetica", 25))
id_entry.pack(pady=5)
id_entry.bind("<Return>", login)  # Bind Enter key to login function

# Name entry (readonly)
ttk.Label(root, text="Emri:").pack(pady=10)
name_entry = ttk.Entry(root, font=("Helvetica", 25), state='readonly')
name_entry.pack(pady=5)

# Message label
message_label = tk.Label(root, text="", font=("Helvetica", 12), height=2)
message_label.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

# Hide buttons (set visibility to hidden)
login_button = ttk.Button(root, text="HYR", command=login)
login_button.pack(pady=20)
login_button.pack_forget()

proceed_button = ttk.Button(root, text="Menu kryesore", command=open_main_form)
proceed_button.pack(pady=10)
proceed_button.pack_forget()

# Set focus to ID entry when the form opens
root.after(0, lambda: id_entry.focus_set())

root.mainloop()
