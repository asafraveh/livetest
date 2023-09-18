import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook
from tkcalendar import Calendar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading


# Function to add data to Excel
def add_data_to_excel():
    customer = customer_combobox.get()
    site = site_combobox.get()
    date = date_cal.get_date()
    net = net_entry.get()
    version = version_entry.get()
    total_event = total_event_entry.get()
    false_event = false_event_entry.get()
    version_bugs = version_bugs_entry.get()
    unique_id = unique_id_entry.get()
    excel_file_path = excel_file_path_entry.get()

    try:
        workbook = load_workbook(excel_file_path)
        if customer in workbook.sheetnames:
            sheet = workbook[customer]
        else:
            sheet = workbook.create_sheet(customer)

        row = [date, site, net, version, total_event, false_event, version_bugs, unique_id]
        sheet.append(row)

        # Add the data to the "follow up" sheet
        follow_up_sheet = workbook["follow up"]

        # Check if the date already exists in the "follow up" sheet
        date_column = follow_up_sheet["A"]
        dates = [cell.value for cell in date_column[1:]]  # Skip the header
        if date in dates:
            row_index = dates.index(date) + 2  # Add 2 to skip the header and convert to 1-based index
        else:
            follow_up_sheet.insert_rows(2)  # Insert a new row at row 2 (adjust as needed)
            follow_up_sheet.cell(row=2, column=1, value=date)
            row_index = 2

        # Mark the site column with a "V"
        site_index = sites.index(site) + 2  # Add 2 to skip the header and convert to 1-based index
        follow_up_sheet.cell(row=row_index, column=site_index, value="V")

        workbook.save(excel_file_path)
        result_label.config(text="Data added to Excel successfully.")
    except Exception as e:
        result_label.config(text=f"Error adding data to Excel: {str(e)}")


# Function to clear all fields
def clear_fields():
    # Clear all entry fields
    net_entry.delete(0, tk.END)
    version_entry.delete(0, tk.END)
    total_event_entry.delete(0, tk.END)
    false_event_entry.delete(0, tk.END)
    version_bugs_entry.delete(0, tk.END)
    unique_id_entry.delete(0, tk.END)
    excel_file_path_entry.delete(0, tk.END)
    excel_file_path_entry.insert(0, "O:\\QA\\live.xlsx")  # Reset Excel file path


# Function to open websites and perform login on a separate thread
def open_website_and_login_threaded(website_url, username_value, password_value):
    # Create a new Chrome browser instance
    driver = webdriver.Chrome()

    try:
        # Open the specified website
        driver.get(website_url)

        # Use explicit wait to wait for the username field to be clickable
        wait = WebDriverWait(driver, 10)
        username_field = wait.until(EC.element_to_be_clickable((By.ID, "username")))

        # Click the username field to focus it
        username_field.click()

        # Send keys to the username field
        username_field.send_keys(username_value)

        # Similarly, handle the password field and login button

        password_field = wait.until(EC.element_to_be_clickable((By.ID, "password")))
        password_field.click()
        password_field.send_keys(password_value)

        login_button = wait.until(EC.element_to_be_clickable((By.ID, "kc-login")))
        login_button.click()

        # Wait for user confirmation before closing the browser
        input("Press Enter to close the browser...")
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        driver.quit()  # Close the browser in all cases


# Function to open the CEMEX website and perform login on a separate thread
def open_cemex_website_threaded():
    threading.Thread(target=open_website_and_login_threaded, args=(
        "https://cemex-manager-app.ception.live/",
        "ceptionuser",
        "ceptioncemex?"
    )).start()


# Function to open the Shafir website and perform login on a separate thread
def open_shafir_website_threaded():
    threading.Thread(target=open_website_and_login_threaded, args=(
        "https://shapir-manager-app.ception.live/",
        "ceptionuser",
        "ceptionshapir!"
    )).start()


# Function to open the Heidelberg website and perform login on a separate thread
def open_heidelberg_website_threaded():
    threading.Thread(target=open_website_and_login_threaded, args=(
        "https://heidelberg-manager-app.ception.live/",
        "ceptionuser",
        "ceptionheidelberg%"
    )).start()


# Create the main window
root = tk.Tk()
root.title("Excel Data Entry and Website Automation")

# Create and configure the frame
frame = ttk.Frame(root, padding=10)
frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))
frame.columnconfigure(1, weight=1)

# Customer selection
customer_label = ttk.Label(frame, text="Customer:")
customer_label.grid(column=0, row=0, sticky=tk.W)
customer_combobox = ttk.Combobox(frame, values=["cemex", "Shafir", "hidenberg"])
customer_combobox.grid(column=1, row=0, sticky=tk.W)

# Site selection
site_label = ttk.Label(frame, text="Site:")
site_label.grid(column=0, row=1, sticky=tk.W)
sites = ["adhalom", "Gdansk", "Gdynia", "golani", "mevocarmel", "modiim", "yvnehe", "roller", "Etziyona", "London1627",
         "London1656"]
site_combobox = ttk.Combobox(frame, values=[])
site_combobox.grid(column=1, row=1, sticky=tk.W)
# Update site choices based on customer selection
def update_site_choices(event):
    selected_customer = customer_combobox.get()

    # Define site choices based on the selected customer
    if selected_customer == "Shafir":
        site_combobox['values'] = ["roller", "Etziyona"]
    elif selected_customer == "cemex":
        site_combobox['values'] = ["adhalom", "Gdansk", "Gdynia", "golani", "mevocarmel", "modiim", "yvnehe"]
    elif selected_customer == "hidenberg":
        site_combobox['values'] = ["London1627", "London1656"]
    else:
        site_combobox['values'] = []


# Bind the update_site_choices function to the Combobox selection event
customer_combobox.bind("<<ComboboxSelected>>", update_site_choices)

# Date entry
date_label = ttk.Label(frame, text="Date:")
date_label.grid(column=0, row=2, sticky=tk.W)
date_cal = Calendar(frame)
date_cal.grid(column=1, row=2, sticky=tk.W)

# Other fields (net, version, total event, false event, version bugs, unique id)
net_label = ttk.Label(frame, text="Net:")
net_label.grid(column=0, row=3, sticky=tk.W)
net_entry = ttk.Entry(frame)
net_entry.grid(column=1, row=3, sticky=tk.W)

version_label = ttk.Label(frame, text="Version:")
version_label.grid(column=0, row=4, sticky=tk.W)
version_entry = ttk.Entry(frame)
version_entry.grid(column=1, row=4, sticky=tk.W)

total_event_label = ttk.Label(frame, text="Total Event:")
total_event_label.grid(column=0, row=5, sticky=tk.W)
total_event_entry = ttk.Entry(frame)
total_event_entry.grid(column=1, row=5, sticky=tk.W)

false_event_label = ttk.Label(frame, text="False Event:")
false_event_label.grid(column=0, row=6, sticky=tk.W)
false_event_entry = ttk.Entry(frame)
false_event_entry.grid(column=1, row=6, sticky=tk.W)

version_bugs_label = ttk.Label(frame, text="Version Bugs:")
version_bugs_label.grid(column=0, row=7, sticky=tk.W)
version_bugs_entry = ttk.Entry(frame)
version_bugs_entry.grid(column=1, row=7, sticky=tk.W)

unique_id_label = ttk.Label(frame, text="Unique ID:")
unique_id_label.grid(column=0, row=8, sticky=tk.W)
unique_id_entry = ttk.Entry(frame)
unique_id_entry.grid(column=1, row=8, sticky=tk.W)

# Excel file path entry
excel_file_path_label = ttk.Label(frame, text="Excel File Path:")
excel_file_path_label.grid(column=0, row=9, sticky=tk.W)
excel_file_path_entry = ttk.Entry(frame)
excel_file_path_entry.grid(column=1, row=9, sticky=(tk.W, tk.E))
excel_file_path_entry.insert(0, "O:\\QA\\live.xlsx")

# Add data button
add_data_button = ttk.Button(frame, text="Add to Excel", command=add_data_to_excel)
add_data_button.grid(column=1, row=10, sticky=tk.E)

# Clear button
clear_button = ttk.Button(frame, text="Clear", command=clear_fields)
clear_button.grid(column=0, row=10, sticky=tk.W)

# Result label
result_label = ttk.Label(frame, text="")
result_label.grid(column=0, row=11, columnspan=2)

# Create buttons for website automation
cemex_button = ttk.Button(frame, text="CEMEX Website", command=open_cemex_website_threaded)
shafir_button = ttk.Button(frame, text="Shafir Website", command=open_shafir_website_threaded)
heidelberg_button = ttk.Button(frame, text="Heidelberg Website", command=open_heidelberg_website_threaded)

# Place website automation buttons in the frame
cemex_button.grid(column=0, row=12, sticky=tk.W)
shafir_button.grid(column=1, row=12, sticky=tk.W)
heidelberg_button.grid(column=2, row=12, sticky=tk.W)

root.mainloop()
