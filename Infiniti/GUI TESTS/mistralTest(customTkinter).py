import customtkinter as ctk
from tkinter import filedialog, messagebox
import json
import os
import appdirs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import glob
import sys
import threading
import ctypes
import pandas as pd
from datetime import datetime

class InfinitiGUI:
    def __init__(self, master):
        self.master = master
        master.title("Infiniti Automation")
        master.geometry("600x400")
        master.resizable(False, False)

        self.config_file = os.path.join(os.getenv("APPDATA"), "infiniti_config.json")
        self.load_config()

        # Create widgets
        self.create_widgets()

        self.automation_thread = None

    def create_widgets(self):
        # Username
        ctk.CTkLabel(self.master, text="Username:").grid(row=0, column=0, sticky="e", padx=10, pady=5)
        self.username_entry = ctk.CTkEntry(self.master)
        self.username_entry.grid(row=0, column=1, padx=10, pady=5)
        self.username_entry.insert(0, self.config.get("username", ""))

        # Password
        ctk.CTkLabel(self.master, text="Password:").grid(row=1, column=0, sticky="e", padx=10, pady=5)
        self.password_entry = ctk.CTkEntry(self.master, show="*")
        self.password_entry.grid(row=1, column=1, padx=10, pady=5)
        self.password_entry.insert(0, self.config.get("password", ""))

        # Download Folder
        ctk.CTkLabel(self.master, text="Download Folder:").grid(row=2, column=0, sticky="e", padx=10, pady=5)
        self.download_entry = ctk.CTkEntry(self.master)
        self.download_entry.grid(row=2, column=1, padx=10, pady=5)
        self.download_entry.insert(0, self.config.get("download_folder", ""))
        ctk.CTkButton(self.master, text="Browse", command=lambda: self.browse_folder(self.download_entry)).grid(row=2, column=2, padx=10, pady=5)

        # Excel and PDF Files Location
        ctk.CTkLabel(self.master, text="Excel/PDF Files:").grid(row=3, column=0, sticky="e", padx=10, pady=5)
        self.files_entry = ctk.CTkEntry(self.master)
        self.files_entry.grid(row=3, column=1, padx=10, pady=5)
        self.files_entry.insert(0, self.config.get("files_folder", ""))
        ctk.CTkButton(self.master, text="Browse", command=lambda: self.browse_folder(self.files_entry)).grid(row=3, column=2, padx=10, pady=5)

        # Summary Excel File Location
        ctk.CTkLabel(self.master, text="Summary Excel:").grid(row=4, column=0, sticky="e", padx=10, pady=5)
        self.summary_entry = ctk.CTkEntry(self.master)
        self.summary_entry.grid(row=4, column=1, padx=10, pady=5)
        self.summary_entry.insert(0, self.config.get("summary_file", ""))
        ctk.CTkButton(self.master, text="Browse", command=lambda: self.browse_file(self.summary_entry)).grid(row=4, column=2, padx=10, pady=5)

        # Checkboxes
        self.step1_var = ctk.IntVar(value=self.config.get("step1", 0))
        self.step2_var = ctk.IntVar(value=self.config.get("step2", 0))
        self.step3_var = ctk.IntVar(value=self.config.get("step3", 0))

        ctk.CTkCheckBox(self.master, text="Step 1", variable=self.step1_var).grid(row=5, column=0, padx=10, pady=5)
        ctk.CTkCheckBox(self.master, text="Step 2", variable=self.step2_var).grid(row=5, column=1, padx=10, pady=5)
        ctk.CTkCheckBox(self.master, text="Step 3", variable=self.step3_var).grid(row=5, column=2, padx=10, pady=5)

        # Run Button
        ctk.CTkButton(self.master, text="Run", command=self.run_automation).grid(row=6, column=1, padx=10, pady=5)

        # Exit Button
        self.exit_button = ctk.CTkButton(self.master, text="Exit", command=self.exit_application, fg_color="red", hover_color="#ff5757")
        self.exit_button.grid(row=7, column=1, padx=10, pady=5)

    def browse_folder(self, entry):
        folder = filedialog.askdirectory()
        if folder:
            entry.delete(0, ctk.END)
            entry.insert(0, folder)

    def browse_file(self, entry):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file:
            entry.delete(0, ctk.END)
            entry.insert(0, file)

    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                self.config = json.load(f)
        else:
            self.config = {}

    def save_config(self):
        config = {
            "username": self.username_entry.get(),
            "password": self.password_entry.get(),
            "download_folder": self.download_entry.get(),
            "files_folder": self.files_entry.get(),
            "summary_file": self.summary_entry.get(),
            "step1": self.step1_var.get(),
            "step2": self.step2_var.get(),
            "step3": self.step3_var.get(),
        }
        with open(self.config_file, "w") as f:
            json.dump(config, f)

    def exit_application(self):
        if messagebox.askokcancel("Exit", "Are you sure you want to exit?"):
            print("Exiting application...")
            if self.automation_thread and self.automation_thread.is_alive():
                thread_id = self.automation_thread.ident
                ret = ctypes.pythonapi.PyThreadState_SetAsyncExc(
                    ctypes.c_long(thread_id), ctypes.py_object(SystemExit)
                )
                if ret > 1:
                    ctypes.pythonapi.PyThreadState_SetAsyncExc(
                        ctypes.c_long(thread_id), None
                    )
                    print("Exception raise failure")
            self.master.quit()
            self.master.destroy()
            sys.exit()

    def run_automation(self):
        self.automation_thread = threading.Thread(target=self._run_automation)
        self.automation_thread.start()

    def _run_automation(self):
        self.save_config()

        chrome_options = Options()
        chrome_options.add_argument("--ignore-ssl-errors=yes")
        chrome_options.add_argument("--ignore-certificate-errors")

        download_dir = os.path.normpath(self.download_entry.get().strip())
        files_dir = os.path.normpath(self.files_entry.get().strip())
        summary_file = os.path.normpath(self.summary_entry.get().strip())

        print(f"Download Directory: {download_dir}")
        print(f"Files Directory: {files_dir}")
        print(f"Summary File: {summary_file}")

        chrome_options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True,
            },
        )

        driver = webdriver.Chrome(options=chrome_options)

        try:
            username = self.username_entry.get()
            password = self.password_entry.get()
            driver.get("https://infiniti.tataconsumer.com/private/tea/billing")
            login(driver, username, password)

            if self.step1_var.get():
                Step1(driver, files_dir)
            if self.step2_var.get():
                Step2(driver, files_dir, summary_file)
            if self.step3_var.get():
                Step3(driver, download_dir)

            self.master.after(
                0,
                lambda: messagebox.showinfo(
                    "Automation Complete", "The automation has finished running."
                ),
            )

        except Exception as e:
            error_message = str(e)
            self.master.after(
                0,
                lambda: messagebox.showerror(
                    "Error", f"An error occurred: {error_message}"
                ),
            )
        finally:
            driver.quit()

def login(driver, username, password):
    WebDriverWait(driver, 20).until(EC.title_contains("Infiniti"))
    username_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='username']"))
    )
    username_field.send_keys(username)
    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']"))
    )
    password_field.send_keys(password)
    try:
        sign_in_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "#location-tabpanel-0 > div > p > div > div:nth-child(4) > div > div > div > div > form > div > button",
                )
            )
        )
    except:
        sign_in_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
        )
    sign_in_button.click()
    print("Successfully clicked the Sign in button")

def Step1(driver, files_dir):
    click_element(driver, "#dashboard", "Home Button")
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )
    click_element(
        driver,
        "#root > div > div > div > header > div > div.i-primary-nav.MuiBox-root.css-1dzmjrl > button:nth-child(2)",
        "Tea Private Button",
    )
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )
    click_element(
        driver,
        "#root > div > div > div > div > main > div.MuiDrawer-root.MuiDrawer-docked.i-drawer.css-1de6c1k > div > ul > li:nth-child(3) > a > div",
        "Billing Button",
    )
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    for filename in os.listdir(files_dir):
        file_path = os.path.join(files_dir, filename)
        if os.path.isfile(file_path) and (
            filename.endswith(".xlsx") or filename.endswith(".xls")
        ):
            click_element(
                driver,
                "#root > div > div > div > div > main > div.rightside-modules > div > div:nth-child(1) > div > div.MuiGrid-root.MuiGrid-item.MuiGrid-grid-xs-2.MuiGrid-grid-sm-2.css-605f2w > button",
                "Upload AutoBill Button",
            )
            file_input = driver.find_element(
                By.CSS_SELECTOR,
                "body > div.MuiModal-root.css-8ndowl > div.uploadOfferModal.MuiBox-root.css-h1jd1n > div > div:nth-child(1) > div > input[type=file]",
            )
            file_input.send_keys(file_path)
            #submitButton = driver.find_element(By.XPATH,'/html/body/div[4]/div[3]/div/div[2]/div/div/button[2]')
            #submitButton.click()
            #Submit button commented out for testing, cancel button used instead
            time.sleep(5)
            click_element(
                driver,
                "body > div.MuiModal-root.css-8ndowl > div.uploadOfferModal.MuiBox-root.css-h1jd1n > div > div:nth-child(2) > div > div > button.MuiButtonBase-root.MuiButton-root.MuiButton-Primary.MuiButton-sizeMedium.MuiButton-SizeMedium.MuiButton-colorPrimary.MuiButton-root.MuiButton-Primary.MuiButton-sizeMedium.MuiButton-SizeMedium.MuiButton-colorPrimary.cancel-btn.css-z1efht",
                "Cancel Button",
            )

def Step2(driver, files_dir, summary_file):
    print("Starting Step2 function")

    # Load the data from the summary Excel file
    df = pd.read_excel(summary_file, header=1)  # Use the second row as the header

    # Print column types for debugging
    print("DataFrame column types:")
    print(df.dtypes)

    # Print out all values that will be input
    print("\nAll values that will be input:")
    for index, row in df.iterrows():
        print(f"\nRow {index + 1}:")
        print(f"PO NUMBER: {row['PO NUMBER']}")
        print(f"BUYING TYPE: {row['BUYING TYPE']}")
        print(f"BUYING CENTRE: {row['BUYING CENTRE']}")
        print(f"VENDOR CODE: {row['VENDOR CODE']}")
        print(f"INVOICE NUMBER: {row['INVOICE NUMBER']}")
        print(f"INVOICE QUANTITY: {row['INVOICE QUANTITY']}")
        print(f"INVOICE VALUE: {row['INVOICE VALUE']}")
        print(f"INVOICE DATE: {row['INVOICE DATE']}")

    # Click on "Invoice Processing"
    print("\nClicking on 'Invoice Processing' button")
    click_element(
        driver,
        "#root > div > div > div > header > div > div.i-primary-nav.MuiBox-root.css-1dzmjrl > button:nth-child(3)",
        "Invoice Processing Button",
    )
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    # Figure out dropdowns, Invoice Number, Invoice Date

    for index, row in df.iterrows():
        print(f"\nProcessing row {index + 1}")
        po_number = str(row["PO NUMBER"])
        print(f"PO Number: {po_number}")
        pdf_file = None

        # Find the PDF file that contains the provided PO number
        for file in os.listdir(files_dir):
            if po_number in file and file.endswith(".pdf"):
                pdf_file = os.path.join(files_dir, file)
                print(f"Found PDF file: {pdf_file}")
                break

        if not pdf_file:
            print(f"PDF file for PO Number {po_number} not found. Skipping this row.")
            continue
        print("Clicking on 'Upload invoice' button")
        click_element(
            driver,
            "#root > div > div > div > div > main > div.invoice-container > section:nth-child(1) > div > div.MuiGrid-root.MuiGrid-item.MuiGrid-grid-xs-12.MuiGrid-grid-sm-12.MuiGrid-grid-md-6.MuiGrid-grid-lg-6.css-iol86l > label",
            "Upload invoice button",
        )

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
        )

        # Upload the PDF file
        print(f"Uploading PDF file: {pdf_file}")
        file_input = driver.find_element(By.CSS_SELECTOR, "input[type=file]")
        file_input.send_keys(pdf_file)

        # Select "Buying Type"
        buying_type = str(row["BUYING TYPE"])
        dropdownOne = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "body > div.MuiModal-root.css-8ndowl > div.invoiceuploadOfferModal.MuiBox-root.css-1kf1z5c > div > div:nth-child(4) > div > div:nth-child(1) > div > div",
                )
            )
        )
        dropdownOne.click()
        PrivateOption = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[text()='{buying_type}']"))
        )
        PrivateOption.click()

        # Select "Buying Centre"
        buying_centre = str(row["BUYING CENTRE"])
        dropdownTwo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "body > div.MuiModal-root.css-8ndowl > div.invoiceuploadOfferModal.MuiBox-root.css-1kf1z5c > div > div:nth-child(4) > div > div:nth-child(2) > div > div",
                )
            )
        )
        dropdownTwo.click()
        parts = buying_centre.split("-")
        first_part = parts[0].strip()
        rest_part = "-".join(parts[1:]).strip()
        xpath = f"//li[normalize-space(.) = '{first_part}-{rest_part}' or (starts-with(normalize-space(.), '{first_part}') and contains(normalize-space(.), '{rest_part}'))]"

        BuyingCenter = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        BuyingCenter.click()
        print(f"Successfully selected buying center: {buying_centre}")

        # Select "Vendor Code" VENDOR CODE WORKING
        vendor_code = str(row["VENDOR CODE"])
        print(f"Entering Vendor Code: {vendor_code}")
        vendor_code_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[3]/div/div/div/input"
        )

        # Clear any existing value
        vendor_code_input.clear()

        # Enter the first few characters of the vendor code
        vendor_code_input.send_keys(vendor_code[:3])

        # Wait for options to appear
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//li[contains(@class, 'MuiAutocomplete-option')]")
            )
        )

        # Find and click the correct option
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    f"//li[contains(@class, 'MuiAutocomplete-option') and contains(text(), '{vendor_code}')]",
                )
            )
        )
        option.click()

        print(f"Successfully selected vendor code: {vendor_code}")

        # Select "PO Number" PO NUMBER WORKING
        print(f"Entering PO Number: {po_number}")
        po_number_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[4]/div/div/div/input"
        )

        # Clear any existing value
        po_number_input.clear()

        # Enter the first few characters of the PO number
        po_number_input.send_keys(po_number[:8])

        # Wait for options to appear
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//li[contains(@class, 'MuiAutocomplete-option')]")
            )
        )

        # Find and click the correct option
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    f"//li[contains(@class, 'MuiAutocomplete-option') and contains(text(), '{po_number}')]",
                )
            )
        )
        option.click()
        actions = ActionChains(driver)
        actions.send_keys(Keys.TAB)
        actions.perform()
        print(f"Successfully selected PO Number: {po_number}")

        # Select "Invoice Number"
        invoice_number = str(row["INVOICE NUMBER"])
        print(f"Selecting Invoice Number: {invoice_number}")

        # Click to open the dropdown
        dropdownThree = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.MuiModal-root.css-8ndowl > div.invoiceuploadOfferModal.MuiBox-root.css-1kf1z5c > div > div:nth-child(4) > div > div:nth-child(6) > div > div > div"))
        )
        dropdownThree.click()

        # Wait for the dropdown options to be visible
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "ul.MuiList-root.MuiMenu-list"))
        )

        # Find and click the correct option
        try:
            option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[@role='option' and contains(@class, 'MuiMenuItem-root') and contains(text(), '{invoice_number}')]"))
            )
            option.click()
            print(f"Successfully selected Invoice Number: {invoice_number}")
        except TimeoutException:
            print(f"Exact match for Invoice Number {invoice_number} not found. Trying partial match.")
            try:
                # If exact match is not found, try matching the first part of the invoice number
                partial_invoice_number = invoice_number[10:]  # Adjust the length as needed
                option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//li[@role='option' and contains(@class, 'MuiMenuItem-root') and starts-with(text(), '{partial_invoice_number}')]"))
                )
                option.click()
                print(f"Selected partial match for Invoice Number: {option.text}")
            except TimeoutException:
                print(f"Failed to select Invoice Number: {invoice_number}")

        # Enter "Invoice Quantity"
        invoice_quantity = str(row["INVOICE QUANTITY"])
        print(f"Entering Invoice Quantity: {invoice_quantity}")
        invoice_quantity_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[7]/div/div/input"
        )
        invoice_quantity_input.send_keys(invoice_quantity)

        # Enter "Invoice Value"
        invoice_value = str(row["INVOICE VALUE"])
        print(f"Entering Invoice Value: {invoice_value}")
        invoice_value_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[8]/div/div/input"
        )
        invoice_value_input.send_keys(invoice_value)

        # Enter "Invoice Date" Needs to be entered in dd-mm-yy
        invoice_date = str(row["INVOICE DATE"])
        date_object = datetime.strptime(invoice_date, "%Y-%m-%d %H:%M:%S")

        # Format the date as dd-mm-yyyy
        formatted_date = date_object.strftime("%d-%m-%Y")

        print(f"Entering date: {formatted_date}")

        # Find the date input field
        date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div[3]/div/div[5]/div/div[3]/div/div/div/input'))
        )

        # Clear any existing value
        for num in range(0,10):
            date_input.send_keys(Keys.BACK_SPACE)

        # Enter the formatted date
        date_input.send_keys(formatted_date)

        print(f"Successfully entered date: {formatted_date}")

        declaration_checkbox = driver.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
        declaration_checkbox.click()
        time.sleep(10)
        # Click on "Submit"
        # print("Clicking Submit button")
        # submit_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        # submit_button.click()
        # time.sleep(2)
        #submit button commented out for testing, cancel button used instead
        cancelButton = driver.find_element(By.XPATH,"/html/body/div[3]/div[3]/div/div[6]/div[2]/div/button[1]")
        cancelButton.click()
        time.sleep(5)

    print("Invoice processing completed.")

def click_element(driver, selector, description):
    try:
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
        )
        element.click()
        print(f"Successfully clicked {description}")
    except Exception as e:
        print(f"Failed to click {description}. Error: {e}")

def Step3(driver, download_dir):
    # Click #dashboard
    click_element(driver, "#dashboard", "Home Button")
    time.sleep(5)

    # Wait for page to load
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    # Click the second button
    click_element(
        driver,
        "#root > div > div > div > header > div > div.i-primary-nav.MuiBox-root.css-1dzmjrl > button:nth-child(2)",
        "Tea Private Button",
    )
    time.sleep(5)
    # Wait for page to load
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    # Click the third list item
    click_element(
        driver,
        "#root > div > div > div > div > main > div.MuiDrawer-root.MuiDrawer-docked.i-drawer.css-1de6c1k > div > ul > li:nth-child(3) > a > div",
        "Billing Button",
    )

    # Set download directory
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # Wait for the table to be present
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiDataGrid-root"))
    )

    page_num = 1
    while True:
        print(f"Processing page {page_num}")
        try:
            rows = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, ".MuiDataGrid-row")
                )
            )
        except TimeoutException:
            print("Timeout waiting for rows to load.")
            break

        for row in rows:
            try:
                cells = row.find_elements(By.CSS_SELECTOR, ".MuiDataGrid-cell")
                bill_no = cells[0].text if cells else "Unknown"

                try:
                    # Find the download SVG icon
                    download_icon = row.find_element(
                        By.CSS_SELECTOR, "svg[data-testid='FileDownloadTwoToneIcon']"
                    )

                    # Find the parent div of the SVG icon
                    download_div = download_icon.find_element(
                        By.XPATH, "./ancestor::div[contains(@class, 'MuiGrid-root')]"
                    )
                except NoSuchElementException:
                    print(f"Download icon not found for Bill No: {bill_no}")
                    continue

                # Use ActionChains to move to the div and click
                ActionChains(driver).move_to_element(download_div).click().perform()

                # Wait for download to complete (adjust time if needed)
                time.sleep(5)

                # Get the most recent file in the download directory
                list_of_files = glob.glob(os.path.join(download_dir, "*"))
                if list_of_files:
                    latest_file = max(list_of_files, key=os.path.getctime)

                    # Rename the file and change extension to .xlsx
                    new_filename = f"{bill_no.replace('/', '-')}.xlsx"
                    new_file_path = os.path.join(download_dir, new_filename)

                    # Check if file already exists
                    if not os.path.exists(new_file_path):
                        os.rename(latest_file, new_file_path)
                        print(f"Downloaded and renamed: {new_filename}")
                    else:
                        print(f"File {new_filename} already exists. Skipping.")
                        os.remove(latest_file)
                else:
                    print(f"No file downloaded for Bill No: {bill_no}")
            except Exception as e:
                print(f"Error processing row: {e}")

        # Check if there's a next page
        try:
            next_button = driver.find_element(
                By.CSS_SELECTOR, "button[aria-label='Go to next page']"
            )
            if next_button.is_enabled():
                next_button.click()
                time.sleep(2)  # Wait for the next page to load
                page_num += 1
            else:
                print("No more pages.")
                break
        except NoSuchElementException:
            print("Next page button not found.")
            break

    print("Download process completed.")

def main():
    root = ctk.CTk()
    gui = InfinitiGUI(root)
    root.protocol(
        "WM_DELETE_WINDOW", gui.exit_application
    )  # Handle window close button
    root.mainloop()

if __name__ == "__main__":
    main()
