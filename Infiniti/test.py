import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import DateEntry
import json
import os
import threading
import ctypes
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import glob
import pandas as pd
from datetime import datetime
from selenium.common.exceptions import WebDriverException
import traceback
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class InfinitiGUI:
    def __init__(self, master):
        self.master = master
        master.title("Infiniti Automation")
        master.geometry("600x750")  # Increased height to accommodate date selection

        self.config_file = os.path.join(os.getenv("APPDATA"), "infiniti_config.json")
        self.load_config()

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="20 20 20 20")
        main_frame.pack(fill=BOTH, expand=YES)

        ttk.Label(
            main_frame, text="Infiniti Automation", font=("TkDefaultFont", 24, "bold")
        ).pack(pady=(0, 20))

        # Username
        ttk.Label(main_frame, text="Username:").pack(fill=X)
        self.username_entry = ttk.Entry(main_frame)
        self.username_entry.pack(fill=X, pady=(0, 10))
        self.username_entry.insert(0, self.config.get("username", ""))

        # Password
        ttk.Label(main_frame, text="Password:").pack(fill=X)
        self.password_entry = ttk.Entry(main_frame, show="*")
        self.password_entry.pack(fill=X, pady=(0, 10))
        self.password_entry.insert(0, self.config.get("password", ""))

        # Download Folder
        ttk.Label(main_frame, text="Download Folder:").pack(fill=X)
        download_frame = ttk.Frame(main_frame)
        download_frame.pack(fill=X, pady=(0, 10))
        self.download_entry = ttk.Entry(download_frame)
        self.download_entry.pack(side=LEFT, expand=YES, fill=X)
        self.download_entry.insert(0, self.config.get("download_folder", ""))
        ttk.Button(
            download_frame,
            text="Browse",
            command=lambda: self.browse_folder(self.download_entry),
        ).pack(side=RIGHT)

        # Excel and PDF Files Location
        ttk.Label(main_frame, text="Excel/PDF Files:").pack(fill=X)
        files_frame = ttk.Frame(main_frame)
        files_frame.pack(fill=X, pady=(0, 10))
        self.files_entry = ttk.Entry(files_frame)
        self.files_entry.pack(side=LEFT, expand=YES, fill=X)
        self.files_entry.insert(0, self.config.get("files_folder", ""))
        ttk.Button(
            files_frame,
            text="Browse",
            command=lambda: self.browse_folder(self.files_entry),
        ).pack(side=RIGHT)

        # Summary Excel File Location
        ttk.Label(main_frame, text="Summary Excel:").pack(fill=X)
        summary_frame = ttk.Frame(main_frame)
        summary_frame.pack(fill=X, pady=(0, 10))
        self.summary_entry = ttk.Entry(summary_frame)
        self.summary_entry.pack(side=LEFT, expand=YES, fill=X)
        self.summary_entry.insert(0, self.config.get("summary_file", ""))
        ttk.Button(
            summary_frame,
            text="Browse",
            command=lambda: self.browse_file(self.summary_entry),
        ).pack(side=RIGHT)

        # Date selection
        date_frame = ttk.Frame(main_frame)
        date_frame.pack(fill=X, pady=(0, 10))

        ttk.Label(date_frame, text="From Date:").pack(side=LEFT)
        self.from_date = DateEntry(date_frame, dateformat="%d-%m-%Y")
        self.from_date.pack(side=LEFT, padx=(0, 10))
        self.from_date.entry.insert(0, self.config.get("from_date", ""))

        ttk.Label(date_frame, text="To Date:").pack(side=LEFT)
        self.to_date = DateEntry(date_frame, dateformat="%d-%m-%Y")
        self.to_date.pack(side=LEFT)
        self.to_date.entry.insert(0, self.config.get("to_date", ""))

        # Checkboxes
        checkbox_frame = ttk.Frame(main_frame)
        checkbox_frame.pack(fill=X, pady=(10, 20))
        self.step1_var = tk.BooleanVar(value=self.config.get("step1", False))
        self.step2_var = tk.BooleanVar(value=self.config.get("step2", False))
        self.step3_var = tk.BooleanVar(value=self.config.get("step3", False))

        ttk.Checkbutton(checkbox_frame, text="Step 1", variable=self.step1_var).pack(
            side=LEFT, expand=YES
        )
        ttk.Checkbutton(checkbox_frame, text="Step 2", variable=self.step2_var).pack(
            side=LEFT, expand=YES
        )
        ttk.Checkbutton(checkbox_frame, text="Step 3", variable=self.step3_var).pack(
            side=LEFT, expand=YES
        )

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=(0, 20))
        ttk.Button(
            button_frame,
            text="Run",
            command=self.run_automation,
            style="success.TButton",
        ).pack(side=LEFT, expand=YES)
        ttk.Button(
            button_frame,
            text="Exit",
            command=self.exit_application,
            style="danger.TButton",
        ).pack(side=RIGHT, expand=YES)

        # Output area
        self.output_area = tk.Text(main_frame, height=10, wrap=tk.WORD)
        self.output_area.pack(fill=BOTH, expand=YES)

        # Redirect stdout to the output area
        sys.stdout = TextRedirector(self.output_area)

    def browse_folder(self, entry):
        folder = filedialog.askdirectory()
        if folder:
            entry.delete(0, tk.END)
            entry.insert(0, folder)

    def browse_file(self, entry):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file:
            entry.delete(0, tk.END)
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
            "from_date": self.from_date.entry.get(),
            "to_date": self.to_date.entry.get(),
        }
        with open(self.config_file, "w") as f:
            json.dump(config, f)

    def exit_application(self):
        if messagebox.askokcancel("Exit", "Are you sure you want to exit?"):
            logging.info("Exiting application...")
            if (
                hasattr(self, "automation_thread")
                and self.automation_thread
                and self.automation_thread.is_alive()
            ):
                thread_id = self.automation_thread.ident
                ret = ctypes.pythonapi.PyThreadState_SetAsyncExc(
                    ctypes.c_long(thread_id), ctypes.py_object(SystemExit)
                )
                if ret > 1:
                    ctypes.pythonapi.PyThreadState_SetAsyncExc(
                        ctypes.c_long(thread_id), None
                    )
                    logging.error("Exception raise failure")
            self.master.quit()
            self.master.destroy()
            sys.exit()

    def run_automation(self):
        self.automation_thread = threading.Thread(target=self._run_automation)
        self.automation_thread.start()

    def _run_automation(self):
        self.save_config()

        # Check for warnings based on selected steps
        warnings = []

        if self.step1_var.get():
            if not self.files_entry.get().strip():
                warnings.append("Step 1 requires Excel/PDF Files directory.")

        if self.step2_var.get():
            if not self.files_entry.get().strip():
                warnings.append("Step 2 requires Excel/PDF Files directory.")
            if not self.summary_entry.get().strip():
                warnings.append("Step 2 requires Summary Excel file.")

        if self.step3_var.get():
            if not self.download_entry.get().strip():
                warnings.append("Step 3 requires Download Folder.")
            if not self.from_date.entry.get() or not self.to_date.entry.get():
                warnings.append("Step 3 requires From Date and To Date.")

        if warnings:
            warning_message = "\n".join(warnings)
            if not messagebox.askyesno(
                "Warning", f"{warning_message}\n\nDo you want to continue anyway?"
            ):
                return

        chrome_options = Options()
        chrome_options.add_argument("--ignore-ssl-errors=yes")
        chrome_options.add_argument("--ignore-certificate-errors")

        download_dir = os.path.normpath(self.download_entry.get().strip())
        files_dir = os.path.normpath(self.files_entry.get().strip())
        summary_file = os.path.normpath(self.summary_entry.get().strip())

        logging.info(f"Download Directory: {download_dir}")
        logging.info(f"Files Directory: {files_dir}")
        logging.info(f"Summary File: {summary_file}")

        chrome_options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.automatic_downloads": 1,  # Allow multiple downloads
            },
        )

        while True:
            try:
                driver = webdriver.Chrome(options=chrome_options)

                try:
                    username = self.username_entry.get()
                    password = self.password_entry.get()
                    driver.get("https://infiniti.tataconsumer.com/private/tea/billing")
                    WebDriverWait(driver, 60).until(
                        EC.visibility_of_element_located(
                            (
                                By.XPATH,
                                "/html/body/div/div/div/div/div/div[2]/div/p/div/div[3]/div/div/div/div/form/div/fieldset/div[1]/div/div/input",
                            )
                        )
                    )
                    login(driver, username, password)
                    WebDriverWait(driver, 60).until(
                        EC.visibility_of_element_located(
                            (
                                By.XPATH,
                                "/html/body/div[1]/div/div/div/header/div/div[1]/button[1]",
                            )
                        )
                    )

                    if self.step1_var.get():
                        Step1(driver, files_dir)

                    if self.step2_var.get():
                        Step2(driver, files_dir, summary_file)

                    if self.step3_var.get():
                        from_date = self.from_date.entry.get()
                        to_date = self.to_date.entry.get()
                        Step3(driver, download_dir, from_date, to_date)

                    messagebox.showinfo(
                        "Automation Complete", "The automation has finished running."
                    )
                    break  # Exit the loop if everything completes successfully

                except WebDriverException:
                    error_message = traceback.format_exc()
                    logging.error(f"Browser was closed unexpectedly. Error: {error_message}")
                    if messagebox.askyesno(
                        "Browser Closed",
                        f"The browser window was closed. Do you want to restart the automation?{error_message}",
                    ):
                        logging.info("Restarting automation...")
                        continue  # Restart the loop
                    else:
                        logging.info("Automation cancelled by user.")
                        break  # Exit the loop

            except Exception as e:
                error_message = str(e)
                logging.error(f"An error occurred: {error_message}")
                messagebox.showerror("Error", f"An error occurred: {error_message}")
                break  # Exit the loop for other exceptions

            finally:
                if "driver" in locals():
                    driver.quit()


class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, str):
        self.widget.insert(tk.END, str)
        self.widget.see(tk.END)

    def flush(self):
        pass


def login(driver, username, password):
    logging.info("Starting login process")
    WebDriverWait(driver, 60).until(EC.title_contains("Infiniti"))
    username_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='username']"))
    )
    username_field.send_keys(username)
    logging.info(f"Entered username: {username}")
    time.sleep(1)  # Increased delay
    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By. CSS_SELECTOR, "input[name='password']"))
    )
    password_field.send_keys(password)
    logging.info("Entered password")
    time.sleep(1)  # Increased delay
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
    logging.info("Successfully clicked the Sign in button")
    time.sleep(1)  # Increased delay


def Step1(driver, files_dir):
    logging.info("Starting Step1 function")
    time.sleep(5)
    click_element(driver, "#dashboard", "Home Button")
    time.sleep(5)
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )
    click_element(
        driver,
        "#root > div > div > div > header > div > div.i-primary-nav.MuiBox-root.css-1dzmjrl > button:nth-child(2)",
        "Tea Private Button",
    )
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )
    click_element(
        driver,
        "#root > div > div > div > div > main > div.MuiDrawer-root.MuiDrawer-docked.i-drawer.css-1de6c1k > div > ul > li:nth-child(3) > a > div",
        "Billing Button",
    )
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "#root > div > div > div > div > main > div.rightside-modules > div > div:nth-child(1) > div > div.MuiGrid-root.MuiGrid-item.MuiGrid-grid-xs-2.MuiGrid-grid-sm-2.css-605f2w > button",
            )
        )
    )
    for filename in os.listdir(files_dir):
        file_path = os.path.join(files_dir, filename)
        if os.path.isfile(file_path) and (
            filename.endswith(".xlsx") or filename.endswith(".xls")
        ):
            logging.info(f"Processing file: {filename}")
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
            logging.info(f"Uploaded file: {file_path}")
            time.sleep(5)
            click_element(
                driver,
                "body > div.MuiModal-root.css-8ndowl > div.uploadOfferModal.MuiBox-root.css-h1jd1n > div > div:nth-child(2) > div > div > button.MuiButtonBase-root.MuiButton-root.MuiButton-Primary.MuiButton-sizeMedium.MuiButton-SizeMedium.MuiButton-colorPrimary.MuiButton-root.MuiButton-Primary.MuiButton-sizeMedium.MuiButton-SizeMedium.MuiButton-colorPrimary.upload-btn.css-z1efht",
                "Upload Button",
            )
            time.sleep(5)  # Increased delay
    logging.info("Step1 completed")


def Step2(driver, files_dir, summary_file):
    logging.info("Starting Step2 function")

    df = pd.read_excel(summary_file)

    logging.info("DataFrame column types:")
    logging.info(df.dtypes)

    logging.info("All values that will be input:")
    for index, row in df.iterrows():
        logging.info(f"\nRow {index + 1}:")
        for column in df.columns:
            logging.info(f"{column}: {row[column]}")

    logging.info("Clicking on 'Invoice Processing' button")
    click_element(
        driver,
        "#root > div > div > div > header > div > div.i-primary-nav.MuiBox-root.css-1dzmjrl > button:nth-child(3)",
        "Invoice Processing Button",
    )
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )
    time.sleep(2)  # Increased delay

    for index, row in df.iterrows():
        logging.info(f"\nProcessing row {index + 1}")
        po_number = str(row["PO NUMBER"])
        logging.info(f"PO Number: {po_number}")
        pdf_file = None

        for file in os.listdir(files_dir):
            if po_number in file and file.endswith(".pdf"):
                pdf_file = os.path.join(files_dir, file)
                logging.info(f"Found PDF file: {pdf_file}")
                break

        if not pdf_file:
            logging.warning(f"PDF file for PO Number {po_number} not found. Skipping this row.")
            continue

        logging.info("Clicking on 'Upload invoice' button")
        click_element(
            driver,
            "#root > div > div > div > div > main > div.invoice-container > section:nth-child(1) > div > div.MuiGrid-root.MuiGrid-item.MuiGrid-grid-xs-12.MuiGrid-grid-sm-12.MuiGrid-grid-md-6.MuiGrid-grid-lg-6.css-iol86l > label",
            "Upload invoice button",
        )

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
        )
        time.sleep(2)  # Increased delay

        logging.info(f"Uploading PDF file: {pdf_file}")
        file_input = driver.find_element(By.CSS_SELECTOR, "input[type=file]")
        file_input.send_keys(pdf_file)
        time.sleep(2)  # Increased delay

        buying_type = str(row["BUYING TYPE"])
        logging.info(f"Selecting Buying Type: {buying_type}")
        dropdownOne = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "body > div.MuiModal-root.css-8ndowl > div.invoiceuploadOfferModal.MuiBox-root.css-1kf1z5c > div > div:nth-child(4) > div > div:nth-child(1) > div > div",
                )
            )
        )
        dropdownOne.click()
        time.sleep(2)  # Increased delay
        PrivateOption = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[text()='{buying_type}']"))
        )
        PrivateOption.click()
        logging.info(f"Successfully selected Buying Type: {buying_type}")
        time.sleep(3)  # Increased delay

        buying_centre = str(row["BUYING CENTRE"])
        logging.info(f"Selecting Buying Centre: {buying_centre}")
        dropdownTwo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "body > div.MuiModal-root.css-8ndowl > div.invoiceuploadOfferModal.MuiBox-root.css-1kf1z5c > div > div:nth-child(4) > div > div:nth-child(2) > div > div",
                )
            )
        )
        dropdownTwo.click()
        time.sleep(2)  # Increased delay
        parts = buying_centre.split("-")
        first_part = parts[0].strip()
        rest_part = "-".join(parts[1:]).strip()
        xpath = f"//li[normalize-space(.) = '{first_part}-{rest_part}' or (starts-with(normalize-space(.), '{first_part}') and contains(normalize-space(.), '{rest_part}'))]"

        BuyingCenter = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        BuyingCenter.click()
        logging.info(f"Successfully selected Buying Centre: {buying_centre}")
        time.sleep(3)  # Increased delay

        vendor_code = str(row["VENDOR CODE"])
        logging.info(f"Entering Vendor Code: {vendor_code}")
        vendor_code_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[3]/div/div/div/input"
        )

        vendor_code_input.clear()
        vendor_code_input.send_keys(vendor_code[:3])
        time.sleep(2)  # Increased delay

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//li[contains(@class, 'MuiAutocomplete-option')]")
            )
        )

        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    f"//li[contains(@class, 'MuiAutocomplete-option') and contains(text(), '{vendor_code}')]",
                )
            )
        )
        option.click()
        logging.info(f"Successfully selected Vendor Code: {vendor_code}")
        time.sleep(3)  # Increased delay

        logging.info(f"Entering PO Number: {po_number}")
        po_number_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[4]/div/div/div/input"
        )

        po_number_input.clear()
        po_number_input.send_keys(po_number[:8])
        time.sleep(2)  # Increased delay

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//li[contains(@class, 'MuiAutocomplete-option')]")
            )
        )

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
        logging.info(f"Successfully selected PO Number: {po_number}")
        time.sleep(3)  # Increased delay

        invoice_number = str(row["INVOICE NUMBER"])
        logging.info(f"Selecting Invoice Number: {invoice_number}")

        dropdownThree = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "body > div.MuiModal-root.css-8ndowl > div.invoiceuploadOfferModal.MuiBox-root.css-1kf1z5c > div > div:nth-child(4) > div > div:nth-child(6) > div > div > div",
                )
            )
        )
        dropdownThree.click()
        time.sleep(2)  # Increased delay

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "ul.MuiList-root.MuiMenu-list")
            )
        )

        try:
            option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        f"//li[@role='option' and contains(@class, 'MuiMenuItem-root') and contains(text(), '{invoice_number}')]",
                    )
                )
            )
            option.click()
            logging.info(f"Successfully selected Invoice Number: {invoice_number}")
        except TimeoutException:
            logging.warning(f"Exact match for Invoice Number {invoice_number} not found. Trying partial match.")
            try:
                partial_invoice_number = invoice_number[10:]
                option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            f"//li[@role='option' and contains(@class, 'MuiMenuItem-root') and starts-with(text(), '{partial_invoice_number}')]",
                        )
                    )
                )
                option.click()
                logging.info(f"Selected partial match for Invoice Number: {option.text}")
            except TimeoutException:
                logging.error(f"Failed to select Invoice Number: {invoice_number}")
        time.sleep(3)  # Increased delay

        invoice_quantity = str(row["INVOICE QUANTITY"])
        logging.info(f"Entering Invoice Quantity: {invoice_quantity}")
        invoice_quantity_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[7]/div/div/input"
        )
        invoice_quantity_input.send_keys(invoice_quantity)
        time.sleep(2)  # Increased delay

        invoice_value = str(row["INVOICE VALUE"])
        logging.info(f"Entering Invoice Value: {invoice_value}")
        invoice_value_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[8]/div/div/input"
        )
        invoice_value_input.send_keys(invoice_value)
        time.sleep(2)  # Increased delay

        invoice_date = str(row["INVOICE DATE"])
        date_object = datetime.strptime(invoice_date, "%Y-%m-%d %H:%M:%S")
        formatted_date = date_object.strftime("%d-%m-%Y")

        logging.info(f"Entering Invoice Date: {formatted_date}")

        date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div[3]/div[3]/div/div[5]/div/div[3]/div/div/div/input",
                )
            )
        )

        for _ in range(10):
            date_input.send_keys(Keys.BACK_SPACE)

        date_input.send_keys(formatted_date)

        logging.info(f"Successfully entered Invoice Date: {formatted_date}")
        time.sleep(3)  # Increased delay

        declaration_checkbox = driver.find_element(
            By.CSS_SELECTOR, "input[type='checkbox']"
        )
        declaration_checkbox.click()
        logging.info("Clicked declaration checkbox")
        time.sleep(10)  # Increased delay

        cancelButton = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[6]/div[2]/div/button[2]"
        )
        cancelButton.click()
        logging.info("Clicked cancel button")
        time.sleep(2)  # Increased delay

    logging.info("Invoice processing completed.")


def Step3(driver, download_dir, from_date, to_date):
    logging.info("Starting Step3 function")
    click_element(driver, "#dashboard", "Home Button")
    time.sleep(2)

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    click_element(
        driver,
        "#root > div > div > div > header > div > div.i-primary-nav.MuiBox-root.css-1dzmjrl > button:nth-child(2)",
        "Tea Private Button",
    )
    time.sleep(2)
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    click_element(
        driver,
        "#root > div > div > div > div > main > div.MuiDrawer-root.MuiDrawer-docked.i-drawer.css-1de6c1k > div > ul > li:nth-child(3) > a > div",
        "Billing Button",
    )

    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
        logging.info(f"Created download directory: {download_dir}")

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiDataGrid-root"))
    )
    time.sleep(2)
    # Add date range selection
    from_date_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "/html/body/div[1]/div/div/div/div/main/div[2]/div/div[2]/div/div[5]/div/div/div[2]/div/div/div/input",
            )
        )
    )
    for num in range(0, 10):
        from_date_input.send_keys(Keys.BACKSPACE)
    from_date_input.send_keys(from_date)
    logging.info(f"Entered From Date: {from_date}")
    time.sleep(2)  # Increased delay

    to_date_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "/html/body/div[1]/div/div/div/div/main/div[2]/div/div[2]/div/div[5]/div/div/div[3]/div/div/div/input",
            )
        )
    )
    for num in range(0, 10):
        to_date_input.send_keys(Keys.BACKSPACE)
    to_date_input.send_keys(to_date)
    logging.info(f"Entered To Date: {to_date}")
    time.sleep(2)  # Increased delay

    # Wait for the table to update
    time.sleep(2)

    page_num = 1
    downloads_remaining = False
    while True:
        logging.info(f"Processing page {page_num}")
        try:
            rows = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, ".MuiDataGrid-row")
                )
            )
        except TimeoutException:
            logging.warning("Timeout waiting for rows to load.")
            break

        for row in rows:
            try:
                cells = row.find_elements(By.CSS_SELECTOR, ".MuiDataGrid-cell")
                bill_no = cells[0].text if cells else "Unknown"

                try:
                    download_icon = row.find_element(
                        By.CSS_SELECTOR, "svg[data-testid='FileDownloadTwoToneIcon']"
                    )
                    download_div = download_icon.find_element(
                        By.XPATH, "./ancestor::div[contains(@class, 'MuiGrid-root')]"
                    )
                except NoSuchElementException:
                    logging.warning(f"Download icon not found for Bill No: {bill_no}")
                    continue

                ActionChains(driver).move_to_element(download_div).click().perform()
                driver.execute_script("window.scrollBy(0, 300);")
                time.sleep(2)

                list_of_files = glob.glob(os.path.join(download_dir, "*"))
                if list_of_files:
                    latest_file = max(list_of_files, key=os.path.getctime)
                    new_filename = f"{bill_no.replace('/', '-')}.xlsx"
                    new_file_path = os.path.join(download_dir, new_filename)

                    if not os.path.exists(new_file_path):
                        os.rename(latest_file, new_file_path)
                        logging.info(f"Downloaded and renamed: {new_filename}")
                        downloads_remaining = True

                        # Scroll down a little after successful download
                        driver.execute_script("window.scrollBy(0, 300);")
                        time.sleep(0.5)  # Short pause to allow scrolling
                    else:
                        logging.info(f"File {new_filename} already exists. Skipping.")
                        os.remove(latest_file)
                else:
                    logging.warning(f"No file downloaded for Bill No: {bill_no}")
            except Exception as e:
                logging.error(f"Error processing row: {e}")

        try:
            time.sleep(2)
            next_button = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[@aria-label='Go to next page']")
                )
            )
            if next_button.is_enabled():
                next_button.click()
                logging.info("Clicked next page button")
                time.sleep(2)
                driver.execute_script("window.scrollTo(0, 0);")
                page_num += 1
            else:
                logging.info("No more pages.")
                break
        except NoSuchElementException:
            logging.warning("Next page button not found.")
            break

    logging.info("Download process completed.")

    if not downloads_remaining:
        logging.info("No downloads left. Closing the browser.")
        driver.quit()

    logging.info("Step 3 completed successfully.")
    messagebox.showinfo("Step 3 Complete", "Step 3 has finished running.")


def click_element(driver, selector, description):
    try:
        element = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
        )
        element.click()
        logging.info(f"Successfully clicked {description}")
        time.sleep(2)  # Increased delay
    except Exception as e:
        logging.error(f"Failed to click {description}. Error: {e}")


def main():
    root = ttk.Window(themename="superhero")
    app = InfinitiGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()