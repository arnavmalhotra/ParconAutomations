import PySimpleGUI as sg
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

class InfinitiGUI:
    def __init__(self):
        sg.theme('DarkTeal9')  # Set the theme for a modern look

        self.config_file = os.path.join(os.getenv("APPDATA"), "infiniti_config.json")
        self.load_config()

        layout = [
            [sg.Text('Infiniti Automation', font=('Helvetica', 20), justification='center', expand_x=True)],
            [sg.HorizontalSeparator()],
            [sg.Text('Username:', size=(15, 1)), sg.InputText(self.config.get("username", ""), key='-USERNAME-')],
            [sg.Text('Password:', size=(15, 1)), sg.InputText(self.config.get("password", ""), key='-PASSWORD-', password_char='*')],
            [sg.Text('Download Folder:', size=(15, 1)), sg.InputText(self.config.get("download_folder", ""), key='-DOWNLOAD-'), sg.FolderBrowse()],
            [sg.Text('Excel/PDF Files:', size=(15, 1)), sg.InputText(self.config.get("files_folder", ""), key='-FILES-'), sg.FolderBrowse()],
            [sg.Text('Summary Excel:', size=(15, 1)), sg.InputText(self.config.get("summary_file", ""), key='-SUMMARY-'), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
            [sg.HorizontalSeparator()],
            [sg.Checkbox('Step 1', default=self.config.get("step1", False), key='-STEP1-'),
             sg.Checkbox('Step 2', default=self.config.get("step2", False), key='-STEP2-'),
             sg.Checkbox('Step 3', default=self.config.get("step3", False), key='-STEP3-')],
            [sg.HorizontalSeparator()],
            [sg.Button('Run', size=(10, 1), button_color=('white', '#007D84')), 
             sg.Button('Exit', size=(10, 1), button_color=('white', '#B22222'))],
            [sg.Output(size=(60, 10), key='-OUTPUT-')]  # Add an output area for logs
        ]

        self.window = sg.Window('Infiniti Automation', layout, finalize=True)
        self.automation_thread = None

    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as f:
                self.config = json.load(f)
        else:
            self.config = {}

    def save_config(self):
        config = {
            "username": self.window['-USERNAME-'].get(),
            "password": self.window['-PASSWORD-'].get(),
            "download_folder": self.window['-DOWNLOAD-'].get(),
            "files_folder": self.window['-FILES-'].get(),
            "summary_file": self.window['-SUMMARY-'].get(),
            "step1": self.window['-STEP1-'].get(),
            "step2": self.window['-STEP2-'].get(),
            "step3": self.window['-STEP3-'].get(),
        }
        with open(self.config_file, 'w') as f:
            json.dump(config, f)

    def run_automation(self):
        self.save_config()

        chrome_options = Options()
        chrome_options.add_argument("--ignore-ssl-errors=yes")
        chrome_options.add_argument("--ignore-certificate-errors")

        download_dir = os.path.normpath(self.window['-DOWNLOAD-'].get().strip())
        files_dir = os.path.normpath(self.window['-FILES-'].get().strip())
        summary_file = os.path.normpath(self.window['-SUMMARY-'].get().strip())

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
            username = self.window['-USERNAME-'].get()
            password = self.window['-PASSWORD-'].get()
            driver.get("https://infiniti.tataconsumer.com/private/tea/billing")
            login(driver, username, password)

            if self.window['-STEP1-'].get():
                Step1(driver, files_dir)
            if self.window['-STEP2-'].get():
                Step2(driver, files_dir, summary_file)
            if self.window['-STEP3-'].get():
                Step3(driver, download_dir)

            sg.popup("Automation Complete", "The automation has finished running.")

        except Exception as e:
            error_message = str(e)
            sg.popup_error(f"An error occurred: {error_message}")
        finally:
            driver.quit()

    def run(self):
        while True:
            event, values = self.window.read()
            if event == sg.WINDOW_CLOSED or event == 'Exit':
                break
            elif event == 'Run':
                self.automation_thread = threading.Thread(target=self.run_automation)
                self.automation_thread.start()

        self.window.close()

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
            time.sleep(5)
            click_element(
                driver,
                "body > div.MuiModal-root.css-8ndowl > div.uploadOfferModal.MuiBox-root.css-h1jd1n > div > div:nth-child(2) > div > div > button.MuiButtonBase-root.MuiButton-root.MuiButton-Primary.MuiButton-sizeMedium.MuiButton-SizeMedium.MuiButton-colorPrimary.MuiButton-root.MuiButton-Primary.MuiButton-sizeMedium.MuiButton-SizeMedium.MuiButton-colorPrimary.cancel-btn.css-z1efht",
                "Cancel Button",
            )

def Step2(driver, files_dir, summary_file):
    print("Starting Step2 function")

    df = pd.read_excel(summary_file, header=1)

    print("DataFrame column types:")
    print(df.dtypes)

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

    print("\nClicking on 'Invoice Processing' button")
    click_element(
        driver,
        "#root > div > div > div > header > div > div.i-primary-nav.MuiBox-root.css-1dzmjrl > button:nth-child(3)",
        "Invoice Processing Button",
    )
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    for index, row in df.iterrows():
        print(f"\nProcessing row {index + 1}")
        po_number = str(row["PO NUMBER"])
        print(f"PO Number: {po_number}")
        pdf_file = None

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

        print(f"Uploading PDF file: {pdf_file}")
        file_input = driver.find_element(By.CSS_SELECTOR, "input[type=file]")
        file_input.send_keys(pdf_file)

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

        vendor_code = str(row["VENDOR CODE"])
        print(f"Entering Vendor Code: {vendor_code}")
        vendor_code_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[3]/div/div/div/input"
        )

        vendor_code_input.clear()
        vendor_code_input.send_keys(vendor_code[:3])

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

        print(f"Successfully selected vendor code: {vendor_code}")

        print(f"Entering PO Number: {po_number}")
        po_number_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[4]/div/div/div/input"
        )

        po_number_input.clear()
        po_number_input.send_keys(po_number[:8])

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
        print(f"Successfully selected PO Number: {po_number}")

        invoice_number = str(row["INVOICE NUMBER"])
        print(f"Selecting Invoice Number: {invoice_number}")

        dropdownThree = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.MuiModal-root.css-8ndowl > div.invoiceuploadOfferModal.MuiBox-root.css-1kf1z5c > div > div:nth-child(4) > div > div:nth-child(6) > div > div > div"))
        )
        dropdownThree.click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "ul.MuiList-root.MuiMenu-list"))
        )

        try:
            option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[@role='option' and contains(@class, 'MuiMenuItem-root') and contains(text(), '{invoice_number}')]"))
            )
            option.click()
            print(f"Successfully selected Invoice Number: {invoice_number}")
        except TimeoutException:
            print(f"Exact match for Invoice Number {invoice_number} not found. Trying partial match.")
            try:
                partial_invoice_number = invoice_number[10:]
                option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//li[@role='option' and contains(@class, 'MuiMenuItem-root') and starts-with(text(), '{partial_invoice_number}')]"))
                )
                option.click()
                print(f"Selected partial match for Invoice Number: {option.text}")
            except TimeoutException:
                print(f"Failed to select Invoice Number: {invoice_number}")

        invoice_quantity = str(row["INVOICE QUANTITY"])
        print(f"Entering Invoice Quantity: {invoice_quantity}")
        invoice_quantity_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[7]/div/div/input"
        )
        invoice_quantity_input.send_keys(invoice_quantity)

        invoice_value = str(row["INVOICE VALUE"])
        print(f"Entering Invoice Value: {invoice_value}")
        invoice_value_input = driver.find_element(
            By.XPATH, "/html/body/div[3]/div[3]/div/div[4]/div/div[8]/div/div/input"
        )
        invoice_value_input.send_keys(invoice_value)

        invoice_date = str(row["INVOICE DATE"])
        date_object = datetime.strptime(invoice_date, "%Y-%m-%d %H:%M:%S")
        formatted_date = date_object.strftime("%d-%m-%Y")

        print(f"Entering date: {formatted_date}")

        date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div[3]/div/div[5]/div/div[3]/div/div/div/input'))
        )

        for _ in range(10):
            date_input.send_keys(Keys.BACK_SPACE)

        date_input.send_keys(formatted_date)

        print(f"Successfully entered date: {formatted_date}")

        declaration_checkbox = driver.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
        declaration_checkbox.click()
        time.sleep(10)

        cancelButton = driver.find_element(By.XPATH,"/html/body/div[3]/div[3]/div/div[6]/div[2]/div/button[1]")
        cancelButton.click()
        time.sleep(5)

    print("Invoice processing completed.")

def Step3(driver, download_dir):
    click_element(driver, "#dashboard", "Home Button")
    time.sleep(5)

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    click_element(
        driver,
        "#root > div > div > div > header > div > div.i-primary-nav.MuiBox-root.css-1dzmjrl > button:nth-child(2)",
        "Tea Private Button",
    )
    time.sleep(5)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    click_element(
        driver,
        "#root > div > div > div > div > main > div.MuiDrawer-root.MuiDrawer-docked.i-drawer.css-1de6c1k > div > ul > li:nth-child(3) > a > div",
        "Billing Button",
    )

    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

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
                    download_icon = row.find_element(
                        By.CSS_SELECTOR, "svg[data-testid='FileDownloadTwoToneIcon']"
                    )
                    download_div = download_icon.find_element(
                        By.XPATH, "./ancestor::div[contains(@class, 'MuiGrid-root')]"
                    )
                except NoSuchElementException:
                    print(f"Download icon not found for Bill No: {bill_no}")
                    continue

                ActionChains(driver).move_to_element(download_div).click().perform()
                time.sleep(5)

                list_of_files = glob.glob(os.path.join(download_dir, "*"))
                if list_of_files:
                    latest_file = max(list_of_files, key=os.path.getctime)
                    new_filename = f"{bill_no.replace('/', '-')}.xlsx"
                    new_file_path = os.path.join(download_dir, new_filename)

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

        try:
            next_button = driver.find_element(
                By.CSS_SELECTOR, "button[aria-label='Go to next page']"
            )
            if next_button.is_enabled():
                next_button.click()
                time.sleep(2)
                page_num += 1
            else:
                print("No more pages.")
                break
        except NoSuchElementException:
            print("Next page button not found.")
            break

    print("Download process completed.")

def click_element(driver, selector, description):
    try:
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
        )
        element.click()
        print(f"Successfully clicked {description}")
    except Exception as e:
        print(f"Failed to click {description}. Error: {e}")

def main():
    gui = InfinitiGUI()
    gui.run()

if __name__ == "__main__":
    main()