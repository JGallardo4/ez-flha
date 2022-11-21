import click
from imap_tools import MailBox
from imap_tools import AND, OR, NOT
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
from datetime import datetime, timedelta
from pytz import timezone
from pdf2image import convert_from_path, convert_from_bytes
import tempfile
from PIL import Image
import io
from os import listdir, rename, path
import configparser
import shutil
import pyautogui

from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)

config = configparser.ConfigParser(allow_no_value=True)
config.read("./config.ini")

email_queue = dict()
sitedocs_submissions = dict()

@click.command()
@click.option("--headless", is_flag=True, flag_value=True, default=False, help="Run browser automation in headless mode.")
def process_flhas(headless):
    """ This program downloads pdf files from email, converts to PNG, and submits to Sitedocs. """

    check_latest_submission_from_memory()
    find_missing_dates(headless=True)
    emails = gather_emails()
    process_emails(emails)
    convert_pdfs()
    submit_flhas(headless=False)

def submit_flhas(headless):
    img_folder = config["IO"]["IMG_FOLDER"]
    img_folder_abs_path = path.abspath(img_folder)

    file_list = sorted(listdir(img_folder), key=lambda x: str_to_date(x[1].partition('_')[1]))
    print(file_list)

    for filename in file_list:
        flha_date_str = filename[1].partition('_')[0]

        full_img_path = f"{img_folder_abs_path}/{filename[1]}"

        if(flha_date_str in email_queue):
            if("images" in email_queue[flha_date_str]):
                email_queue[flha_date_str]["images"].append(full_img_path)
            else:
                email_queue[flha_date_str]["images"] = [full_img_path]
    
    for flha in email_queue:
        if(not check_file_is_archive(flha)):
            print(f"We need to process {flha}")
            driver = navigate_to_flhas(headless)

            # New FLHA
            new_flha_menu_option = driver.find_element(By.CSS_SELECTOR, "[data-id='sidebar-submenu-new-form']")
            new_flha_menu_option.click()

            # Label
            label = driver.find_element(By.ID, "company-info-label")
            label.send_keys(filename[1].rpartition('_')[0])

            # Permit
            permit = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[4]/div[2]/div[1]/div/div[2]/div[1]/div/div/textarea[1]")
            permit_value = None

            if(email_queue[flha]["permit"]):
                permit_value = email_queue[flha_date_str]["permit"]
            else:
                permit_value = click.prompt(f"Please enter the permit number for {flha_date_str}", type=str)
            
            permit.send_keys(permit_value)

            # Expiry
            expiry = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[4]/div[2]/div[3]/div/div[2]/div[1]/div/div/textarea[1]")
            expiry.send_keys(config["FORM"]["PERMIT_EXPIRY"])

            # EMP
            emp = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[4]/div[2]/div[7]/div/div[2]/div[1]/div/div/textarea[1]")
            emp.send_keys(config["FORM"]["MUSTER_POINTS"])

            # EAA
            eaa = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[4]/div[2]/div[8]/div/div[2]/div[1]/div/div/textarea[1]")
            eaa.send_keys(config["FORM"]["PROJECT"])

            # PPE
            ppe = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[4]/div[2]/div[9]/div/div[2]/div[1]/div[1]/button[1]")
            ppe.click()

            # Fit for duty
            ffd = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[4]/div[2]/div[10]/div/div[2]/div[1]/div/button[1]/span")
            ffd.click()

            # Vehicle
            vehicle = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[5]/div[2]/div[1]/div/div[2]/div[1]/div/div/textarea[1]")
            vehicle.send_keys(config["FORM"]["TRUCK"])

            # Walk around
            walk_around = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[5]/div[2]/div[2]/div/div[2]/div[1]/div/button[1]/span")
            walk_around.click()

            # Issues
            issues = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[5]/div[2]/div[3]/div/div[2]/div[1]/div/button[2]/span")
            issues.click()

            # Photos
            for image_path in email_queue[flha]["images"]:
                # Upload image
                upload_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[1]/div[9]/div[1]/button[1]")
                upload_button.click()
                pyautogui.write(image_path)
                pyautogui.press('enter')
                time.sleep(2)

            # Sign
            sign = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/main/div/div[2]/div[2]/div/button/span[1]")
            sign.click()

            # Worker
            worker = driver.find_element(By.XPATH, "/html/body/div[3]/div[3]/div/div[2]/div[3]/div/div[1]/div/div/div[1]/li/div/div/div[1]/div/h5")
            worker.click()

            # Wait for user input
            click.pause("Please submit your signature, then press any key to continue!")

            # Share
            # share_button = driver.find_elements(By.XPATH, "/html/body/div[1]/div/div[4]/main/header/div/button[2]/span[1]/svg")
            # share_button.click()
            
        else:
            # Move file to archive
            old = f'{config["IO"]["IMG_FOLDER"]}/{filename}'
            new = f'{config["IO"]["IMG_ARCHIVE"]}/{filename}'
            shutil.move(old, new)

def check_file_is_archive(date_str):
    file_date = str_to_date(date_str)
    latest_str = config["APP"]["LATEST_SUBMISSION"]
    latest_date = str_to_date(latest_str)

    return (latest_date - file_date) >= timedelta(0)

def str_to_date(date_str):
    return datetime.strptime(date_str, "%Y.%m.%d").replace(tzinfo=timezone("Canada/Mountain"))

def date_to_str(date_obj):
    return date_obj.strftime('%Y.%m.%d')

def print_days_since(latest_date, latest_date_str):
    days = get_time_since(latest_date).days

    print(f"It has been {days} days since your last submission on {latest_date_str}")

def check_up_to_date(latest_date):
    if(get_time_since(latest_date) <= timedelta(2)):
        print_days_since(latest_date, latest)
        print("You are up to date")
        quit()

def get_time_since(latest_date):
    result = datetime.today().replace(tzinfo=timezone("Canada/Mountain")) - latest_date
    return result

def save_pdf(path_dest, file):
    if (not path.exists(path_dest)):
            with open(path_dest, 'wb') as f:
                f.write(file.payload)
    else:
        filename = path.basename(path_dest)
        print(f"{filename} already exists")

def convert_pdfs():
    # Convert pdfs to jpegs and save locally
    pdf_folder = config["IO"]["PDF_FOLDER"]
    pdf_archive_folder = config["IO"]["PDF_ARCHIVE"]

    for index, filename in enumerate(listdir(pdf_folder)):
        file_date_str = filename.partition('_')[0]

        if(file_date_str not in sitedocs_submissions):
            if(not file_date_str in listdir(pdf_folder)):
                file_date = str_to_date(file_date_str)

                if (file_date > latest_date):
                    full_pdf_path = f'{pdf_folder}/{filename}'
                    with tempfile.TemporaryDirectory() as path:
                        images_from_path = convert_from_path(f"{full_pdf_path}", output_folder=config["IO"]["IMG_FOLDER"], fmt="jpeg", output_file=filename)
            else:
                print(f"jpgs for date {file_date_str} already exist")
        else:
            # Move file to archive
            old = f'{pdf_folder}/{filename}'
            new = f'{pdf_archive_folder}/{filename}'
            shutil.move(old, new)

    # Fix extra suffix
    img_folder = config["IO"]["IMG_FOLDER"]
    for index, filename in enumerate(listdir(img_folder)):
        dst = re.sub(r'(?<=A)(.*?)(?=-)', '', filename)
        dst = re.sub(r'-', '_', dst)
        src =f"{img_folder}/{filename}"
        dst =f"{img_folder}/{dst}"
        rename(src, dst)

def process_emails(emails):
    for email in emails:
        date = email.date.strftime('%Y.%m.%d')
        ast = config["FORM"]["ASSISTANT"].title().replace(" ", "")
        pc = config["FORM"]["PARTY_CHIEF"].title().replace(" ", "")
        pdf_filename = f"{date}_{pc}_{ast}_FLHA.pdf"
        pdf_folder = config["IO"]["PDF_FOLDER"]
        full_pdf_path = f"{pdf_folder}/{pdf_filename}"
        attachment = email.attachments[0]

        save_pdf(full_pdf_path, attachment)

def gather_emails():
    emails = []
    smtp_server = config["EMAIL"]["SMTP_SERVER"]
    from_email = config["EMAIL"]["FROM_EMAIL"]
    pwd = config["EMAIL"]["APP_PWD"]
    scanner_email = config["SCANNER"]["SCANNER_EMAIL"]

    with MailBox(smtp_server).login(from_email, pwd, 'INBOX') as mailbox:
        current_date = None
        for msg in mailbox.fetch(AND(seen=False, from_=scanner_email, subject="Attached Image"), mark_seen=False):
            if(len(msg.attachments)==1 and msg.date.date() not in sitedocs_submissions and (current_date == None or msg.date.day != current_date)):
                for att in msg.attachments:
                    if(att.content_type=="application/pdf"):
                        emails.append(msg)
                        current_date = msg.date.day
    
    print("Found the following files:")
    for msg in emails:        
        date_str = date_to_str(msg.date)

        permit = None

        if(msg.subject.isdigit() and msg.subject.length == 3):
            permit = msg.subject

        email_queue[date_str] = { "permit": permit }

        print(f"Date: {date_str} File: {msg.attachments[0].filename}")

    return emails

def navigate_to_flhas(headless):
    # Initialize driver
    options = webdriver.ChromeOptions()
    if(headless):
        options.add_argument("--headless")
    options.add_argument("--disable-notifications")

    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(100)
    driver.set_page_load_timeout(-1)
    
    # Log in
    log_in_sitedocs(driver=driver)
    time.sleep(5)
    
    # Select project
    project_radio_button = driver.find_element(By.XPATH, "/html/body/div[3]/div[3]/div/div[3]/ul/div[1]/div/div/div[1]/li/div/div/div[1]/div")
    project_radio_button.click()
    time.sleep(5)

    # Open forms menu
    forms_menu = driver.find_element(By.XPATH, '//*[@id="forms-nav-item"]')
    forms_menu.click()
    time.sleep(5)
    
    # Open FLHAs Menu
    flha_menu = driver.find_element(By.CSS_SELECTOR, "[data-id='sidebar-form-item-name-button']")
    flha_menu.click()

    return driver

def find_missing_dates(headless):
    driver = navigate_to_flhas(headless)

    # Check previously signed FLHAs 
    history_menu = driver.find_elements(By.XPATH, "/html/body/div[1]/div/div[3]/div/div/div/div[22]/div/div/div[2]/nav/a[4]/div/div/div/div")
    history_menu[0].click()

    # Collect history items
    history_items = driver.find_elements(By.CSS_SELECTOR, "[data-id='sidebar-signed-form-item-label-text']")
    latest_submissions = []

    for item in history_items:
        if(config["FORM"]["PARTY_CHIEF"].title().replace(" ", "") in item.text and config["FORM"]["ASSISTANT"].title().replace(" ", "") in item.text):
            latest_submissions.append(item.text)
    
    date_strings = []
    for item in latest_submissions:
        date_strings.append(item[:item.index("_")])

    for date in date_strings:
        sitedocs_submissions.append(datetime.strptime(date, "%Y.%m.%d").replace(tzinfo=timezone("Canada/Mountain")))

    driver.quit()

def log_in_sitedocs(driver):
    try:
        print("Logging in to SiteDocs...")

        driver.get(config["SITEDOCS"]["URL"])
        
        username_field = driver.find_element(By.ID, "Username")
        username_field.send_keys(config["SITEDOCS"]["USERNAME"])

        next_button = driver.find_element(By.XPATH, "/html/body/div/div/div/div[2]/div/form/fieldset/button")
        next_button.click()
        
        username_field = driver.find_element(By.XPATH, '//*[@id="Password"]')
        username_field.send_keys(config["SITEDOCS"]["PWD"])

        next_button = driver.find_element(By.XPATH, "/html/body/div/div/div/div[2]/div/form/fieldset/button")
        next_button.click()

        web_app_button = driver.find_element(By.XPATH, "/html/body/div/div/div/div[2]/div/form/fieldset/div/div[2]/a")
        web_app_button.click()

        print("Log in successful!")
    except Exception as ex:
        print(ex)
        print("Unable to log in")
        quit()

def test():
    submit_flhas(True)

if __name__ == '__main__':
    process_flhas()
    # test()
