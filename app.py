import click
import config
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
from os.path import exists as file_exists
from pdf2image import convert_from_path, convert_from_bytes
import tempfile
from PIL import Image
import io
from os import listdir, rename

from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)

@click.command()
def process_flhas():
    """ This program downloads pdf files from email, converts to PNG, and submits to Sitedocs. """

    latest_submission = find_missing_dates()
    update_latest_submission(latest_submission)
    emails = gather_emails(latest_submission)
    process_emails(emails)
    convert_pdfs()
    submit_flhas()

def submit_flhas():
    queue = dict()
    for index, filename in enumerate(listdir(config.IMG_FOLDER)):


def update_latest_submission(latest_submission):
    

def save_pdf(path, file):
    if (not file_exists(path)):
            with open(path, 'wb') as f:
                f.write(file.payload)
    else:
        print(f"{path} already exists")

def convert_pdfs():
    # Convert pdfs to jpegs and save locally
    for index, filename in enumerate(listdir(config.PDF_FOLDER)):    
        full_pdf_path = f"{config.PDF_FOLDER}/{filename}"

        with tempfile.TemporaryDirectory() as path:
            images_from_path = convert_from_path(f"{full_pdf_path}", output_folder=config.IMG_FOLDER, fmt="jpeg", output_file=filename)

    # Fix extra suffix
    for index, filename in enumerate(listdir(config.IMG_FOLDER)):
        dst = re.sub(r'(?<=A)(.*?)(?=-)', '', filename)
        dst = re.sub(r'-', '_', dst)
        src =f"{config.IMG_FOLDER}/{filename}"
        dst =f"{config.IMG_FOLDER}/{dst}"
        rename(src, dst)

def process_emails(emails):
    for email in emails:
        date = email.date.strftime('%Y.%m.%d')
        ast = config.ASSISTANT.title().replace(" ", "")
        pc = config.PARTY_CHIEF.title().replace(" ", "")
        pdf_filename = f"{date}_{pc}_{ast}_FLHA.pdf"
        full_pdf_path = f"{config.PDF_FOLDER}/{pdf_filename}"
        attachment = email.attachments[0]

        save_pdf(full_pdf_path, attachment)

def gather_emails(latest_submission):
    emails = []

    with MailBox(config.SMTP_SERVER).login(config.FROM_EMAIL, config.APP_PWD, 'INBOX') as mailbox:
        current_date = None
        for msg in mailbox.fetch(AND(seen=False, from_=config.SCANNER_EMAIL, subject="Attached Image"), mark_seen=False):
            if(len(msg.attachments)==1 and msg.date.date() > latest_submission.date() and (current_date == None or msg.date.day != current_date)):
                for att in msg.attachments:
                    if(att.content_type=="application/pdf"):
                        emails.append(msg)
                        current_date = msg.date.day
    
    print("Found the following files")
    for msg in emails:
        date = msg.date.strftime('%Y.%m.%d')
        print(f"Date: {date} File: {msg.attachments[0].filename}")

    return emails

def find_missing_dates():
    # Initialize driver
    options = webdriver.FirefoxOptions()
    options.add_argument("--headless")
    options.set_preference("dom.webnotifications.enabled", False)

    driver = webdriver.Firefox(options=options)
    driver.implicitly_wait(20)    

    # Log in
    log_in_sitedocs(driver=driver)

    time.sleep(5)

    # Select project
    project_radio_button = driver.find_element(By.XPATH, "/html/body/div[3]/div[3]/div/div[3]/ul/div[1]/div/div/div[1]/li/div/div/div[1]/div/span/span/input")
    project_radio_button.click()

    time.sleep(5)

    # Open forms menu
    forms_menu = driver.find_element(By.XPATH, '//*[@id="forms-nav-item"]')
    forms_menu.click()

    time.sleep(5)

    # Open FLHAs Menu
    flha_menu = driver.find_element(By.CSS_SELECTOR, "[data-id='sidebar-form-item-name-button']")
    flha_menu.click()

    # Check previously signed FLHAs 
    history_menu = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[3]/div/div/div/div[22]/div/div/div[2]/nav/a[4]/div/div/div/div/div[1]/div/span")
    history_menu.click()

    # Collect history items
    history_items = driver.find_elements(By.CSS_SELECTOR, "[data-id='sidebar-signed-form-item-label-text']")
    latest_submissions = []

    for item in history_items:
        if(config.PARTY_CHIEF.title().replace(" ", "") in item.text and config.ASSISTANT.title().replace(" ", "") in item.text):
            latest_submissions.append(item.text)
    
    date_strings = []
    for item in latest_submissions:
        date_strings.append(item[:item.index("_")])

    dates = []
    for date in date_strings:
        dates.append(datetime.strptime(date, "%Y.%m.%d").replace(tzinfo=timezone("Canada/Mountain")))

    latest = max(dates)
    time_since = datetime.today().replace(tzinfo=timezone("Canada/Mountain")) - latest
    latest_str = latest.strftime('%Y.%m.%d')

    print(f"It has been {time_since.days} days since your last submission on {latest_str}")

    if(time_since <= timedelta(2)):
        print("You are up to date")
        quit()
    else:
        print("You are not up to date")
        return latest

    driver.quit()

def log_in_sitedocs(driver):
    try:
        print("Firefox Headless Browser Invoked")
        print("Logging in to SiteDocs...")
        driver.get(config.URL)
        
        username_field = driver.find_element(By.XPATH, '//*[@id="Username"]')
        username_field.send_keys(config.USERNAME)

        next_button = driver.find_element(By.XPATH, "/html/body/div/div/div/div[2]/div/form/fieldset/button")
        next_button.click()
        
        username_field = driver.find_element(By.XPATH, '//*[@id="Password"]')
        username_field.send_keys(config.PWD)

        next_button = driver.find_element(By.XPATH, "/html/body/div/div/div/div[2]/div/form/fieldset/button")
        next_button.click()

        print("Log in successful!")
    except:
        driver.quit()
        print("Unable to log in")

# Check email
# Check time
# Check attachment file
# Check number of pages
# Check number of matches (1 per day)


if __name__ == '__main__':
    process_flhas()
