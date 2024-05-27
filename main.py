from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import time
import pandas as pd
import smtplib


options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                          options=options)

#Accessing Data
driver.get("https://my.fibank.bg/EBank/public/offices")
driver.maximize_window()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, ".//*[contains(text(), 'Всички клонове')]"))
)

input_element = driver.find_element(By. CLASS_NAME, "class=btn-group bootstrap-select form-control dropdown ng-pristine ng-untouched ng-valid ng-scope open")
input_element.send_keys("class=btn-group bootstrap-select form-control dropdown ng-pristine ng-untouched ng-valid ng-scope open" + Keys.ENTER)

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'С удължени раб. време')]"))
)

input_element = driver.find_element(By.CLASS_NAME, "span.class=filter-option" + Keys.ENTER) 

data = []
for office in offices:
    name = office.find_element(By.CLASS_NAME, 'office-name').text
    address = office.find_element(By.CLASS_NAME, 'office-address').text
    phone = office.find_element(By.CLASS_NAME, 'office-phone').text
    
    try:
        saturday_hours = office.find_element(By.XPATH, ".//*[contains(text(), 'събота')]").text.split(': ')[1]
    except:
        saturday_hours = 'N/A'
    
    try:
        sunday_hours = office.find_element(By.XPATH, ".//*[contains(text(), 'неделя')]").text.split(': ')[1]
    except:
        sunday_hours = 'N/A'
    
    data.append({
        'Office Name': name,
        'Address': address,
        'Phone': phone,
        'Saturday Hours': saturday_hours,
        'Sunday Hours': sunday_hours
    })

#Save data to Excel 
df = pd.DataFrame(data)
file_path = 'C:/PythonApp/fibank_branches.xlsx'
df.to_excel(file_path, index=False)

#Sending Email
def send_email(subject, body, to_email, attachment_path):
    from_email = "my_email@gmail.com"
    from_password = "my_password"

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {C:\PythonApp}fibank_branches")
        msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, from_password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()

send_email(
    subject="Fibank Branches Data",
    body="Please find the attached Excel file with Fibank branches working hours and working on weekends.",
    to_email="db.rpa@fibank.bg",
    attachment_path=file_path
)

print("Email sent successfully.")

time.sleep(10)

driver.quit()