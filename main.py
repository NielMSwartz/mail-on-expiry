import pandas as pd
import datetime
import smtplib

SENDER = "nielswartz11@gmail.com"
RECEIVER = "nielswartz7@gmail.com"
password = pd.read_csv(f"C:/Users/Niel/Documents/Personal Projects/cert_expiry_reminder/password.csv")
PASSWORD = password["password"][0]
SUBJECT = "Expiry Dates!"
MESSAGE = "Hellow"

today = datetime.datetime.now().date()

df = pd.read_excel("C:/Users/Niel/Documents/practice/exp_dates.xlsx")

names = df["names"]
exp_date = df["exp_date"]

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()

expired = []
expires_in_7 = []
expires_in_30 = []
expires_in_60 = []

for _, row in df.iterrows():
    name = row["names"]
    expiry_date = row["exp_date"].date()
    
    days_to_expire = (expiry_date - today).days
    
    if days_to_expire < 0:
        expired.append([name, expiry_date])
    
    elif 1 <= days_to_expire <= 7:
        expires_in_7.append([name, expiry_date])
    
    elif 8 <= days_to_expire <= 31:
        expires_in_30.append([name, expiry_date])
        
    elif 32 <= days_to_expire <= 60:
        expires_in_60.append([name, expiry_date])
    
df_expired_items = pd.DataFrame(expired, columns=["Items", "Expiry_date"])
df_expires_in_7 = pd.DataFrame(expires_in_7, columns=["Items", "Expiry_date"])
df_expires_in_30 = pd.DataFrame(expires_in_30, columns=["Items", "Expiry_date"])
df_expires_in_60 = pd.DataFrame(expires_in_60, columns=["Items", "Expiry_date"])

df_expired_items.to_csv("Expired_items.csv")
df_expires_in_7.to_csv("Expires_in_7.csv")
df_expires_in_30.to_csv("Expires_in_30.csv")
df_expires_in_60.to_csv("Expires_in_60.csv")

try:
    df_expired_items = pd.read_csv("Expired_items.csv")
    df_expires_in_7 = pd.read_csv("Expires_in_7.csv")
    df_expires_in_30 = pd.read_csv("Expires_in_30.csv")
    df_expires_in_60 = pd.read_csv("Expires_in_60.csv")
except FileNotFoundError as e:
    print(f"Error: csv file not found {e}")
    exit(1)

email_body = "Items to keep an eye on: \n\n"

if not df_expired_items.empty:
    email_body += "The following items are already expired: \n"
    for _, row in df_expired_items.iterrows():
        item = row["Items"]
        item_expired_date = row["Expiry_date"]
        email_body += f"{item}: {item_expired_date}\n"
    print(email_body)
else:
    email_body += "No Items are Expired\n"
    
if not df_expires_in_7.empty:
    email_body += "The following items expire in 7 days: \n"
    for _, row in df_expires_in_7.iterrows():
        item_7 = row["Items"]
        item_expire_in_7_date = row["Expiry_date"]
        email_body += f"{item_7}: {item_expire_in_7_date}\n"
    print(email_body)
else:
    email_body += "No items 7 days from expiry\n"
    
if not df_expires_in_30.empty:
    email_body += "The following items expire in 30 days: \n"
    for _, row in df_expires_in_30.iterrows():
        item_30 = row["Items"]
        item_expire_in_30_date = row["Expiry_date"]
        email_body += f"{item_30}: {item_expire_in_30_date}\n"
    print(email_body)
else:
    email_body += "No items 30 days from expiry\n"
    
if not df_expires_in_60.empty:
    email_body += "The following items expire in 60 days: \n"
    for _, row in df_expires_in_60.iterrows():
        item_60 = row["Items"]
        item_expire_in_60_date = row["Expiry_date"]
        email_body += f"{item_60}: {item_expire_in_60_date}\n"
    print(email_body)
else:
    email_body += "No items 60 days from expiry\n"


try:
    server.login(SENDER, PASSWORD)
    print("Logged in.")
    server.sendmail(SENDER, RECEIVER, email_body)
    print("Sent")
except smtplib.SMTPAuthenticationError:
    print("Unable to sign in.")

server.quit()
