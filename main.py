import pandas as pd
import datetime
import smtplib

SENDER = "nielswartz11@gmail.com"
RECEIVER = "nielswartz7@gmail.com"
password = pd.read_csv(f"cert_expiry_reminder/password.csv")
PASSWORD = password["password"][0]
SUBJECT = "Expiry Dates!"

today = datetime.datetime.now()

df = pd.read_excel("C:/Users/Niel/Documents/practice/exp_dates.xlsx")

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()

for _,rows in df.iterrows():
    names = rows["names"]
    exp_date = rows["exp_date"]
    MESSAGE = f"{names} Expires on: {exp_date.date()}"
    
    if exp_date < today:
        print(f"The {names} has already expired on {exp_date}")
    
    elif exp_date <= today + datetime.timedelta(7):
        print(f"{names} expires within 7 days on: {exp_date.date()}")
        try:
            server.login(SENDER, PASSWORD)
            print("Logged in.")
            server.sendmail(SENDER, RECEIVER, MESSAGE)
            print("Sent")
        except smtplib.SMTPAuthenticationError:
            print("Unable to sign in.")
    
    elif exp_date <= today + datetime.timedelta(30):
        print(f"{names} expires within 30 days on: {exp_date.date()}")
        try:
            server.login(SENDER, PASSWORD)
            print("Logged in.")
            server.sendmail(SENDER, RECEIVER, MESSAGE)
            print("Sent")
        except smtplib.SMTPAuthenticationError:
            print("Unable to sign in.")
            
    elif exp_date <= today + datetime.timedelta(60):
        print(f"{names} expires within 60 days on: {exp_date.date()}")
        try:
            server.login(SENDER, PASSWORD)
            print("Logged in.")
            server.sendmail(SENDER, RECEIVER, MESSAGE)
            print("Sent")
        except smtplib.SMTPAuthenticationError:
            print("Unable to sign in.")

server.quit()

