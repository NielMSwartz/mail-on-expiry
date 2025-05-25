import pandas as pd
import datetime
import smtplib
from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout

class MyButton(App):
    def build(self):
        layout = BoxLayout(orientation="vertical")
        
        medical_button = Button(text="Medicals")
        cert_button = Button(text="Certificates")
        tool_button = Button(text="Tools & Vehicles")
        
        medical_button.bind(on_press=lambda instance: send_medical_mail())
        cert_button.bind(on_press=lambda instance: send_cerificate_mail())
        tool_button.bind(on_press=lambda instance: send_tool_mail())
        
        layout.add_widget(medical_button)
        layout.add_widget(cert_button)
        layout.add_widget(tool_button)
        return layout
    

SENDER = "nielswartz11@gmail.com"
RECEIVER = "nielswartz7@gmail.com"
password = pd.read_csv(f"C:/Users/Niel/Documents/cert_expiry_project/cert_expiry_reminder/password.csv")
PASSWORD = password["password"][0]

today = datetime.datetime.now().date()
medicals = pd.read_excel("C:/Users/Niel/Documents/cert_expiry_project/cert_expiry_reminder/exp_dates.xlsx", sheet_name="Medicals")

certificates = pd.read_excel("C:/Users/Niel/Documents/cert_expiry_project/cert_expiry_reminder/exp_dates.xlsx", sheet_name="Certificates")

tools = pd.read_excel("C:/Users/Niel/Documents/cert_expiry_project/cert_expiry_reminder/exp_dates.xlsx", sheet_name="Tools")

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()

medical_message = "Medicals to keep an eye on: \n\n"
certificates_message = "Certificates to keep an eye on: \n\n"
tools_message = "Tools to keep an eye on: \n\n"

for _, row in medicals.iterrows():
    name = row["names"]
    expiry_date = row["exp_date"].date()
    
    days_to_expire = (expiry_date - today).days
    
    medical_message
    
    if days_to_expire < 0:
        medical_message += "Expired items: \n"
        medical_message += f"\t{name} expired on {expiry_date}\n"
        
    elif 1 <= days_to_expire <= 7:
        medical_message += "Items that are expiring in 7 days: \n"
        medical_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"
        
    elif 8 <= days_to_expire <= 30:
        medical_message += "Items that are expiring in 30 days: \n"
        medical_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"
        
    elif 31 <= days_to_expire <= 60:
        medical_message += "Items that are expiring in 60 days: \n"
        medical_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"
        
for _, row in certificates.iterrows():
    name = row["names"]
    expiry_date = row["exp_date"].date()
    
    days_to_expire = (expiry_date - today).days
    
    certificates_message
    
    if days_to_expire < 0:
        certificates_message += "Expired items: \n"
        certificates_message += f"\t{name} expired on {expiry_date}\n"
        
    elif 1 <= days_to_expire <= 7:
        certificates_message += "Items that are expiring in 7 days: \n"
        certificates_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"
        
    elif 8 <= days_to_expire <= 30:
        certificates_message += "Items that are expiring in 30 days: \n"
        certificates_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"
        
    elif 31 <= days_to_expire <= 60:
        certificates_message += "Items that are expiring in 60 days: \n"
        certificates_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"
        
for _, row in tools.iterrows():
    name = row["names"]
    expiry_date = row["exp_date"].date()
    
    days_to_expire = (expiry_date - today).days
    
    tools_message
    
    if days_to_expire < 0:
        tools_message += "Expired items: \n"
        tools_message += f"\t{name} expired on {expiry_date}\n"
        
    elif 1 <= days_to_expire <= 7:
        tools_message += "Items that are expiring in 7 days: \n"
        tools_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"
        
    elif 8 <= days_to_expire <= 30:
        tools_message += "Items that are expiring in 30 days: \n"
        tools_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"
        
    elif 31 <= days_to_expire <= 60:
        tools_message += "Items that are expiring in 60 days: \n"
        tools_message += f"\t{name} expires on {expiry_date} in {days_to_expire} days. \n"

def send_medical_mail():
    try:
        server.login(SENDER, PASSWORD)
        print("Logged in.")
        server.sendmail(SENDER, RECEIVER, medical_message)
        print("Sent")
        print(medical_message)
    except smtplib.SMTPAuthenticationError:
        print("Unable to sign in.")
        
def send_cerificate_mail():
    try:
        server.login(SENDER, PASSWORD)
        print("Logged in.")
        server.sendmail(SENDER, RECEIVER, certificates_message)
        print("Sent")
        print(certificates_message)
    except smtplib.SMTPAuthenticationError:
        print("Unable to sign in.") 
        
def send_tool_mail():
    try:
        server.login(SENDER, PASSWORD)
        print("Logged in.")
        server.sendmail(SENDER, RECEIVER, tools_message)
        print("Sent")
        print(tools_message)
    except smtplib.SMTPAuthenticationError:
        print("Unable to sign in.")
    
        
if __name__ == "__main__":
    MyButton().run()

server.quit()
