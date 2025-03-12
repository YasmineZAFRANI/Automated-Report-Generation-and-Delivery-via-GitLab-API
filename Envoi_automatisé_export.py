import smtplib
import os
import pandas as pd
import datetime
import holidays
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
#  https://support.google.com/mail/?p=InvalidSecondFactor necessaire pour la double authentification
# Configurations
EMAIL_ADDRESS = "adresse@gmail.com"
EMAIL_PASSWORD = "app_mdp"
TO_EMAIL = "adresse@gmail.com"
SUBJECT = "Envoi automatique du fichier Excel"
BODY = "Bonjour,\n\nVeuillez trouver ci-joint le rapport des tickets de cette semaine.\n\nCordialement."
FILE_PATH = r"C:\Users\yzafrani\issues_export_2025-02-18.xlsx"


# Définir le jour d'envoi
fr_holidays = holidays.France()
today = datetime.date.today()

# Trouver le prochain mercredi
send_date = today + datetime.timedelta((2 - today.weekday()) % 7)

# Si jour férié, envoyer la veille
while send_date in fr_holidays:
    send_date -= datetime.timedelta(days=1)

# Vérifier si aujourd'hui est le jour d'envoi
if today == send_date:
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = TO_EMAIL
        msg['Subject'] = SUBJECT
        msg.attach(MIMEText(BODY, 'plain'))

        # Attacher le fichier
        attachment = open(FILE_PATH, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(FILE_PATH)}")
        msg.attach(part)
        attachment.close()

        # Connexion SMTP et envoi
    #Pour Outlook    server = smtplib.SMTP('smtp.office365.com', 587)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        text = msg.as_string()
        server.sendmail(EMAIL_ADDRESS, TO_EMAIL, text)
        server.quit()

        print("Email envoyé avec succès!")
    except Exception as e:
        print(f"Erreur lors de l'envoi : {e}")


