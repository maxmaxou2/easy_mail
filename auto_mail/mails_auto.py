class Receivers :
    def __init__(self, dataframe, nb_receivers, nb_columns) :
        self.dataframe = dataframe
        self.nb_receivers = nb_receivers
        self.nb_columns = nb_columns

    def set_message(self, message: str) :
        self.message = message

    def format_message(self, index) :
        a = ""+self.message
        tab = self.dataframe.iloc[index]
        for i in range(self.nb_columns) :
            label = self.dataframe.columns[i]
            print("Index :", i, " and label:", label)
            if i != 0 :
                a = a.replace(label, tab[i])
        return a

class Data :
    def __init__(self, sender_email="", sender_name="", sender_password="", signature="", subject="", path="", message="", mail_type="0", excel_path="./annuaire.xlsx", cci="charles.barthelemy@student-cs.fr,andres.calderas@student-cs.fr"):
        self.sender_email = sender_email
        self.sender_name = sender_name
        self.sender_password = sender_password
        self.signature = signature
        self.subject = subject
        self.path = path
        self.message = message
        self.mail_type = mail_type
        self.excel_path = excel_path
        self.cci = cci

    def set(self, key, val) :
        match key :
            case "sender_email" :
                self.sender_email = val
            case "sender_name" :
                self.sender_name = val
            case "sender_password" :
                self.sender_password = val
            case "signature" :
                self.signature = val
            case "subject" :
                self.subject = val
            case "path" :
                self.path = val
            case "message" :
                self.message = val
            case "mail_type" :
                self.mail_type = val
            case "excel_path" :
                self.excel_path = val
            case "cci" :
                self.cci = val
            


import smtplib
from tkinter import *
from tkinter import ttk
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd

def load_data() :
    f = open("./auto_mail/config.txt", "r")
    data_str = "".join(f.readlines())
    f.close()
    data = Data()
    if data_str.strip() == "" :
        return data
    tab = data_str.split("<!!!>")
    for line in tab :
        line_tab = line.split("=", 1)
        data.set(line_tab[0].strip(), line_tab[1].strip())
    return data

def load_receivers(excel_path) :
    df = pd.read_excel(excel_path)
    nb_columns = df.columns.size
    nb_index = df.index.size
    return Receivers(df, nb_index, nb_columns)

def on_closing() :
    # Récupérer les informations saisies par l'utilisateur
    sender_email_type = listeCombo_type_mail.get()
    sender_email = sender_entry.get()
    sender_password = password_entry.get()
    subject = subject_entry.get()
    message = message_text.get("1.0", END)
    sender_name = sender_name_entry.get()
    signature = sender_signature_entry.get()
    path_to_attached_file = path_to_attached_file_entry.get()
    excel_path = excel_path_entry.get()
    cci = cci_entry.get()

    data = Data(sender_email, sender_name, sender_password, signature, subject, path_to_attached_file, message, ("0" if sender_email_type == "CentraleSupélec" else "1"), excel_path, cci)

    f = open("./auto_mail/config.txt", "w")
    to_write = "<!!!>".join(("sender_email="+data.sender_email, "sender_name="+data.sender_name, "sender_password="+data.sender_password, "signature="+data.signature, "subject="+data.subject, "path="+data.path, "message="+data.message, "mail_type="+data.mail_type, "excel_path="+data.excel_path, "cci="+data.cci))
    f.write(to_write)
    f.close()
    window.destroy()
    return

def set_text(e, text):
    e.delete(0,END)
    e.insert(0,text)
    return

def format(text):
    return text.replace("\n", "<br>").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")

def send_email():
    # Récupérer les informations saisies par l'utilisateur
    sender_name = sender_name_entry.get()
    sender_email_type = listeCombo_type_mail.get()
    sender_email = sender_entry.get()
    sender_password = password_entry.get()
    subject = subject_entry.get()
    message = message_text.get("1.0", END)
    signature = sender_signature_entry.get()
    path_to_attached_file = path_to_attached_file_entry.get()
    excel_path = excel_path_entry.get()
    cci = cci_entry.get()

    receivers = load_receivers(excel_path)

    message = message.replace("{Prénom et nom}", sender_name)
    receivers.set_message(message)

    for i in range(receivers.nb_receivers) :
        message = receivers.format_message(i)

        msg = MIMEMultipart('alternative')
        msg['From'] = sender_email
        msg['To'] = receivers.dataframe.iloc[i][0]
        msg['Subject'] = subject
        msg['Bcc'] = ", ".join([mail.strip() for mail in cci.split(",")])

        with open(path_to_attached_file, 'rb') as f:
            attach = MIMEApplication(f.read(),_subtype='pdf')
            attach.add_header('Content-Disposition','attachment',filename=str(path_to_attached_file.split("/")[-1]))
            msg.attach(attach)

        msg.attach(MIMEText(format(message)+"<br>"+"<html>"+signature+"</html>", 'html'))

        # Se connecter au serveur SMTP de l'expéditeur
        with smtplib.SMTP("smtp.gmail.com" if sender_email_type == "Gmail" else "smtp.office365.com", 587) as smtp:
            smtp.starttls()
            smtp.login(sender_email, sender_password)
            
            # Envoyer le message à l'adresse e-mail du destinataire
            smtp.send_message(msg)
        
    # Afficher une notification de confirmation
    confirmation_label.config(text="E-mail(s) envoyé(s) avec succès!")

#Charger les données sauvegardées en config
data = load_data()

# Créer une interface graphique
window = Tk()
window.title("Envoyer un e-mail")

sender_name_label = Label(window, text="{Prénom et nom}:")
sender_name_label.pack()
sender_name_entry = Entry(window)
set_text(sender_name_entry, data.sender_name)
sender_name_entry.pack()

label_liste_type_mail = Label(window, text = "Type de boîte mail d'envoi:")
label_liste_type_mail.pack()
liste_type_mails=["CentraleSupélec", "Gmail"]
listeCombo_type_mail = ttk.Combobox(window, values=liste_type_mails)
listeCombo_type_mail.current(int(data.mail_type))
listeCombo_type_mail.pack()

# Créer des champs de saisie pour l'expéditeur et le destinataire
sender_label = Label(window, text="Adresse e-mail de l'expéditeur:")
sender_label.pack()
sender_entry = Entry(window)
set_text(sender_entry, data.sender_email)
sender_entry.pack()

password_label = Label(window, text="Mot de passe:")
password_label.pack()
password_entry = Entry(window, show="*")
set_text(password_entry, data.sender_password)
password_entry.pack()

cci_label = Label(window, text="CCI")
cci_label.pack()
cci_entry = Entry(window)
set_text(cci_entry, data.cci)
cci_entry.pack()

sender_signature_label = Label(window, text="Code signature corpo:")
sender_signature_label.pack()
sender_signature_entry = Entry(window)
set_text(sender_signature_entry, data.signature)
sender_signature_entry.pack()

# Créer des champs de saisie pour le sujet et le message
subject_label = Label(window, text="Sujet:")
subject_label.pack()
subject_entry = Entry(window)
set_text(subject_entry, data.subject)
subject_entry.pack()

path_to_attached_file_label = Label(window, text="Chemin de la pièce jointe:")
path_to_attached_file_label.pack()
path_to_attached_file_entry = Entry(window)
set_text(path_to_attached_file_entry, data.path)
path_to_attached_file_entry.pack()

excel_path_label = Label(window, text="Chemin de l'annuaire excel:")
excel_path_label.pack()
excel_path_entry = Entry(window)
set_text(excel_path_entry, data.excel_path)
excel_path_entry.pack()

# Créer un bouton pour envoyer l'e-mail
send_button = Button(window, text="Envoyer", command=send_email)
send_button.pack()

# Créer une étiquette pour afficher la confirmation de l'envoi
confirmation_label = Label(window, text="")
confirmation_label.pack()

message_label = Label(window, text="Message:")
message_label.pack()
message_text = Text(window)
message_text.delete(1.0, END)
message_text.insert(END, data.message)
message_text.pack(fill="x")

# Afficher la fenêtre
window.protocol("WM_DELETE_WINDOW", on_closing)
window.mainloop()