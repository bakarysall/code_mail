import streamlit as st
import pandas as pd
import smtplib 
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Code pour récupérer les infos de connexion mis dans le sidebar
def user_auth():
    image = st.sidebar.image('logo.jpg')
    st.sidebar.title("Entrez vos informations d'identification Gmail")
    email = st.sidebar.text_input("E-mail")
    password = st.sidebar.text_input("Mot de passe", type="password")
    return email, password

# Information du compte qui va envoyer le mail
def send_email(sender_email, password, receiver_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Connexion au niveau du compte Gmail via le serveur SMTP sur le port SSL (port 465)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl.create_default_context()) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, msg.as_string())

# Récupération des informations dans le fichier Excel
def send_student_notes(df, sender_email, password, nom_de_matiere):
    for index, row in df.iterrows():
        student_name = row['Nom']
        student_email = row['Email']
        student_not_dev = row['Note_dev']
        student_not_ex = row['Note_ex']
        student_not_moy = row['Moyenne']

        # Voici le corps du message avec les variables en haut qui récupèrent à partir du fichier Excel
        message_body = f"Cher {student_name},\n\nVotre note de 40% est : {student_not_dev},\n\nVotre note de 60% est : {student_not_ex},\n\nVotre note de Moyenne est : {student_not_moy},\n\nCordialement,\nSALL"
        send_email(sender_email, password, student_email, f"Résultats des notes de {nom_de_matiere}", message_body)

# Nouvelle fonction pour envoyer les emplois du temps
def send_class_schedule(df, sender_email, password, nom_de_matiere, heure_debut, heure_fin, nom_prof, classe_cours):
    for index, row in df.iterrows():
        student_name = row['Nom']
        student_email = row['Email']
        class_schedule = f"Matière : {nom_de_matiere}\nClasse : {classe_cours}\nHeure de début : {heure_debut}\nHeure de fin : {heure_fin}\nProfesseur : {nom_prof}"
        message_body = f"Cher {student_name},\n\nVoici votre emploi du temps :\n\n{class_schedule}\n\nCordialement,\nSALL"
        send_email(sender_email, password, student_email, f"Emploi du temps de {nom_de_matiere}", message_body)

def main():
    st.sidebar.title('Bienvenue cher Gouverneur')
    st.title("Envoi des résultats par e-mail")

    # Appel de la fonction user_auth pour la connexion via le compte de l'envoyeur
    sender_email, password = user_auth()

    st.sidebar.write("L'application vous permet d'envoyer les notes ou les emplois du temps via mail. Pour l'utiliser, vous devez charger le classeur Excel. Vérifiez que le fichier est un fichier Excel. Ensuite, entrez le nom de la matière qui sera considéré comme l'objet de votre mail. Merci.")

    # Bouton pour uploader un fichier via l'ordinateur
    excel_file = st.file_uploader("Téléchargez le fichier Excel", type=["xlsx"])

    # Chargement du fichier Excel
    if excel_file is not None:
        df = pd.read_excel(excel_file)

        # Récupérer le nom de la matière pour en faire l'objet
        nom_de_matiere = st.text_input("Donner le nom de la matière")

        # Ajout d'options pour choisir entre envoyer les notes ou les emplois du temps
        choix_action = st.radio("Choisissez l'action à effectuer :", ["Envoyer les notes", "Envoyer les emplois du temps"])

        if choix_action == "Envoyer les notes":
            if st.button("Envoyer les e-mails"):
                # D'abord, on vérifie
                if nom_de_matiere:
                    send_student_notes(df, sender_email, password, nom_de_matiere)
                    st.success("Les e-mails des notes ont été envoyés avec succès.")
                # Vérification pour voir si l'utilisateur a renseigné le nom de la matière
                else:
                    st.warning("Petit saisissez d'abord le nom de la matière.")

        elif choix_action == "Envoyer les emplois du temps":
            heure_debut = st.text_input("Heure de début du cours")
            heure_fin = st.text_input("Heure de fin du cours")
            nom_prof = st.text_input("Nom du professeur")
            classe_cours= st.text_input("Entrer le nom de la classe")

            if st.button("Envoyer les e-mails"):
                # D'abord, on vérifie
                if nom_de_matiere and heure_debut and heure_fin and nom_prof:
                    send_class_schedule(df, sender_email, password, nom_de_matiere, heure_debut, heure_fin, nom_prof,classe_cours)
                    st.success("Les e-mails des emplois du temps ont été envoyés avec succès.")
                # Vérification pour voir si l'utilisateur a renseigné toutes les informations nécessaires
                else:
                    st.warning("Veuillez saisir toutes les informations nécessaires pour les emplois du temps.")
    
if __name__ == '__main__':
    main()
