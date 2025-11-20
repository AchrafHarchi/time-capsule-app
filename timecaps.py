import streamlit as st
from datetime import datetime
import random
import gspread
from google.oauth2.service_account import Credentials
import os
import pickle
import google_auth_oauthlib.flow
import googleapiclient.discovery
import googleapiclient.errors
import json
import pandas as pd

import io
import string
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email import encoders
import base64
import tempfile
import time
from googleapiclient.http import MediaFileUpload

# ---------- CONFIG + AUTH (lecture depuis Streamlit secrets) ----------
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import json
import base64
import tempfile
import os

# --- Defaults (comme dans ton ancien code) ---
DEFAULT_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
DEFAULT_YOUTUBE_SCOPES = ["https://www.googleapis.com/auth/youtube.upload"]

# ---------- SHEETS IDS ----------
# lire depuis le TOML -> [sheets]
INFO_SHEET_ID = st.secrets["sheets"]["info_sheet_id"]
DIST_SHEET_ID = st.secrets["sheets"]["dist_sheet_id"]

# ---------- SCOPES ----------
# Optionnel : tu peux définir une clé 'scopes' dans st.secrets si tu veux override
SCOPES = st.secrets.get("scopes", {}).get("list", DEFAULT_SCOPES)

# ---------- SERVICE ACCOUNT (remplace credentials.json) ----------
# Exige une section [google_service_account] dans ton TOML (format JSON-like)
service_account_info = dict(st.secrets["google_service_account"])
# On crée les creds directement depuis le dict (plus besoin de credentials.json)
creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)

# Si tu avais besoin d'un fichier credentials.json (pour une lib qui exige un chemin),
# on peut écrire un fichier temporaire et donner son chemin :
CREDENTIALS_FILE = None
try:
    tmp_dir = tempfile.gettempdir()
    tmp_cred_path = os.path.join(tmp_dir, "credentials_from_secrets.json")
    with open(tmp_cred_path, "w", encoding="utf-8") as f:
        json.dump(service_account_info, f)
    CREDENTIALS_FILE = tmp_cred_path
except Exception:
    CREDENTIALS_FILE = None  # fallback, la plupart des libs fonctionneront avec `creds` objet

# ---------- GSpread authorisation ----------
gc = gspread.authorize(creds)
sheet = gc.open_by_key(INFO_SHEET_ID).sheet1
rows = sheet.get_all_values()
header = rows[0] if rows else []

dist_sheet = gc.open_by_key(DIST_SHEET_ID).sheet1
dist_rows = dist_sheet.get_all_values()
dist_header = dist_rows[0] if dist_rows else []

# ---------- YOUTUBE / GMAIL OAUTH client_secret.json replacement ----------
# Si tu as [youtube_oauth] ou [gmail_oauth] dans le TOML, on les récupère en dict
youtube_client_secret_info = st.secrets.get("youtube_oauth", None)
gmail_client_secret_info = st.secrets.get("gmail_oauth", None)

CLIENT_SECRET_FILE = None  # chemin temporaire vers client_secret.json si créé
if youtube_client_secret_info:
    try:
        tmp_youtube_secret = os.path.join(tmp_dir, "client_secret_youtube.json")
        with open(tmp_youtube_secret, "w", encoding="utf-8") as f:
            json.dump(dict(youtube_client_secret_info), f)
        CLIENT_SECRET_FILE = tmp_youtube_secret
    except Exception:
        CLIENT_SECRET_FILE = None

# Tu peux aussi exposer le dict directement (pour les flows OAuth qui acceptent un dict)
YOUTUBE_CLIENT_SECRET = youtube_client_secret_info
GMAIL_CLIENT_SECRET = gmail_client_secret_info

# ---------- YOUTUBE SCOPES ----------
YOUTUBE_SCOPES = st.secrets.get("youtube_scopes", {}).get("list", DEFAULT_YOUTUBE_SCOPES)

# ---------- TOKENS (token.pkl / token_gmail.pkl) ----------
# Attendu dans TOML : [tokens] youtube = "<base64 string>" , gmail = "<base64 string>"
tokens_section = st.secrets.get("tokens", {})

PICKLE_FILE = None
GMAIL_TOKEN_FILE = None

def _write_base64_to_temp(base64_str: str, filename: str) -> str:
    """Decode base64 and write bytes to a temp file. Return path or None."""
    if not base64_str:
        return None
    try:
        data = base64.b64decode(base64_str)
        path = os.path.join(tempfile.gettempdir(), filename)
        with open(path, "wb") as f:
            f.write(data)
        return path
    except Exception:
        return None

# youtube token (token.pkl)
youtube_b64 = tokens_section.get("youtube", "")
if youtube_b64:
    PICKLE_FILE = _write_base64_to_temp(youtube_b64, "token.pkl")

# gmail token (token_gmail.pkl)
gmail_b64 = tokens_section.get("gmail", "")
if gmail_b64:
    GMAIL_TOKEN_FILE = _write_base64_to_temp(gmail_b64, "token_gmail.pkl")

# Si les tokens n'existent pas dans st.secrets, PICKLE_FILE / GMAIL_TOKEN_FILE resteront None
# et tu pourras lancer le flow OAuth normal pour générer les fichiers et (éventuellement)
# convertir puis coller en base64 dans ton TOML.

# ---------- Résumé des variables exposées ----------
# INFO_SHEET_ID, DIST_SHEET_ID,
# CREDENTIALS_FILE (path temporaire si créé), creds (Credentials object),
# SCOPES,
# CLIENT_SECRET_FILE (path temporaire si créé),
# YOUTUBE_CLIENT_SECRET (dict), GMAIL_CLIENT_SECRET (dict),
# YOUTUBE_SCOPES,
# PICKLE_FILE (path to token.pkl if present in secrets),
# GMAIL_TOKEN_FILE (path to token_gmail.pkl if present in secrets)


# ---------- SESSION STATE ----------
if "current_step" not in st.session_state:
    st.session_state["current_step"] = 1
if "client_id" not in st.session_state:
    st.session_state["client_id"] = None
if "distribution_validated" not in st.session_state:
    st.session_state["distribution_validated"] = False

# ---------- SIDEBAR ----------
st.sidebar.title("Time Capsule Steps")

def step_label(step_number, label):
    if st.session_state["current_step"] > step_number or (
        step_number==3 and st.session_state["distribution_validated"]
    ):
        return f"✅ <span style='color:green'>{step_number}️⃣ {label}</span>"
    elif st.session_state["current_step"] == step_number:
        return f"🟡 {step_number}️⃣ {label}"
    else:
        return f"{step_number}️⃣ {label}"

st.sidebar.markdown(step_label(1, "Step 1: Client Info"), unsafe_allow_html=True)
st.sidebar.markdown(step_label(2, "Step 2: Upload Video"), unsafe_allow_html=True)
st.sidebar.markdown(step_label(3, "Step 3: Distribution List"), unsafe_allow_html=True)
st.sidebar.markdown(step_label(4, "Step 4: Conditions de Diffusion"), unsafe_allow_html=True)
st.sidebar.markdown(step_label(5, "Step 5: Revue et Validation"), unsafe_allow_html=True)
st.sidebar.markdown(step_label(6, "Step 6: Paiement & Envoi"), unsafe_allow_html=True)

# ---------- STEP 1 ----------
st.title("🌿 Time Capsule - Étape 1 : Informations Client")
with st.form("client_info_form"):
    prénom = st.text_input("Prénom", value=st.session_state.get("prénom", ""))
    nom = st.text_input("Nom", value=st.session_state.get("nom", ""))
    date_naissance = st.date_input(
        "Date de naissance",
        value=st.session_state.get("date_naissance", datetime(1980,1,1)),
        min_value=datetime(1900,1,1),
        max_value=datetime.today()
    )
    email = st.text_input("Email", value=st.session_state.get("email",""))
    téléphone = st.text_input("Téléphone", value=st.session_state.get("téléphone",""))
    adresse = st.text_input("Adresse", value=st.session_state.get("adresse",""))
    submitted = st.form_submit_button("Enregistrer")

if submitted:
    if not (prénom and nom and date_naissance):
        st.error("Veuillez remplir au minimum : prénom, nom et date de naissance.")
    else:
        st.session_state.update({
            "prénom": prénom,
            "nom": nom,
            "date_naissance": date_naissance,
            "email": email,
            "téléphone": téléphone,
            "adresse": adresse
        })

        client_id = None
        data_rows = rows[1:]
        for idx, row in enumerate(data_rows, start=2):
            if (row[header.index("prénom")].strip().lower() == prénom.strip().lower() and
                row[header.index("nom")].strip().lower() == nom.strip().lower() and
                row[header.index("date_naissance")].strip() == date_naissance.strftime("%d/%m/%Y")):
                client_id = row[header.index("client_ID")]
                st.info(f"Client existant trouvé. Client ID: {client_id}")
                break

        if client_id is None:
            client_id = str(random.randint(10000000,99999999))
            now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
            new_row = [
                client_id, prénom, nom, date_naissance.strftime("%d/%m/%Y"),
                email, téléphone, adresse, "", "", "en_attente", now_str, now_str
            ]
            sheet.append_row(new_row)
            st.success(f"Nouveau client enregistré. Client ID: {client_id}")

        # ⚡ Mise à jour step AVANT tout feedback pour corriger le décalage
        st.session_state["client_id"] = client_id
        st.session_state["current_step"] = 2

# ---------- STEP 2 ----------
if st.session_state["current_step"] >= 2 and st.session_state["client_id"]:
    client_id = st.session_state["client_id"]
    st.title("🌿 Time Capsule - Étape 2 : Upload Video")
    uploaded_file = st.file_uploader("Sélectionnez une vidéo", type=["mp4","mov","avi"])
    video_title = st.text_input("Titre de la vidéo")
    video_description = st.text_area("Description de la vidéo")

    if uploaded_file and st.button("Uploader sur YouTube"):
        temp_video_path = "temp_video.mp4"
        with open(temp_video_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # --- Auth YouTube ---
        credentials = None
        if os.path.exists(PICKLE_FILE):
            with open(PICKLE_FILE, "rb") as f:
                credentials = pickle.load(f)
        else:
            flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, YOUTUBE_SCOPES
            )
            credentials = flow.run_local_server(port=0)
            with open(PICKLE_FILE, "wb") as f:
                pickle.dump(credentials, f)

        youtube = googleapiclient.discovery.build("youtube", "v3", credentials=credentials)

        # --- Upload avec progress bar ---
        media = googleapiclient.http.MediaFileUpload(temp_video_path, chunksize=1024*1024, resumable=True)
        request = youtube.videos().insert(
            part="snippet,status",
            body={
                "snippet": {
                    "title": video_title or "Time Capsule Video",
                    "description": video_description or "",
                    "tags": ["time capsule"],
                    "categoryId": "22"
                },
                "status": {"privacyStatus": "unlisted"}
            },
            media_body=media
        )

        progress_bar = st.progress(0)
        status_text = st.empty()
        response = None

        try:
            while response is None:
                status, response = request.next_chunk()
                if status:
                    progress_bar.progress(int(status.progress() * 100))
                    status_text.text(f"Upload vidéo en cours : {int(status.progress() * 100)}%")
        finally:
            # ⚡ Fermer le flux pour libérer le fichier
            if media.stream():
                media.stream().close()
            # ⚡ Supprimer uniquement si le fichier existe
            if os.path.exists(temp_video_path):
                os.remove(temp_video_path)

        progress_bar.empty()
        status_text.text("Upload terminé ✅")

        # --- Mettre à jour le lien vidéo dans Google Sheets ---
        video_url = f"https://youtu.be/{response['id']}"
        st.success(f"Vidéo uploadée ! Voir ici : {video_url}")

        cell = sheet.find(client_id)
        sheet.update_cell(cell.row, header.index("lien_vidéo")+1, video_url)

        # Étape 2 terminée → passage à l'étape 3 immédiatement en vert
        st.session_state["current_step"] = 3

# ---------- STEP 3 ----------
if st.session_state["current_step"] >= 3 and st.session_state["client_id"]:
    client_id = st.session_state["client_id"]
    st.title("🌿 Time Capsule - Étape 3 : Liste de Distribution")

    with st.form("add_recipient_form"):
        prénom_dest = st.text_input("Prénom du destinataire")
        nom_dest = st.text_input("Nom du destinataire")
        email_dest = st.text_input("Email du destinataire")
        phone_dest = st.text_input("Téléphone (optionnel)")
        addr_dest = st.text_input("Adresse (optionnelle)")
        message_dest = st.text_area("Message")
        add_recipient = st.form_submit_button("Ajouter ce destinataire")

    if add_recipient:
        if prénom_dest and nom_dest and email_dest:
            now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
            new_row = [
                client_id, prénom_dest, nom_dest, email_dest, phone_dest, addr_dest, message_dest, now_str
            ]
            dist_sheet.append_row(new_row)
            st.success(f"Destinataire {prénom_dest} {nom_dest} ajouté !")
        else:
            st.error("Veuillez remplir au minimum prénom, nom et email du destinataire.")

    # Afficher les destinataires existants
    all_dests = dist_sheet.get_all_values()[1:]
    client_dests = [r for r in all_dests if r[0]==client_id]
    if client_dests:
        st.subheader("Destinataires ajoutés :")
        for idx, d in enumerate(client_dests, start=1):
            st.write(f"{idx}. {d[1]} {d[2]}, Email: {d[3]}")
            st.write(f"Message: {d[6]}")
            st.write("---")

    # Valider la liste de distribution
    if client_dests and st.button("Valider la liste finale"):
        st.session_state["distribution_validated"] = True
        st.session_state["current_step"] = 4  # passage immédiat à l'étape 4
        st.success("Liste de distribution validée ✅")


# ---------- STEP 4 ----------
if st.session_state["current_step"] >= 4 and st.session_state["client_id"]:
    st.title("🌿 Time Capsule - Étape 4 : Conditions de Diffusion")
    st.write("Définissez quand la vidéo sera libérée : après votre décès ou à une date donnée.")

    with st.form("conditions_form"):
        post_mortem = st.radio("Souhaitez-vous que la capsule soit diffusée après votre mort ?", ["Oui", "Non"])
        release_date = st.date_input("Date maximale de diffusion", min_value=datetime.today())
        submit_conditions = st.form_submit_button("Enregistrer les conditions")

    if submit_conditions:
        condition_data = {
            "post_mortem": True if post_mortem == "Oui" else False,
            "date": release_date.strftime("%Y-%m-%d")
        }
        condition_json = json.dumps(condition_data)

        client_id = st.session_state["client_id"]
        cell = sheet.find(client_id)
        sheet.update_cell(cell.row, header.index("condition")+1, condition_json)

        st.success("✅ Conditions enregistrées avec succès !")
        st.balloons()
        st.session_state["current_step"] = 5

# ---------- STEP 5 ----------
if st.session_state["current_step"] >= 5 and st.session_state["client_id"]:
    st.title("🌿 Time Capsule - Étape 5 : Revue et Validation des Informations")

    client_id = st.session_state["client_id"]
    client_row = next((r for r in rows if r and r[0] == client_id), None)
    if client_row:
        client_data = dict(zip(header, client_row))
        st.subheader("Informations du Client")
        st.table(pd.DataFrame(client_data.items(), columns=["Champ", "Valeur"]))

    video_link = client_data.get("lien_vidéo", "Non disponible")
    st.markdown(f"🎬 **Lien Vidéo :** [{video_link}]({video_link})" if video_link != "Non disponible" else "🎬 Vidéo non encore uploadée")

    client_dests = [r for r in dist_rows[1:] if r[0] == client_id]
    if client_dests:
        st.subheader("Liste de Distribution")
        st.table(pd.DataFrame(client_dests, columns=dist_header))

    if client_data.get("condition"):
        try:
            condition = json.loads(client_data["condition"])
            st.subheader("Conditions de Diffusion")
            st.json(condition)
        except:
            st.write("⚠️ Données de condition non valides.")

    if st.button("✅ Valider toutes les informations"):
        st.session_state["current_step"] = 6
        st.success("Toutes les informations ont été validées ! Passage à l’étape suivante.")


# ---------- STEP 6 ----------

def generate_access_code(length=8):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

def create_recipient_pdf(recipient, client_data, video_url, access_code, condition_text):
    fd, pdf_path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)

    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4

    # Header
    c.setFont("Helvetica-Bold", 16)
    c.drawString(30*mm, height - 30*mm, "Capsule Temporelle - Message pour destinataire")

    # Recipient info
    c.setFont("Helvetica", 11)
    y = height - 40*mm
    c.drawString(30*mm, y, f"Destinataire : {recipient.get('prénom','')} {recipient.get('nom','')}")
    y -= 7*mm
    c.drawString(30*mm, y, f"Email : {recipient.get('email','')}")
    y -= 7*mm
    if recipient.get('phone'):
        c.drawString(30*mm, y, f"Téléphone : {recipient.get('phone')}")
        y -= 7*mm
    if recipient.get('addr'):
        c.drawString(30*mm, y, f"Adresse : {recipient.get('addr')}")
        y -= 7*mm

    # Message
    y -= 4*mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(30*mm, y, "Message du client :")
    y -= 6*mm
    c.setFont("Helvetica", 10)
    text = c.beginText(30*mm, y)
    for line in (recipient.get('message') or "").splitlines():
        text.textLine(line)
    c.drawText(text)

    # Access code & condition
    bottom_y = 70*mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(30*mm, bottom_y, f"Code d'accès : {access_code}")
    c.setFont("Helvetica", 10)
    c.drawString(30*mm, bottom_y - 7*mm, f"Condition de diffusion : {condition_text}")

    # QR Code centered
    qr = qrcode.QRCode(box_size=6, border=2)
    qr.add_data(video_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    qr_fd, qr_path = tempfile.mkstemp(suffix=".png")
    os.close(qr_fd)
    img.save(qr_path)
    qr_size = 60*mm
    c.drawImage(qr_path, (width - qr_size)/2, (height - qr_size)/2, width=qr_size, height=qr_size)

    # Footer
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(30*mm, 20*mm, "Pour visionner la vidéo, rendez-vous sur l'application Time Capsule et entrez votre code d'accès.")
    c.showPage()
    c.save()

    try: os.remove(qr_path)
    except: pass

    return pdf_path

def create_client_pdf(client_data, client_dests, condition_text):
    fd, pdf_path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)

    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4

    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(width/2, height - 30*mm, "Capsule Temporelle - Récapitulatif Client")

    c.setFont("Helvetica", 11)
    y = height - 45*mm
    c.drawString(20*mm, y, f"Client : {client_data.get('prénom','')} {client_data.get('nom','')}")
    y -= 7*mm
    c.drawString(20*mm, y, f"Email : {client_data.get('email','')}")
    y -= 7*mm
    c.drawString(20*mm, y, f"Date de naissance : {client_data.get('date_naissance','')}")
    y -= 10*mm

    c.setFont("Helvetica-Bold", 12)
    c.drawString(20*mm, y, "Liste des destinataires :")
    y -= 8*mm
    c.setFont("Helvetica", 10)
    for idx, d in enumerate(client_dests, start=1):
        line = f"{idx}. {d.get('prénom_dest','')} {d.get('nom_dest','')} - {d.get('email_dest','')} - Code: {d.get('access_code','')}"
        c.drawString(22*mm, y, line)
        y -= 6*mm
        if y < 40*mm:
            c.showPage()
            y = height - 30*mm

    y -= 6*mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(20*mm, y, "Conditions de diffusion :")
    y -= 7*mm
    c.setFont("Helvetica", 10)
    c.drawString(20*mm, y, condition_text)

    c.showPage()
    c.save()
    return pdf_path

def make_mime_attachment(file_path, filename):
    with open(file_path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
    part.add_header('Content-Disposition', 'attachment', filename=filename)
    return part

def send_email_with_attachments_gmail(subject, body, to_email, attachments_paths, creds_file=GMAIL_TOKEN_FILE):
    with open(creds_file, "rb") as f:
        creds = pickle.load(f)
    service = googleapiclient.discovery.build("gmail", "v1", credentials=creds)
    message = MIMEMultipart()
    message["to"] = to_email
    message["from"] = "me"
    message["subject"] = subject
    message.attach(MIMEText(body, "plain"))
    for p in attachments_paths:
        part = make_mime_attachment(p["path"], p["name"])
        message.attach(part)
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    send_message = service.users().messages().send(userId="me", body={"raw": raw}).execute()
    return send_message

# UI: display payer button
if st.session_state["current_step"] >= 6 and st.session_state.get("client_id"):
    client_id = st.session_state["client_id"]
    st.title("🌿 Time Capsule - Étape 6 : Paiement & Envoi des Fiches (DEMO)")

    st.markdown("""
    Cliquez sur **Payer** pour simuler la confirmation du paiement et lancer l'envoi des documents.
    (Démo : aucun paiement réel n'est effectué.)
    """)
    pay_col1, pay_col2 = st.columns([1,3])
    with pay_col1:
        pay = st.button("💳 Payer", key="pay_button")
    with pay_col2:
        st.markdown("<div style='padding:10px;border-radius:8px'>Bouton de paiement (démo). Après clic, génération des PDF et envoi d'emails.</div>", unsafe_allow_html=True)


    if pay:
        try:
            rows = sheet.get_all_values()
            header = rows[0]
            client_row = next((r for r in rows if r and r[0] == client_id), None)
            if not client_row:
                st.error("Client introuvable.")
                st.stop()
            client_data = dict(zip(header, client_row))

            dist_rows = dist_sheet.get_all_values()
            dist_header = dist_rows[0]
            client_dests_rows = [r for r in dist_rows[1:] if r and r[0] == client_id]

            client_dests = []
            for r in client_dests_rows:
                rec = {
                    "client_ID": r[0],
                    "prénom_dest": r[1] if len(r) > 1 else "",
                    "nom_dest": r[2] if len(r) > 2 else "",
                    "email_dest": r[3] if len(r) > 3 else "",
                    "phone_dest": r[4] if len(r) > 4 else "",
                    "addr_dest": r[5] if len(r) > 5 else "",
                    "message": r[6] if len(r) > 6 else "",
                    "date_création": r[7] if len(r) > 7 else ""
                }
                client_dests.append(rec)

            condition_text = "Non définie"
            if client_data.get("condition"):
                try:
                    condition = json.loads(client_data["condition"])
                    condition_text = f"Post-mortem: {condition.get('post_mortem')} - Date max: {condition.get('date')}"
                except:
                    condition_text = str(client_data.get("condition"))

            attachments = []
            created_files = []
            for rec in client_dests:
                code = generate_access_code(8)
                rec["access_code"] = code
                try:
                    access_idx = dist_header.index("access_code")
                except ValueError:
                    st.error("Colonne 'access_code' introuvable.")
                    st.stop()

                target_row_idx = None
                for i, row in enumerate(dist_rows[1:], start=2):
                    if row and row[0] == client_id and len(row) > 3 and row[3] == rec["email_dest"]:
                        target_row_idx = i
                        break
                if access_idx is not None and target_row_idx:
                    dist_sheet.update_cell(target_row_idx, access_idx + 1, code)

                video_url = client_data.get("lien_vidéo", "")
                pdf_path = create_recipient_pdf({
                    "prénom": rec["prénom_dest"],
                    "nom": rec["nom_dest"],
                    "email": rec["email_dest"],
                    "phone": rec["phone_dest"],
                    "addr": rec["addr_dest"],
                    "message": rec["message"]
                }, client_data, video_url, code, condition_text)

                created_files.append(pdf_path)
                attachments.append({"path": pdf_path, "name": f"{rec['prénom_dest']}_{rec['nom_dest']}_fiche.pdf"})

            client_summary_path = create_client_pdf(client_data, client_dests, condition_text)
            created_files.append(client_summary_path)
            attachments.append({"path": client_summary_path, "name": f"{client_data.get('prénom','')}_{client_data.get('nom','')}_summary.pdf"})

            client_email = client_data.get("email")
            subject = "Votre Capsule Temporelle - Documents et confirmation"
            body = (
                f"Bonjour {client_data.get('prénom','')},\n\n"
                "Merci. Votre Capsule Temporelle a été finalisée (DEMO).\n"
                "Vous trouverez en pièces jointes :\n"
                "- Le récapitulatif client\n"
                "- Les fiches destinataires (une par destinataire)\n\n"
                "Les codes d'accès ont été générés et stockés.\n\n"
                "Cordialement,\nL'équipe Capsule Temporelle"
            )
            send_email_with_attachments_gmail(subject, body, client_email, attachments, creds_file="token_gmail.pkl")

            for fpath in created_files:
                try: os.remove(fpath)
                except: pass

            st.balloons()
            st.success("🎉 Félicitations — Opération 100% réussie")
            st.markdown("""
            <div style="padding:12px;background:#F6FFEE;color:black;border-left:4px solid #2ECC71;border-radius:6px">
            <strong>Votre capsule est prête.</strong><br>
            Les fiches destinataires ont été générées et envoyées.<br><br>
            Vous Pouvez Fermer Cette Page ! <br>
            Merci ! <br>
            L'Equipe Capsule Temporelle.
            </div>
            """, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Erreur lors du traitement : {e}")
            import traceback
            st.text(traceback.format_exc())
        
