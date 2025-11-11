import streamlit as st
from datetime import datetime
import random
import gspread
from google.oauth2.service_account import Credentials

# ---------- CONFIG ----------
INFO_SHEET_ID = "1MFtDX7ZduN9W1VVrlItafDl09OGDscKYgJDe1dnMay0"  # timeCapsule_info_sheet
CREDENTIALS_FILE = "credentials.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ---------- AUTHENTICATION ----------
creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)

# Connect to info sheet
sheet = gc.open_by_key(INFO_SHEET_ID).sheet1

# ---------- STREAMLIT FORM ----------
st.title("🌿 Time Capsule - Étape 1 : Informations Client")

with st.form("client_info_form"):
    prénom = st.text_input("Prénom")
    nom = st.text_input("Nom")
    
    # ---------- DATE INPUT STARTING AT 1900 ----------
    date_naissance = st.date_input(
        "Date de naissance",
        value=datetime(1980, 1, 1),      # default date
        min_value=datetime(1900, 1, 1),  # earliest selectable date
        max_value=datetime.today()       # latest selectable date
    )
    
    email = st.text_input("Email")
    téléphone = st.text_input("Téléphone")
    adresse = st.text_input("Adresse")
    
    submitted = st.form_submit_button("Enregistrer")

if submitted:
    if not (prénom and nom and date_naissance):
        st.error("Veuillez remplir au minimum : prénom, nom et date de naissance.")
    else:
        # Fetch all rows to check for existing client
        rows = sheet.get_all_values()  # 2D list
        header = rows[0]
        data_rows = rows[1:]

        # Find existing client (prénom + nom + date_naissance)
        client_id = None
        for idx, row in enumerate(data_rows, start=2):  # start=2 because sheet rows start at 1 and first row is header
            row_prénom = row[header.index("prénom")].strip().lower()
            row_nom = row[header.index("nom")].strip().lower()
            row_date = row[header.index("date_naissance")].strip()
            
            if (row_prénom == prénom.strip().lower() and
                row_nom == nom.strip().lower() and
                row_date == date_naissance.strftime("%d/%m/%Y")):
                client_id = row[header.index("client_ID")]
                st.info(f"Client existant trouvé. Client ID : {client_id}")
                break

        if client_id is None:
            # Generate new 8-digit client_ID
            client_id = str(random.randint(10000000, 99999999))
            now_str = datetime.now().strftime("%d/%m/%Y %H:%M")

            new_row = [
                client_id,
                prénom,
                nom,
                date_naissance.strftime("%d/%m/%Y"),
                email,
                téléphone,
                adresse,
                "",  # lien_vidéo
                "",  # dossier_vidéo
                "",  # condition
                "",  # lien_liste_distribution
                "en_attente",  # statut
                now_str,  # date_création
                now_str   # date_mise_à_jour
            ]
            sheet.append_row(new_row)
            st.success(f"Nouveau client enregistré. Client ID : {client_id}")
