import streamlit as st
import pandas as pd
import sqlite3
import datetime
from fpdf import FPDF
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut

# --- CONFIGURATION ---
st.set_page_config(page_title="Mirvanie Sénégal - Estimateur", page_icon="🇸🇳", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { background-color: #00853F; color: white; border-radius: 5px; } /* Vert Sénégal */
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    </style>
""", unsafe_allow_html=True)

# --- INITIALISATION SQLITE ---
def init_db():
    conn = sqlite3.connect('mirvanie_estimations_sn.db')
    c = conn.cursor()
    c.execute('DROP TABLE IF EXISTS estimations')
    c.execute('''CREATE TABLE estimations
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  date TEXT, adresse TEXT, type_bien TEXT, segment TEXT,
                  surface_habitable REAL, surface_terrain REAL, 
                  nb_pieces INTEGER, nb_chambres INTEGER, nb_sdb INTEGER, etage INTEGER, cour_jardin TEXT,
                  annee INTEGER, etat TEXT, docs TEXT, equipements TEXT, parking TEXT,
                  estimation_totale REAL)''')
    conn.commit()
    conn.close()

init_db()

# --- CHARGEMENT EXCEL ---
@st.cache_data
def load_zones():
    try:
        df = pd.read_excel('mirvanie_grille_senegal.xlsx', sheet_name='Zones')
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        st.error(f"Erreur de chargement de l'Excel : {e}")
        return pd.DataFrame()

df_zones = load_zones()

# --- GÉNÉRATION DU PDF ---
def generate_pdf(data, details, estimation_finale):
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_fill_color(0, 133, 63) # Vert Sénégal
    pdf.rect(0, 0, 210, 35, 'F')
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", 'B', 22)
    pdf.cell(0, 15, "MIRVANIE SENEGAL", ln=True, align='C')
    pdf.set_font("helvetica", 'I', 12)
    pdf.cell(0, 10, "Rapport d'Estimation Immobiliere", ln=True, align='C')
    
    pdf.set_text_color(0, 0, 0)
    pdf.ln(15)
    pdf.set_font("helvetica", 'B', 14)
    pdf.cell(0, 10, "1. Description complete du bien", ln=True)
    pdf.set_font("helvetica", size=10)
    
    col_width = 60
    def add_row(label, value):
        pdf.cell(col_width, 7, label, border=0)
        pdf.cell(0, 7, str(value), border=0, ln=True)

    add_row("Localisation :", data['Localisation'])
    add_row("Type de bien :", data['Type'])
    add_row("Marche :", data['Marche'])
    
    if data['Type'] != "Appartement":
        add_row("Surface habitable :", f"{data['Surface habitable']} m2")
        add_row("Surface du terrain :", f"{data['Surface terrain']} m2")
    else:
        add_row("Surface habitable :", f"{data['Surface habitable']} m2")
    
    if data['Type'] != "Terrain":
        add_row("Nombre de pieces :", data['Nb_pieces'])
        add_row("Nombre de chambres :", data['Nb_chambres'])
        add_row("Salles de bain :", data['Nb_sdb'])
        add_row("Etage :", data['Etage'])
        if data['Type'] == "Maison/villa":
            add_row("Cour / Jardin :", data['Cour_jardin'])
        add_row("Annee de construction :", data['Annee'])
        add_row("Etat du bien :", data['Etat'])

    add_row("Documents fonciers :", data['Documents'])
    add_row("Equipements :", data['Equipements'])
    add_row("Stationnement :", data['Parking'])

    pdf.ln(10)
    pdf.set_font("helvetica", 'B', 14)
    pdf.cell(0, 10, "2. Synthese de l'estimation", ln=True)
    
    pdf.set_fill_color(240, 255, 240) # Vert très clair
    pdf.set_font("helvetica", 'B', 16)
    pdf.cell(0, 15, f"Valeur estimee : {estimation_finale:,.0f} FCFA", border=1, ln=True, align='C', fill=True)
    
    pdf.set_font("helvetica", 'I', 10)
    pdf.cell(0, 10, f"Fourchette de confiance (-10% / +10%) : {estimation_finale * 0.90:,.0f} FCFA  -  {estimation_finale * 1.10:,.0f} FCFA", ln=True, align='C')

    pdf.ln(10)
    pdf.set_font("helvetica", 'B', 14)
    pdf.cell(0, 10, "3. Methodologie de calcul", ln=True)
    pdf.set_font("helvetica", size=10)
    
    if data['Type'] == "Appartement":
        texte_explication = f"L'estimation se base sur la valeur du bati (Indice de la zone : {details['prix_bati_m2']:,.0f} FCFA/m2) et des ajustements hedoniques."
    else:
        texte_explication = f"L'estimation separe la valeur du bati (Indice : {details['prix_bati_m2']:,.0f} FCFA/m2) et la valeur du terrain (Indice : {details['prix_terrain_m2']:,.0f} FCFA/m2)."
    pdf.multi_cell(0, 6, texte_explication)
    
    pdf.ln(5)
    pdf.set_font("helvetica", 'B', 11)
    pdf.cell(0, 8, f"Valeur de base theorique (Sans ajustements) : {details['valeur_base_simple']:,.0f} FCFA", ln=True)
    
    pdf.ln(5)
    pdf.set_font("helvetica", 'B', 12)
    pdf.cell(0, 10, "Detail des ajustements appliques :", ln=True)
    
    pdf.set_fill_color(200, 200, 200)
    pdf.set_font("helvetica", 'B', 9)
    pdf.cell(110, 8, "Critere", border=1, fill=True)
    pdf.cell(30, 8, "Impact %", border=1, fill=True, align='C')
    pdf.cell(45, 8, "Impact Valeur (FCFA)", border=1, fill=True, align='C', ln=True)
    
    pdf.set_font("helvetica", size=9)
    if not details['ajustements']:
        pdf.cell(0, 7, "Aucun ajustement majeur applique.", border=1, ln=True, align='C')
    else:
        for key, val in details['ajustements'].items():
            impact_pct = val['pct']
            impact_val = val['valeur']
            
            pdf.cell(110, 7, str(key), border=1)
            
            if impact_pct > 0:
                pdf.set_text_color(0, 128, 0)
                pdf.cell(30, 7, f"+{impact_pct:.1f}%", border=1, align='C')
            elif impact_pct < 0:
                pdf.set_text_color(255, 0, 0)
                pdf.cell(30, 7, f"{impact_pct:.1f}%", border=1, align='C')
            else:
                pdf.set_text_color(0, 0, 0)
                pdf.cell(30, 7, "0.0%", border=1, align='C')
                
            pdf.set_text_color(0, 0, 0)
            pdf.cell(45, 7, f"{impact_val:+,.0f}", border=1, align='R', ln=True)

    pdf.ln(15)
    pdf.set_font("helvetica", 'I', 8)
    pdf.set_text_color(100, 100, 100)
    pdf.multi_cell(0, 4, "Ce rapport est genere automatiquement par l'algorithme Mirvanie Senegal. Il s'agit d'une estimation indicative qui ne remplace pas une expertise immobiliere officielle (Notaire, Geometre).", align='C')

    return bytes(pdf.output())

# --- LOGIQUE D'ESTIMATION (Spécifique Sénégal) ---
def calculer_estimation(loc, type_bien, segment, surf_hab, surf_terrain, nb_pieces, nb_chambres, nb_sdb, etage, cour_jardin, annee, etat, docs, equipements, parking_choices):
    
    res_bati = df_zones[(df_zones['localisation'] == loc) & 
                        (df_zones['type_bien'] == type_bien) & 
                        (df_zones['segment'] == segment)]
    if res_bati.empty:
        return None, None, "Données non disponibles pour cette combinaison (Type/Segment)."
    prix_bati_m2 = (res_bati.iloc[0]['prix_min_m2'] + res_bati.iloc[0]['prix_max_m2']) / 2

    if type_bien != "Appartement":
        res_terrain = df_zones[(df_zones['localisation'] == loc) & (df_zones['type_bien'] == "Terrain")]
        prix_terrain_m2 = ((res_terrain.iloc[0]['prix_min_m2'] + res_terrain.iloc[0]['prix_max_m2']) / 2) if not res_terrain.empty else 0
    else:
        prix_terrain_m2 = 0

    if type_bien == "Terrain":
        valeur_bati_base = 0
        valeur_terrain_base = surf_terrain * prix_bati_m2 
        valeur_base_simple = valeur_terrain_base
    elif type_bien == "Appartement":
        valeur_bati_base = surf_hab * prix_bati_m2
        valeur_terrain_base = 0
        valeur_base_simple = valeur_bati_base
    else:
        valeur_bati_base = surf_hab * prix_bati_m2
        surf_terrain_libre = max(0, surf_terrain - surf_hab)
        valeur_terrain_base = surf_terrain_libre * prix_terrain_m2
        valeur_base_simple = valeur_bati_base + valeur_terrain_base

    ajustements = {}
    valeur_bati_courante = valeur_bati_base
    valeur_terrain_courante = valeur_terrain_base

    if type_bien != "Terrain":
        age = max(0, datetime.datetime.now().year - annee)
        coef_vetuste = max(0.70, 1 - (age * 0.01))
        if coef_vetuste < 1:
            impact = valeur_bati_courante * (coef_vetuste - 1)
            ajustements["Decote Vetuste (Age)"] = {"pct": (coef_vetuste - 1) * 100, "valeur": impact}
            valeur_bati_courante += impact

        coef_etat_map = {"excellent": 1.10, "bon": 1.00, "moyen": 0.85, "a renover": 0.60}
        coef_etat = coef_etat_map.get(etat.lower(), 1.0)
        if coef_etat != 1.0:
            impact = valeur_bati_courante * (coef_etat - 1)
            ajustements["Ajustement Etat General"] = {"pct": (coef_etat - 1) * 100, "valeur": impact}
            valeur_bati_courante += impact

        if nb_chambres > 3:
            bonus_conf_pct = min((nb_chambres - 3) * 0.02, 0.10)
            impact = valeur_bati_courante * bonus_conf_pct
            ajustements["Configuration (Chambres)"] = {"pct": bonus_conf_pct * 100, "valeur": impact}
            valeur_bati_courante += impact

        if type_bien == "Appartement":
            if etage == 0:
                impact = valeur_bati_courante * 0.02
                ajustements["Rez-de-chaussee"] = {"pct": 2.0, "valeur": impact}
                valeur_bati_courante += impact
            elif etage > 3 and "Ascenseur" not in equipements:
                impact = valeur_bati_courante * -0.05
                ajustements["Etage eleve sans ascenseur"] = {"pct": -5.0, "valeur": impact}
                valeur_bati_courante += impact

        bonus_equip_pct = 0
        if "Piscine" in equipements: bonus_equip_pct += 0.08
        if "Ascenseur" in equipements: bonus_equip_pct += 0.05
        if "Groupe electrogene" in equipements: bonus_equip_pct += 0.04
        if "panneau solaire" in equipements: bonus_equip_pct += 0.03
        if "Electricite H24" in equipements: bonus_equip_pct += 0.04
        if "Securite 24h/24" in equipements: bonus_equip_pct += 0.04
        if "Climatisation" in equipements: bonus_equip_pct += 0.03
        if "Balcon/Terrasse" in equipements: bonus_equip_pct += 0.03
        if "Cuisine equipee" in equipements: bonus_equip_pct += 0.03
        if "Chauffage" in equipements: bonus_equip_pct += 0.03
        if "eau chaude" in equipements: bonus_equip_pct += 0.02
        
        bonus_equip_pct = min(bonus_equip_pct, 0.30)
        if bonus_equip_pct > 0:
            impact = valeur_bati_courante * bonus_equip_pct
            ajustements["Equipements de confort"] = {"pct": bonus_equip_pct * 100, "valeur": impact}
            valeur_bati_courante += impact

        bonus_parking_pct = 0
        if "Prive" in parking_choices: bonus_parking_pct += 0.03
        if "Garage" in parking_choices: bonus_parking_pct += 0.05
        bonus_parking_pct = min(bonus_parking_pct, 0.09)
        if bonus_parking_pct > 0:
            impact = valeur_bati_courante * bonus_parking_pct
            ajustements["Stationnement"] = {"pct": bonus_parking_pct * 100, "valeur": impact}
            valeur_bati_courante += impact

    # --- AJUSTEMENTS TERRAIN (Spécifique Sénégal) ---
    if valeur_terrain_courante > 0:
        # Coefficient de sécurité juridique Sénégal
        coef_doc_map = {
            "Titre Foncier": 1.20,             # Le Graal
            "Bail Emphytéotique": 1.10,        # Très sécurisé
            "Acte Administratif": 0.95,        # Intermédiaire
            "Délibération Municipale": 0.80,   # Précaire
            "Permis d'Occuper": 0.75,          # Très précaire
            "Non communiqué": 0.70
        }
        coef_doc = max([coef_doc_map.get(d, 0.70) for d in docs]) if docs else 0.70
        
        if coef_doc != 1.0:
            impact = valeur_terrain_courante * (coef_doc - 1)
            ajustements["Securisation Fonciere (Terrain)"] = {"pct": (coef_doc - 1) * 100, "valeur": impact}
            valeur_terrain_courante += impact

        if type_bien == "Maison/villa" and cour_jardin == "Oui":
            impact = valeur_terrain_courante * 0.05
            ajustements["Amenagement Cour / Jardin"] = {"pct": 5.0, "valeur": impact}
            valeur_terrain_courante += impact

    estimation_finale = valeur_bati_courante + valeur_terrain_courante
    details = {
        "prix_bati_m2": prix_bati_m2,
        "prix_terrain_m2": prix_terrain_m2,
        "valeur_base_simple": valeur_base_simple,
        "ajustements": ajustements
    }
    return estimation_finale, details, None

# --- INTERFACE ---
tab1, tab2 = st.tabs(["🏠 Nouvelle Estimation", "🔐 Admin Mirvanie"])

with tab1:
    st.title("🏠 Estimateur Immobilier Mirvanie Sénégal")
    
    if not df_zones.empty:
        with st.form("estimation_form"):
            st.subheader("📍 Localisation (Géocodage)")
            
            # --- AMÉLIORATION : GÉOCODAGE NOMINATIM ---
            adresse_recherche = st.text_input("Recherchez votre adresse (ex: Almadies, Dakar) :", "")
            if adresse_recherche:
                with st.spinner("Recherche de l'adresse..."):
                    geolocator = Nominatim(user_agent="mirvanie_senegal")
                    try:
                        location = geolocator.geocode(f"{adresse_recherche}, Senegal")
                        if location: 
                            st.success(f"📍 Adresse trouvée : {location.address}")
                        else: 
                            st.warning("Adresse introuvable. Utilisez la liste ci-dessous.")
                    except GeocoderTimedOut: 
                        st.warning("Service de géocodage indisponible.")

            # Liste de secours pour lier l'adresse à la micro-zone
            communes = sorted(df_zones['localisation'].unique())
            loc_choisie = st.selectbox("Ou sélectionnez la micro-zone correspondante :", communes)
            
            types_dispo = sorted(df_zones[df_zones['localisation'] == loc_choisie]['type_bien'].unique())
            type_choisi = st.selectbox("Type de bien :", types_dispo)
            
            segments_dispo = sorted(df_zones[(df_zones['localisation'] == loc_choisie) & 
                                             (df_zones['type_bien'] == type_choisi)]['segment'].unique())
            segment_choisi = st.selectbox("Marché (Neuf ou Ancien) :", segments_dispo)

            st.subheader("📐 Configuration & Caractéristiques")
            col1, col2 = st.columns(2)
            
            with col1:
                surf_hab_defaut = 100 if type_choisi != "Terrain" else 0
                surface_habitable = st.number_input("Surface habitable (m²) :", min_value=0, value=surf_hab_defaut)
                
                if type_choisi == "Appartement":
                    surface_terrain = 0
                    st.info("Un appartement ne dispose pas de terrain privatif.")
                else:
                    surface_terrain = st.number_input("Surface totale du terrain (m²) :", min_value=10, value=100)
                
                if type_choisi != "Terrain":
                    nb_pieces = st.number_input("Nombre de pièces :", min_value=1, value=3)
                    nb_chambres = st.number_input("Nombre de chambres :", min_value=0, value=2)
                    nb_sdb = st.number_input("Nombre de salles de bain :", min_value=1, value=1)
                    etage = st.number_input("Étage (0 pour RDC) :", min_value=0, value=0) if type_choisi == "Appartement" else 0
                    
                    if type_choisi == "Maison/villa":
                        cour_jardin = st.radio("Cour / Jardin :", ["Non", "Oui"])
                    else:
                        cour_jardin = "Non"
                        
                    annee_construction = st.number_input("Année de construction :", min_value=1900, max_value=2026, value=2015)
                    etat_bien = st.selectbox("État du bien :", ["excellent", "bon", "moyen", "à rénover"])
                else:
                    nb_pieces, nb_chambres, nb_sdb, etage = 0, 0, 0, 0
                    cour_jardin = "Oui"
                    annee_construction = 2026
                    etat_bien = "bon"

            with col2:
                # Documents fonciers Sénégalais
                docs_foncier = st.multiselect("Document(s) foncier(s) :", [
                    "Titre Foncier", "Bail Emphytéotique", "Acte Administratif", 
                    "Délibération Municipale", "Permis d'Occuper", "Non communiqué"
                ], default=["Titre Foncier"])
                
                parking_choices = st.multiselect("Stationnement :", ["Aucun", "Rue", "Privé", "Garage"], default=["Aucun"])
                
                equipements = st.multiselect("Équipements :", [
                    "Balcon/Terrasse", "Ascenseur", "Piscine", "Chauffage", 
                    "eau chaude", "panneau solaire", "Groupe électrogène",
                    "Climatisation", "Électricité H24", "Sécurité 24h/24", "Cuisine équipée"
                ])

            submit = st.form_submit_button("💰 Lancer l'estimation", use_container_width=True)

        if submit:
            estimation, details, erreur = calculer_estimation(
                loc_choisie, type_choisi, segment_choisi, 
                surface_habitable, surface_terrain, nb_pieces, nb_chambres, nb_sdb, etage, cour_jardin,
                annee_construction, etat_bien, docs_foncier, equipements, parking_choices
            )
            
            if erreur:
                st.warning(erreur)
            else:
                st.success("### 📊 Estimation Mirvanie Sénégal")
                col_res1, col_res2 = st.columns([1, 1])
                with col_res1:
                    st.metric(label="Valeur estimée", value=f"{estimation:,.0f} FCFA")
                    st.info(f"📉 **Basse (-10%)** : {estimation * 0.90:,.0f} FCFA\n\n📈 **Haute (+10%)** : {estimation * 1.10:,.0f} FCFA")
                with col_res2:
                    st.write("**Méthodologie :**")
                    st.write(f"Valeur de base : {details['valeur_base_simple']:,.0f} FCFA")
                    for k, v in details['ajustements'].items():
                        st.write(f"- {k} : **{v['pct']:+.1f}%**")
                
                conn = sqlite3.connect('mirvanie_estimations_sn.db')
                c = conn.cursor()
                c.execute('''INSERT INTO estimations 
                             (date, adresse, type_bien, segment, surface_habitable, surface_terrain, nb_pieces, nb_chambres, nb_sdb, etage, cour_jardin, annee, etat, docs, equipements, parking, estimation_totale)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                          (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), loc_choisie, type_choisi, segment_choisi,
                           surface_habitable, surface_terrain, nb_pieces, nb_chambres, nb_sdb, etage, cour_jardin,
                           annee_construction, etat_bien, ", ".join(docs_foncier), ", ".join(equipements), ", ".join(parking_choices),
                           estimation))
                conn.commit()
                conn.close()

                pdf_data = {
                    "Localisation": loc_choisie, "Type": type_choisi, "Marche": segment_choisi,
                    "Surface habitable": surface_habitable, "Surface terrain": surface_terrain,
                    "Nb_pieces": nb_pieces, "Nb_chambres": nb_chambres, "Nb_sdb": nb_sdb,
                    "Etage": etage, "Cour_jardin": cour_jardin, "Annee": annee_construction,
                    "Etat": etat_bien, "Documents": ", ".join(docs_foncier),
                    "Equipements": ", ".join(equipements) if equipements else "Aucun",
                    "Parking": ", ".join(parking_choices) if parking_choices else "Aucun"
                }
                
                pdf_bytes = generate_pdf(pdf_data, details, estimation)
                
                st.download_button(
                    label="📄 Télécharger le rapport détaillé (PDF)", 
                    data=pdf_bytes, 
                    file_name="rapport_mirvanie_senegal.pdf", 
                    mime="application/pdf"
                )

# --- ONGLET ADMIN ---
with tab2:
    st.title("🔐 Espace Administrateur Sénégal")
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        password = st.text_input("Mot de passe :", type="password")
        if st.button("Connexion"):
            if password == st.secrets.get("ADMIN_PASSWORD", "admin123"):
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
    else:
        st.success("Connecté.")
        if st.button("Déconnexion"):
            st.session_state.authenticated = False
            st.rerun()
        
        conn = sqlite3.connect('mirvanie_estimations_sn.db')
        df_estimations = pd.read_sql_query("SELECT * FROM estimations ORDER BY id DESC", conn)
        conn.close()
        
        if not df_estimations.empty:
            st.subheader("📊 Indicateurs Clés")
            kpi1, kpi2, kpi3 = st.columns(3)
            
            total_est = len(df_estimations)
            prix_moyen = df_estimations['estimation_totale'].mean()
            zone_top = df_estimations['adresse'].mode()[0] if not df_estimations['adresse'].empty else "N/A"
            
            kpi1.metric("Total Estimations", total_est)
            kpi2.metric("Valeur Moyenne", f"{prix_moyen:,.0f} FCFA")
            kpi3.metric("Zone Top", zone_top)
            
            st.divider()
            st.subheader("📋 Historique des estimations")
            csv = df_estimations.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Exporter en CSV", data=csv, file_name="estimations_senegal.csv", mime="text/csv")
            st.dataframe(df_estimations[['date', 'adresse', 'type_bien', 'segment', 'estimation_totale']], use_container_width=True)
        else:
            st.info("Aucune estimation enregistrée.")