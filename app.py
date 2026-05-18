import streamlit as st
import datetime as dt
import time
from portfolio_engine import process_portfolio
from zoneinfo import ZoneInfo

st.set_page_config(
    page_title="Analyseur de Portefeuille Yahoo",
    page_icon="📈",
    layout="centered"
)

st.title("📈 Portefeuille Yahoo")
st.write("Téléversez votre fichier Excel de portefeuille, lancez l’analyse, puis téléchargez le fichier mis à jour.")


# Modèle USA / NYSE
with open("Fonds-Compagnies-Suivi_MODELE.xlsx", "rb") as f1:
    st.download_button(
        label="Télécharger modèle USA / NYSE (GitHub → votre appareil)",
        data=f1.read(),
        file_name="Fonds-Compagnies-Suivi_MODELE.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Modèle Canada / CDR / TO
with open("Fonds-Compagnies-Suivi_MODELE_CDR_TO.xlsx", "rb") as f2:
    st.download_button(
        label="Télécharger modèle Canada / CDR / TO (GitHub → votre appareil)",
        data=f2.read(),
        file_name="Fonds-Compagnies-Suivi_MODELE_CDR_TO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

uploaded_file = st.file_uploader(
    "Choisir votre fichier Excel (votre appareil → Streamlit)",
    type=["xlsx"]
)

if uploaded_file is not None:

    st.success("Fichier Excel téléversé avec succès.")

    if st.button("Lancer l’analyse"):
        try:
            status_box = st.empty()

            progress_bar = st.progress(0)

            etapes = [
                ("Synchronisation Yahoo Finance...", 15),
                ("Calcul des métriques...", 40),
                ("Application des règles...", 70),
                ("Génération du fichier Excel...", 90),
            ]

            for message, pct in etapes:
                status_box.info(message)
                progress_bar.progress(pct)
                time.sleep(1.0)

            output_file = process_portfolio(uploaded_file)

            try:
                with open("/tmp/portfolio_log.txt", "r", encoding="utf-8") as f:
                    logs = f.read()

                st.text_area(
                    "Journal d’analyse",
                    logs,
                    height=300
                )
            except:
                st.warning("Aucun journal disponible.")
    
            
            progress_bar.progress(100)
            time.sleep(0.5)
            progress_bar.empty()

            status_box.success("Analyse terminée avec succès.")

            eastern_now = dt.datetime.now(ZoneInfo("America/New_York"))

            uploaded_name = uploaded_file.name.upper()

            if "CDR_TO" in uploaded_name:
                 prefix = "Portefeuille_MAJ_CDR_TO"
            else:
                 prefix = "Portefeuille_MAJ_USA"

            file_name = f"{prefix}_{eastern_now.strftime('%Y-%m-%d_%H%M')}.xlsx"

            st.download_button(
                label="Télécharger le fichier Excel mis à jour (Streamlit → votre appareil)",
                data=output_file,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
          

        except Exception as e:
            st.error(f"Erreur : {e}")

