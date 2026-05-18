import streamlit as st
import datetime as dt
import time
from portfolio_engine import process_portfolio

st.set_page_config(
    page_title="Analyseur de Portefeuille Yahoo",
    page_icon="📈",
    layout="centered"
)

st.title("📈 Portefeuille Yahoo")
st.write("Téléversez votre fichier Excel de portefeuille, lancez l’analyse, puis téléchargez le fichier mis à jour.")
with open("Fonds-Compagnies-Suivi_MODELE.xlsx", "rb") as template_file:
    st.download_button(
        label="Télécharger le fichier modèle (GitHub → votre appareil)",
        data=template_file,
        file_name="Fonds-Compagnies-Suivi_MODELE.xlsx",
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

            
            progress_bar.progress(100)
            time.sleep(0.5)
            
            status_box.success("Analyse terminée avec succès.")

            st.download_button(
                label="Télécharger le fichier Excel mis à jour (Streamlit → votre appareil)",
                data=output_file,
                file_name=f"Portefeuille_MAJ_{dt.datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erreur : {e}")

