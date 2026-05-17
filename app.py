import streamlit as st
import datetime as dt
from portfolio_engine import process_portfolio

st.set_page_config(
    page_title="Analyseur de Portefeuille Yahoo",
    page_icon="📈",
    layout="centered"
)

st.title("📈 Analyseur de Portefeuille Yahoo")
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
            log_box = st.empty()
            status_box = st.empty()

            status_box.info("Analyse en cours...")

            output_file = process_portfolio(uploaded_file)

            try:
                with open("/tmp/portfolio_log.txt", "r", encoding="utf-8") as f:
                    logs = f.read()
                    log_box.text_area("Journal d’analyse", logs, height=300)
            except:
                log_box.warning("Aucun journal disponible.")

            status_box.success("Analyse terminée avec succès.")

            st.download_button(
                label="Télécharger le fichier Excel mis à jour (Streamlit → votre appareil)",
                data=output_file,
                file_name=f"Portefeuille_MAJ_{dt.datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erreur : {e}")

