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

uploaded_file = st.file_uploader(
    "Choisir votre fichier Excel",
    type=["xlsx"]
)

if uploaded_file is not None:

    st.success("Fichier Excel téléversé avec succès.")

    if st.button("Lancer l’analyse"):

        with st.spinner("Traitement des données Yahoo Finance... veuillez patienter..."):

            try:
                output_file = process_portfolio(uploaded_file)

                st.success("Analyse terminée avec succès.")

                st.download_button(
                    label="Télécharger le fichier Excel mis à jour",
                    data=output_file,
                    file_name=f"Portefeuille_MAJ_{dt.datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Erreur : {e}")
