import io
import datetime as dt

import pandas as pd
import streamlit as st


# ---------- Fonctions utilitaires ----------

def lire_fec(uploaded_file) -> pd.DataFrame:
    """Lit un FEC en devinant le séparateur le plus probable."""
    filename = uploaded_file.name.lower()

    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        df = pd.read_excel(uploaded_file, dtype=str)
    else:
        # Essais successifs de séparateurs classiques
        content = uploaded_file.read()
        for sep in [";", "\t", ",", "|"]:
            try:
                df = pd.read_csv(
                    io.BytesIO(content),
                    sep=sep,
                    dtype=str,
                    engine="python",
                )
                # Si on a au moins 5 colonnes, on considère que c'est bon
                if df.shape[1] >= 5:
                    break
            except Exception:
                df = None
        if df is None:
            st.error("Impossible de lire le fichier. Merci de vérifier le séparateur.")
            st.stop()

    # Normalisation des noms de colonnes
    df.columns = [c.strip() for c in df.columns]

    # Conversion des montants
    for col in ["Debit", "Credit"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .str.replace(" ", "", regex=False)
                .str.replace(",", ".", regex=False)
                .astype(float)
            )
        else:
            st.error(f"Colonne manquante dans le FEC : {col}")
            st.stop()

    # Conversion des dates
    if "PieceDate" in df.columns:
        df["PieceDate"] = pd.to_datetime(df["PieceDate"], errors="coerce")
    else:
        st.error("Colonne 'PieceDate' manquante dans le FEC.")
        st.stop()

    return df


def calc_creances_ouvertes(df: pd.DataFrame, date_anciennete: dt.date) -> pd.DataFrame:
    """
    Calcule les créances clientes ouvertes par facture (PieceRef)
    à une date d'ancienneté donnée.

    Logique :
    - on filtre les comptes 411* (paramétrable si besoin)
    - on regroupe par client + pièce
    - Solde = somme(Débit - Crédit)
    - Montant facture = somme des débits
    - Règlement partiel = Montant facture - Solde
    - On garde uniquement les soldes non nuls et les pièces antérieures à date_anciennete
    """
    df = df.copy()

    # Filtre comptes clients (tu peux élargir à 410 / 418 si besoin)
    df["CompteNum"] = df["CompteNum"].astype(str)
    mask_clients = df["CompteNum"].str.startswith("411")
    df_clients = df[mask_clients].copy()

    if df_clients.empty:
        st.warning("Aucune écriture de compte 411* trouvée dans le FEC.")
        return pd.DataFrame()

    # Colonne Solde par ligne
    df_clients["Solde_ligne"] = df_clients["Debit"] - df_clients["Credit"]

    # Si pas de compte auxiliaire, on remplace par compte général
    if "CompAuxNum" not in df_clients.columns:
        df_clients["CompAuxNum"] = df_clients["CompteNum"]
    if "CompAuxLib" not in df_clients.columns:
        df_clients["CompAuxLib"] = df_clients["CompteLib"]

    group_cols = [
        "CompAuxNum",
        "CompAuxLib",
        "PieceRef",
        "PieceDate",
    ]

    if "PieceRef" not in df_clients.columns:
        st.error("La colonne 'PieceRef' est manquante dans le FEC. On en a besoin pour identifier les factures.")
        st.stop()

    grp = df_clients.groupby(group_cols, dropna=False)

    synthese = grp.agg(
        Montant_facture=("Debit", "sum"),
        Total_credit=("Credit", "sum"),
        Solde=("Solde_ligne", "sum"),
    ).reset_index()

    # Nettoyage
    synthese["Solde"] = synthese["Solde"].round(2)
    synthese["Montant_facture"] = synthese["Montant_facture"].round(2)
    synthese["Total_credit"] = synthese["Total_credit"].round(2)

    # Règlement partiel = montant payé sur cette facture
    synthese["Reglement_partiel"] = (synthese["Montant_facture"] - synthese["Solde"]).clip(lower=0).round(2)

    # Filtre : pièces antérieures à la date d'ancienneté
    synthese = synthese[synthese["PieceDate"].dt.date <= date_anciennete]

    # On garde seulement les factures encore ouvertes (solde != 0)
    synthese = synthese[synthese["Solde"].abs() > 0.01]

    # Tri par client puis date
    synthese = synthese.sort_values(["CompAuxNum", "PieceDate", "PieceRef"])

    return synthese


def fabriquer_tableau_client(df_ouvert: pd.DataFrame, client_code: str) -> pd.DataFrame:
    """Construit le tableau à envoyer par mail pour un client donné."""
    df_client = df_ouvert[df_ouvert["CompAuxNum"] == client_code].copy()
    if df_client.empty:
        return df_client

    df_client["PieceDate"] = df_client["PieceDate"].dt.strftime("%d/%m/%Y")

    # Colonnes à envoyer au client + colonnes à renseigner
    df_client = df_client[[
        "PieceDate",
        "PieceRef",
        "Montant_facture",
        "Reglement_partiel",
        "Solde",
    ]]

    df_client = df_client.rename(columns={
        "PieceDate": "Date facture",
        "PieceRef": "Référence facture",
        "Montant_facture": "Montant facture TTC",
        "Reglement_partiel": "Règlement(s) déjà reçu(s)",
        "Solde": "Solde restant dû",
    })

    # Colonnes que le client devra compléter
    df_client["Créance douteuse ? (Oui/Non)"] = ""
    df_client["Si douteuse, montant ou % douteux"] = ""
    df_client["Manque-t-il un avoir ? (Oui/Non)"] = ""
    df_client["Si payé, date de règlement"] = ""
    df_client["Commentaires (client)"] = ""

    return df_client


def proposer_mail(client_name: str,
                  client_code: str,
                  date_situation: dt.date,
                  date_anciennete: dt.date) -> str:
    """Génère une proposition de mail à adapter par le collaborateur."""
    objet = f"Point sur vos factures en attente au {date_situation.strftime('%d/%m/%Y')}"
    corps = f"""Objet : {objet}

Bonjour {client_name},

Dans le cadre de l'arrêté de vos comptes, nous réalisons un point sur les factures clients en attente de règlement.

Vous trouverez en pièce jointe un tableau récapitulatif des créances encore ouvertes sur votre compte (code client {client_code}) pour des factures antérieures au {date_anciennete.strftime('%d/%m/%Y')}.

Pour chaque ligne, nous vous remercions de bien vouloir :
- confirmer si la créance est ou non douteuse,
- préciser, le cas échéant, le montant ou le pourcentage que vous considérez comme douteux,
- nous indiquer s'il manque un avoir,
- nous préciser la date de règlement lorsque la facture a déjà été soldée,
- compléter, si nécessaire, la colonne “Commentaires”.

Ces informations nous permettront :
- d’actualiser la situation de vos comptes clients,
- et, le cas échéant, d’évaluer les provisions pour créances douteuses à comptabiliser.

Nous vous remercions par avance pour votre retour, idéalement sous 8 jours, en nous renvoyant le fichier complété.

Restant à votre disposition pour toute précision,

Cordialement,

[Nom du collaborateur]
[Cabinet]
[Téléphone]
[Email]
"""
    return corps


# ---------- Interface Streamlit ----------

st.set_page_config(page_title="Relances clients à partir du FEC", layout="wide")

st.title("📂 Relances clients à partir du FEC")
st.write(
    "Cette application lit un FEC, identifie les **créances clients encore ouvertes** "
    "et prépare un **mail de relance** et un **tableau à envoyer au client**."
)

uploaded_file = st.file_uploader(
    "Importer le FEC (format CSV / TXT / Excel)",
    type=["csv", "txt", "xlsx", "xls"],
)

if uploaded_file is not None:
    df_fec = lire_fec(uploaded_file)

    st.success("FEC importé avec succès ✅")

    col1, col2 = st.columns(2)
    with col1:
        date_situation = st.date_input(
            "Date de situation (date des comptes / relance)",
            value=dt.date.today(),
        )
    with col2:
        date_anciennete = st.date_input(
            "Prendre les créances antérieures au",
            value=dt.date.today(),
            help="Seules les factures avec une date pièce antérieure ou égale à cette date seront retenues.",
        )

    # Calcul des créances ouvertes
    df_ouvert = calc_creances_ouvertes(df_fec, date_anciennete=date_anciennete)

    if df_ouvert.empty:
        st.info("Aucune créance ouverte trouvée selon les critères définis.")
        st.stop()

    st.subheader("Synthèse des créances clients ouvertes")
    st.write(
        "Il s'agit des factures en comptes 411* dont le solde reste non nul "
        f"pour des factures antérieures au {date_anciennete.strftime('%d/%m/%Y')}."
    )
    st.dataframe(df_ouvert.head(100), use_container_width=True)

    # Choix du client
    clients = (
        df_ouvert[["CompAuxNum", "CompAuxLib"]]
        .drop_duplicates()
        .sort_values("CompAuxNum")
    )

    st.subheader("Préparation du mail par client")
    client_labels = {
        f"{row.CompAuxNum} - {row.CompAuxLib}": row.CompAuxNum
        for row in clients.itertuples()
    }

    choix_label = st.selectbox(
        "Sélectionner un client",
        options=list(client_labels.keys()),
    )

    if choix_label:
        client_code = client_labels[choix_label]
        client_name = clients.loc[clients["CompAuxNum"] == client_code, "CompAuxLib"].iloc[0]

        # Tableau spécifique à ce client
        df_client_mail = fabriquer_tableau_client(df_ouvert, client_code)

        if df_client_mail.empty:
            st.warning("Aucune créance ouverte pour ce client.")
        else:
            st.markdown(f"### Tableau des créances pour : **{client_name}** ({client_code})")
            st.dataframe(df_client_mail, use_container_width=True)

            # Export Excel pour pièce jointe
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_client_mail.to_excel(writer, index=False, sheet_name="Relance client")
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger le tableau Excel à joindre au mail",
                data=buffer,
                file_name=f"relance_client_{client_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Proposition de mail
            st.markdown("### Proposition de mail (modifiable)")
            mail_suggestion = proposer_mail(
                client_name=client_name,
                client_code=client_code,
                date_situation=date_situation,
                date_anciennete=date_anciennete,
            )

            texte_mail = st.text_area(
                "Texte du mail à copier/coller dans votre messagerie :",
                value=mail_suggestion,
                height=400,
            )

else:
    st.info("➡️ Commence par importer un FEC pour continuer.")
