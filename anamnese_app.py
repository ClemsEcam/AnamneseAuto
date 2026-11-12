from datetime import date
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
import streamlit as st

st.set_page_config(page_title="Anamnèse Psychomotricité", layout="wide")

st.title("📋 Assistant de Rédaction - Anamnèse")
st.markdown(
    "Remplissez les champs pertinents lors de l'entretien. Les champs vides"
    " seront automatiquement ignorés lors de la rédaction."
)

# --- FORMULAIRE PAR ONGLETS ---
tabs = st.tabs([
    "1. Identité & Famille",
    "2. Motif",
    "3. Antécédents & Santé",
    "4. Développement",
    "5. Vie Quotidienne",
    "6. École & Garde",
    "7. Comportement",
    "8. Suivis & Attentes",
])

with tabs[0]:
    col1, col2, col3 = st.columns(3)
    with col1:
        nom = st.text_input("Nom de l'enfant")
        prenom = st.text_input("Prénom de l'enfant")
        sexe = st.selectbox("Sexe", ["Garçon", "Fille"])
    with col2:
        ddn = st.date_input("Date de naissance", value=date(2017, 1, 1))
        age_texte = st.text_input(
            "Âge (ex: 8 ans 6 mois)", placeholder="8 ans"
        )
        scolarite = st.text_input("Classe / Scolarité", placeholder="CE2")
    with col3:
        ville = st.text_input("Ville de résidence", placeholder="Chatou")
        prescripteur = st.text_input(
            "Médecin prescripteur", placeholder="Dr P. Martin (Chatou)"
        )
        generaliste = st.text_input(
            "Médecin généraliste", placeholder="Dr D. Dupont (Le Vesinet)"
        )

    st.subheader("Composition familiale")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        schema_familial = st.selectbox(
            "Schéma familial",
            [
                "Vit avec ses deux parents",
                "Parents séparés (garde alternée)",
                "Vit principalement avec sa mère",
                "Vit principalement avec son père",
                "Autre",
            ],
        )
        fratrie = st.text_input(
            "Fratrie (ex: fils unique / une sœur de 5 ans)",
            value="enfant unique",
        )
        animaux = st.text_input(
            "Animaux de compagnie (optionnel)",
            placeholder="une petite gerbille prénommée Zoé",
        )
    with col_f2:
        presents = st.text_input(
            "Personne(s) présente(s) à l'entretien", value="sa maman"
        )
        langues = st.text_input(
            "Langues parlées à la maison", value="le français"
        )
        chambre = st.radio(
            "Chambre", ["Dispose de sa propre chambre", "Partage sa chambre"]
        )

with tabs[1]:
    indication = st.text_area(
        "Indication / Motif principal de consultation",
        placeholder="Difficultés remontées par l’école, difficultés d’inhibition verbale",
    )
    orientateur = st.text_input(
        "Qui a orienté vers la psychomotricité ?",
        placeholder="L'enseignante / Le médecin",
    )
    detail_motif = st.text_area(
        "Contexte de la demande et observations initiales",
        placeholder="L’enseignante décrit une classe bavarde...",
    )

with tabs[2]:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        sante_globale = st.selectbox(
            "État de santé général",
            [
                "Va bien, pas de problème particulier",
                "Antécédents médicaux / suivi particulier",
            ],
        )
        antecedents_details = st.text_area(
            "Détails médicaux (hospitalisations, opérations, traitements, antécédents familiaux)"
        )
    with col_s2:
        controles_sensoriels = st.checkbox(
            "Contrôles visuels et auditifs à jour", value=True
        )
        vision_audition = st.text_input(
            "Précisions vue / audition",
            placeholder="Porte des lunettes pour une myopie",
        )
        sensorialite_tactile = st.text_input(
            "Particularités sensorielles (toucher, bruit, etc.)"
        )

with tabs[3]:
    st.subheader("Grossesse et naissance")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        grossesse = st.selectbox(
            "Grossesse et accouchement",
            [
                "Au terme d'une grossesse et d'un accouchement dans la norme",
                "Grossesse ou accouchement avec particularités",
            ],
        )
        cesarienne = st.checkbox("Naissance par césarienne")
        naissance_details = st.text_input(
            "Précisions naissance / premiers mois",
            placeholder="Décrit comme un bébé très vif dès la naissance",
        )
    with col_g2:
        allaitement = st.selectbox(
            "Alimentation nourrisson",
            [
                "Allaitement sans difficulté",
                "Allaitement difficile (difficultés de succion), passage au biberon",
                "Biberon dès la naissance",
            ],
        )

    st.subheader("Acquisitions motrices et langagières")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        etapes_motrices = st.selectbox(
            "Développement moteur global",
            [
                "Étapes acquises dans les temps",
                "Quelques décalages dans les acquisitions",
            ],
        )
        details_moteur = st.text_area(
            "Âges clés et aisance motrice (4 pattes, marche, vélo, nage, coordination)",
            placeholder="A marché vers 12 mois après du quatre pattes vers 6 mois. Fait du vélo sans petites roues. À l'aise dans l'eau.",
        )
    with col_m2:
        langage = st.selectbox(
            "Acquisition du langage",
            [
                "Dans les temps, sans difficulté",
                "Un peu tardive",
                "Difficultés d'articulation / suivi orthophonique",
            ],
        )
        details_langage = st.text_area(
            "Précisions langage / communication",
            placeholder="Ce qui peut s'expliquer par le bilinguisme familial.",
        )
        lateralite = st.selectbox(
            "Latéralité observée", ["Droitière", "Gauchère", "Non fixée / Ambiguë"]
        )

with tabs[4]:
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        alimentation = st.text_area(
            "Alimentation actuelle et passée",
            placeholder="Son alimentation est aujourd'hui équilibrée, même s'il a présenté plus jeune des difficultés avec certaines textures.",
        )
        sommeil = st.text_area(
            "Sommeil (endormissement, réveils, cauchemars)",
            placeholder="Le sommeil est globalement bon. Il lui arrive de rêver de sa journée, réveils matinaux sans difficulté.",
        )
    with col_v2:
        autonomie = st.text_area(
            "Autonomie quotidienne (habillage, lacets, cartable, tâches)",
            placeholder="Sait s'habiller seul et préparer son sac. Commence à apprendre à faire ses lacets.",
        )
        ecrans = st.text_input(
            "Rapport aux écrans",
            placeholder="Les écrans sont encadrés au domicile (jeux vidéo occasionnels).",
        )
        loisirs = st.text_area(
            "Jeux à la maison et activités extrascolaires",
            placeholder="Apprécie les Lego, le piano. Pratique le Krav Maga et la Capoeira.",
        )

with tabs[5]:
    garde = st.text_input(
        "Mode de garde avant l'école",
        placeholder="Accueilli par une nourrice de 3 mois à 2 ans, puis crèche de 2 à 3 ans.",
    )
    vecu_ecole = st.text_area(
        "Rapport à l'école et apprentissages",
        placeholder="Dans l'ensemble, il apprécie l'école, bien qu'il puisse parfois être contrarié à l'idée d'y aller.",
    )
    sociabilite_ecole = st.text_area(
        "Relations avec les pairs et l'enseignant",
        placeholder="Les relations sociales à l'école peuvent être délicates : peu d'amis en classe, contrairement au sport.",
    )

with tabs[6]:
    caractere = st.text_area(
        "Caractère et gestion émotionnelle (selon les parents)",
        placeholder="Enfant expressif, parfois dans sa bulle. Intériorise beaucoup ses émotions et cherche à éviter les conflits.",
    )

with tabs[7]:
    suivis_anterieurs = st.text_area(
        "Autres bilans et suivis (orthophonie, psy, etc.)",
        placeholder="A déjà bénéficié d'un suivi en orthophonie ainsi que d'un bilan psychologique sans éléments inquiétants.",
    )
    attentes_parents = st.text_area(
        "Attentes des parents vis-à-vis du bilan",
        placeholder="Les parents expriment le souhait que le bilan permette à l'enfant de s'épanouir pleinement.",
    )

# --- LOGIQUE DE RÉDACTION CONDITIONNELLE ---
def generer_texte_anamnese():
    il_elle = "Il" if sexe == "Garçon" else "Elle"
    p = prenom.strip() if prenom.strip() else "L'enfant"

    paragraphes = []

    # Paragraphe 1 : Grossesse & Naissance
    p1 = []
    if grossesse.startswith("Au terme"):
        p1.append(
            f"{p} est né{'e' if sexe=='Fille' else ''} au terme d’une grossesse et"
            " d’un accouchement s’étant déroulés dans la norme."
        )
    elif naissance_details:
        p1.append(f"Concernant la naissance, {naissance_details}.")

    if cesarienne:
        p1.append(
            f"{il_elle} est né{'e' if sexe=='Fille' else ''} par césarienne."
        )
    if naissance_details and grossesse.startswith("Au terme"):
        p1.append(f"{naissance_details}.")

    if allaitement.startswith("Allaitement difficile"):
        p1.append(
            f"L’allaitement a été difficile en raison de difficultés de succion."
            f" {il_elle} est donc rapidement passé{'e' if sexe=='Fille' else ''}"
            " au biberon."
        )
    elif allaitement == "Allaitement sans difficulté":
        p1.append("L'allaitement s'est déroulé sans difficulté.")
    if p1:
        paragraphes.append(" ".join(p1))

    # Paragraphe 2 : Famille & Garde
    p2 = []
    if "unique" in fratrie.lower():
        p2.append(
            f"{p} est {'fille unique' if sexe=='Fille' else 'fils unique'} et"
            f" {chambre.lower()}."
        )
    else:
        p2.append(f"{p} a {fratrie} et {chambre.lower()}.")

    fam_str = f"{il_elle} {schema_familial.lower()}"
    if animaux:
        fam_str += f", ainsi qu’avec {animaux}"
    p2.append(fam_str + ".")

    if langues:
        p2.append(f"À la maison, {langues} sont parlées.")
    if garde:
        p2.append(f"Concernant le mode de garde avant l'école, {garde[0].lower() + garde[1:]}.")
    if p2:
        paragraphes.append(" ".join(p2))

    # Paragraphe 3 : Santé & Sensorialité
    p3 = []
    if sante_globale.startswith("Va bien"):
        p3.append(f"Sur le plan de la santé, {p} va bien.")
    if antecedents_details:
        p3.append(antecedents_details)
    if controles_sensoriels:
        p3.append("Les contrôles visuels et auditifs sont à jour.")
    if vision_audition:
        p3.append(f"{vision_audition}.")
    if sensorialite_tactile:
        p3.append(f"Au niveau sensoriel, {sensorialite_tactile}.")
    if p3:
        paragraphes.append(" ".join(p3))

    # Paragraphe 4 : Vie quotidienne (Alimentation & Sommeil)
    p4 = []
    if alimentation:
        p4.append(alimentation)
    if sommeil:
        p4.append(sommeil)
    if p4:
        paragraphes.append(" ".join(p4))

    # Paragraphe 5 : Développement moteur & Langage
    p5 = []
    if etapes_motrices.startswith("Étapes acquises"):
        p5.append(
            f"Sur le plan moteur, {p} a acquis ses étapes de développement dans"
            " les temps."
        )
    if details_moteur:
        p5.append(details_moteur)
    if langage == "Dans les temps, sans difficulté":
        p5.append("L'acquisition du langage s'est faite dans les temps.")
    elif langage == "Un peu tardive":
        p5.append("L’acquisition du langage a été un peu plus tardive.")
    if details_langage:
        p5.append(details_langage)
    if p5:
        paragraphes.append(" ".join(p5))

    # Paragraphe 6 : Autonomie, Écrans & Loisirs
    p6 = []
    if autonomie:
        p6.append(f"Sur le plan de l’autonomie quotidienne, {autonomie[0].lower() + autonomie[1:]}")
    if ecrans:
        p6.append(ecrans)
    if loisirs:
        p6.append(f"Concernant ses centres d'intérêt et activités, {loisirs[0].lower() + loisirs[1:]}")
    if p6:
        paragraphes.append(" ".join(p6))

    # Paragraphe 7 : Scolarité & Comportement
    p7 = []
    if scolarite:
        p7.append(f"{p} est scolarisé{'e' if sexe=='Fille' else ''} en {scolarite}.")
    if vecu_ecole:
        p7.append(vecu_ecole)
    if sociabilite_ecole:
        p7.append(sociabilite_ecole)
    if detail_motif:
        p7.append(detail_motif)
    if caractere:
        p7.append(f"Selon {presents}, {caractere[0].lower() + caractere[1:]}")
    if attentes_parents:
        p7.append(attentes_parents)
    if p7:
        paragraphes.append(" ".join(p7))

    return paragraphes


# --- GÉNÉRATION DU FICHIER WORD ---
def creer_docx(paragraphes):
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    titre = doc.add_paragraph()
    run_t = titre.add_run("CABINET DE PSYCHOMOTRICITÉ LE PECQ\n")
    run_t.bold = True
    run_t.font.size = Pt(14)
    run_t.font.color.rgb = RGBColor(0x2B, 0x7B, 0xB9)
    titre.add_run("34 Rue du Président Wilson, 78230, Le Pecq")

    doc.add_paragraph(
        f"BILAN PSYCHOMOTEUR - Réalisé le : {date.today().strftime('%d/%m/%Y')}"
    )

    table = doc.add_table(rows=4, cols=2)
    table.cell(0, 0).text = f"Nom : {nom.upper()}"
    table.cell(0, 1).text = f"Prénom : {prenom}"
    table.cell(1, 0).text = f"Date de naissance : {ddn.strftime('%d/%m/%Y')}"
    table.cell(1, 1).text = f"Âge : {age_texte}   |   Scolarité : {scolarite}"
    table.cell(2, 0).text = f"Indication : {indication}"
    table.cell(3, 0).text = f"Médecin prescripteur : {prescripteur}"
    table.cell(3, 1).text = f"Médecin généraliste : {generaliste}"

    doc.add_paragraph()

    h_anam = doc.add_paragraph()
    r_anam = h_anam.add_run("Anamnèse :")
    r_anam.bold = True
    r_anam.underline = True

    for p in paragraphes:
        doc.add_paragraph(p)

    if suivis_anterieurs.strip():
        h_autre = doc.add_paragraph()
        r_autre = h_autre.add_run("Autres bilans :")
        r_autre.bold = True
        r_autre.underline = True
        doc.add_paragraph(suivis_anterieurs)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# --- VALIDATION ET EXPORT ---
st.divider()
if st.button("✅ Valider et générer le compte-rendu Word", type="primary"):
    texte_genere = generer_texte_anamnese()

    st.subheader("Aperçu du texte généré :")
    for para in texte_genere:
        st.write(para)

    fichier_word = creer_docx(texte_genere)
    nom_fichier = f"Anamnese_{prenom}_{nom}.docx".replace(" ", "_")

    st.download_button(
        label="📥 Télécharger le fichier Word (.docx)",
        data=fichier_word,
        file_name=nom_fichier,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
