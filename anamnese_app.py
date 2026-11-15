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
    "1. Données Générales et Motif",
    "2. Antécédents et développement périnatal",
    "3. Évolution moteur",
    "4. Fonctions vitales et hygiène de vie",
    "5. Sphère affective, comportement et sociale",
    "6. Parcours et vécu scolaire",
    "7. Attentes de la famille",
])

with tabs[0]:
    st.subheader("Information Générale")
    col1, col2, col3, col4 = st.columns(4)
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
        ville = st.text_input("Ville de résidence", placeholder="Paris")
        prescripteur = st.text_input(
            "Médecin prescripteur", placeholder="Dr P. Martin (Paris)"
        )
        generaliste = st.text_input(
            "Médecin généraliste", placeholder="Dr D. Dupont (Paris)"
        )
    with col4:
        ecole = st.text_input("École", placeholder="Saint Anne à Paris")
    st.subheader("Composition familiale")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        garde_enfant = st.selectbox(
            "Garde enfant",
            [
                "ses deux parents",
                "sa mère",
                "son père",
                "Autre",
            ],
        )
        if garde_enfant == "Autre":
            garde_enfant = st.text_input(
                "Précisez la garde de l'enfant", 
                placeholder="Oncle"
            )
        schema_familial = st.selectbox(
            "Schéma familliale",
            [
                "parents unis",
                "séparés",
                "garde alternée",
            ],      
        )                
        nb_enfants = st.number_input("Nombre d'enfants dans la fratrie", min_value=1, value=1, step=1)
        if nb_enfants > 1:
            rang = st.number_input("Position dans la fratrie", min_value=1, max_value=nb_enfants, value=1, step=1)
        else:
            rang = 1
        animaux = st.text_input(
            "Animaux de compagnie (optionnel)",
            placeholder="un chiot",
        )
    with col_f2:
        presents = st.text_input(
            "Personne(s) présente(s) à l'entretien", placeholder="sa maman"
        )

        nb_langues = st.number_input("Nombre de langues parlées à la maison", min_value=1, value=1, step=1)
        options_langues = ["français", "anglais", "espagnol", "arabe", "portugais", "italien", "allemand"]
        if nb_langues > 1:
            langues = st.multiselect("Langues parlées", options=options_langues)
        else:
            saisie_unique = st.text_input("Langue parlée", value="français")
            langues = [saisie_unique] if saisie_unique else []

    st.subheader("Motif")
    orientateur = st.text_input(
            "Qui a orienté vers la psychomotricité ?",
            placeholder="L'enseignante / Le médecin",
    )
    indication = st.text_area(
        "Raison de la demande",
        placeholder="un très fort dynamisme au quotidien qui demande beaucoup d’attention",
    )
    date_motif = st.text_input(
        "Depuis quand les difficultés existent",
        placeholder="3 ans",
    )
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        options_diagnostics = ["TDAH", "TSA", "TDC", "Dysgraphie", "hypotonie", "hypertonie"]
        nb_diagnostics_pose = st.number_input("Nombre de diagnostics posé", min_value=0, value=0, step=1)
        diagnostics_pose = []
        if nb_diagnostics_pose > 0:
                diagnostics_pose = st.multiselect("Diagnostic posé", options=options_diagnostics)
    with col_g2:
        nb_diagnostics_en_cours = st.number_input("Nombre de diagnostics en cours", min_value=0, value=0, step=1)
        diagnostics_en_cours = []
        if nb_diagnostics_en_cours > 0:
                diagnostics_en_cours = st.multiselect("Diagnostic en cours", options=options_diagnostics)

with tabs[1]:
    
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
    suivis_anterieurs = st.text_area(
        "Autres bilans et suivis (orthophonie, psy, etc.)",
        placeholder="A déjà bénéficié d'un suivi en orthophonie ainsi que d'un bilan psychologique sans éléments inquiétants.",
    )
    attentes_parents = st.text_area(
        "Attentes des parents vis-à-vis du bilan",
        placeholder="Les parents expriment le souhait que le bilan permette à l'enfant de s'épanouir pleinement.",
    )


# Fonction pour gérer l'élision (le -> l' si voyelle)
    def formater_langue(langue):
        voyelles = ("a", "e", "i", "o", "u", "y", "é", "è", "ê", "ë", "à", "â")
        if langue.lower().startswith(voyelles):
            return f"l'{langue}"
        else:
            return f"le {langue}"

# --- LOGIQUE DE RÉDACTION CONDITIONNELLE ---
def generer_texte_anamnese():
    il_elle = "Il" if sexe == "Garçon" else "Elle"
    article = "le" if sexe == "Garçon" else "la"
    p = prenom.strip() if prenom.strip() else "L'enfant"
    
    # Formatage de la date de naissance et de la date du jour
    ddn_str = ddn.strftime('%d/%m/%Y')
    date_jour = date.today().strftime('%d/%m/%Y')

    paragraphes = []

    # 1. Identité
    p1 = []
    identite_parts = [f"**Identité :** {p} {nom} "]
    identite_parts.append(f"né{'e' if sexe == 'Fille' else ''} le {ddn_str}")
    if age_texte:
        identite_parts.append(f"âgé(e) de {age_texte}")
    if scolarite:
        ecole_str = f" à {ecole}" if ecole else ""
        identite_parts.append(f"scolarisé(e) en classe de {scolarite}{ecole_str}")
    p1.append(", ".join(identite_parts) + ".")
    paragraphes.append(" ".join(p1))

    # 2 : Configuration familiale
    p2 = []
    famille_parts = [f"**Configuration familiale :**  {p}"]
    famille_parts.append(f"vit auprès de {garde_enfant} dans le cadre d'un schéma familial {schema_familial}.")
    p2.append(", ".join(famille_parts))
    if nb_enfants == 1:
        statut_fratrie = "fils unique" if sexe == "Garçon" else "fille unique"
        frat = f"{il_elle} est {statut_fratrie}"
    else:
        if rang == 1:
            ordinal = "1er" if sexe == "Garçon" else "1re"
        else:
            ordinal = f"{rang}e"
        frat = f"{il_elle} est {article} {ordinal} d'une fratrie de {nb_enfants} enfants"
    p2.append(frat + ".")
    if animaux:
        p2.append(f"{il_elle} vit aussi avec {animaux}.")
    if len(langues) == 1:
        p2.append(f"La langue parlée au domicile est {formater_langue(langues[0])}.")
    else:
        langues_formatees = [formater_langue(l) for l in langues]
        if len(langues_formatees) == 2:
            str_langues = f"{langues_formatees[0]} et {langues_formatees[1]}"
        else:
            str_langues = ", ".join(langues_formatees[:-1]) + f" et {langues_formatees[-1]}"
        p2.append(f"Les langues parlées au domicile sont {str_langues}.")
    paragraphes.append(" ".join(p2))

    # 3 : Motif
    p3 = [f"**Motif :**"]
    p3.append(f"L'anamnèse est réalisée le {date_jour} en présence de {presents}.")
    p3.append(f"La consultation fait suite à une orientation de {orientateur}")
    p3.append(f"motivée par {indication},")
    p3.append(f"difficultés observées depuis {date_motif}.")
    
    # Génération des phrases de diagnostics
    diag_phrases = []
    # 1. Diagnostics posés
    diags_poses_valides = [d.strip() for d in diagnostics_pose if d and d.strip()]
    if len(diags_poses_valides) == 1:
        diag_phrases.append(f"Un diagnostic de {diags_poses_valides[0]} est posé.")
    elif len(diags_poses_valides) > 1:
        if len(diags_poses_valides) == 2:
            str_dp = f"{diags_poses_valides[0]} et {diags_poses_valides[1]}"
        else:
            str_dp = ", ".join(diags_poses_valides[:-1]) + f" et {diags_poses_valides[-1]}"
        diag_phrases.append(f"Des diagnostics de {str_dp} sont posés.")

    # 2. Diagnostics en cours
    diags_cours_valides = [d.strip() for d in diagnostics_en_cours if d and d.strip()]
    if len(diags_cours_valides) == 1:
        diag_phrases.append(f"Un diagnostic de {diags_cours_valides[0]} est en cours.")
    elif len(diags_cours_valides) > 1:
        if len(diags_cours_valides) == 2:
            str_dc = f"{diags_cours_valides[0]} et {diags_cours_valides[1]}"
        else:
            str_dc = ", ".join(diags_cours_valides[:-1]) + f" et {diags_cours_valides[-1]}"
        diag_phrases.append(f"Des diagnostics de {str_dc} sont en cours.")

    p3.append(" ".join(diag_phrases))
    paragraphes.append(" ".join(p3))


    # Paragraphe 4 : Grossesse & Naissance
    p4 = []
    if grossesse.startswith("Au terme"):
        p4.append(
            f"{p} est né{'e' if sexe=='Fille' else ''} au terme d’une grossesse et"
            " d’un accouchement s’étant déroulés dans la norme."
        )
    elif naissance_details:
        p4.append(f"Concernant la naissance, {naissance_details}.")

    if cesarienne:
        p4.append(
            f"{il_elle} est né{'e' if sexe=='Fille' else ''} par césarienne."
        )
    if naissance_details and grossesse.startswith("Au terme"):
        p4.append(f"{naissance_details}.")

    if allaitement.startswith("Allaitement difficile"):
        p4.append(
            f"L’allaitement a été difficile en raison de difficultés de succion."
            f" {il_elle} est donc rapidement passé{'e' if sexe=='Fille' else ''}"
            " au biberon."
        )
    elif allaitement == "Allaitement sans difficulté":
        p4.append("L'allaitement s'est déroulé sans difficulté.")
    if p4:
        paragraphes.append(" ".join(p4))

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

    # Boucle corrigée pour transformer le markdown (**en gras**) en vrai gras Word
    for p in paragraphes:
        para_doc = doc.add_paragraph()
        if p.startswith("**") and "**" in p[2:]:
            fin_bold = p.find("**", 2)
            label = p[2:fin_bold]
            reste = p[fin_bold+2:].strip()
            
            run_label = para_doc.add_run(label + " ")
            run_label.bold = True
            para_doc.add_run(reste)
        else:
            para_doc.add_run(p)

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
