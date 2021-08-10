def get_comment(content, volet2mesures):
    volet2comment = {}
    mesure2comment = {}
    body = content.body[:min(len(content.body),142)] #Supprime le dernier retour à la ligne parasite
    # Les mesures doivent apparaitre dans le même ordre que le document
    ordered_mesures = [mesure for mesures in volet2mesures.values() for mesure in mesures]
    compteur_mesure = 0 #Permet de se retrouver dans les mesures affiliées aux commentaires
    voletecologie = 0
    voletcompetitivite = 46
    voletcohesion = 92
    liste_indice_volet = [voletecologie, voletcompetitivite, voletcohesion]
    liste_volet = ["Ecologie", "Competitivite", 'Cohesion']
    # Boucle while pour sortir texte_content
    position = 0
    while position < 100: #Dernier commentaires correspond au volet cohesion indice 92
        textbox_content = ""
        while textbox_content == "" and position < 100:
            if position in liste_indice_volet:
                position += 1
                textbox_content = body[position][0][0]
                position += 1
            else:
                # Dans une page normée, il y a 5 entités. La 6e entité est vide si il y a un commentaire après, ou ne l'est pas si on change de page
                if body[position + 6] == [[[""]]]:
                    position += 7 #Zone du commentaire
                    textbox_content = body[position][0][0]
                    position += 1 # On rajoute 1: Retour sur la première ligne de la page suivante
                else:
                    position += 6
                    compteur_mesure += 1
        #Traitement sur le textbox_content
        # Filtrer des retours à la lignes et potentiels num page
        while len(textbox_content) > 0 and (textbox_content[0] == '' or textbox_content[0].strip().isdigit()):
            textbox_content = textbox_content[1:]
                
        # On extrait le commentaire
        textbox_content = extract_comment(textbox_content)
        textbox_content.replace(";", ",")
        # On associe la mesure au commentaire
        if position-2 in liste_indice_volet: #On traite un commentaire de volet
            encode_volet = encode_name(liste_volet[liste_indice_volet.index(position-2)])
            volet2comment[encode_volet] = textbox_content
        else:
            encoded_mesure = encode_name(ordered_mesures[compteur_mesure])
            compteur_mesure += 1 
            mesure2comment[encoded_mesure] = textbox_content  #textbox_content

    return mesure2comment, volet2comment


def recuperation_commentaire_all(path_to_transposed_report, path_to_image = image_folder):
    #On ne conserve que les fiches docx
    liste_fichier = os.listdir(path_to_transposed_report)
    dico_commentaire_volet = {}
    dico_commentaire_mesure = {}
    for fichier in liste_fichier:
        if fichier.endswith('.docx'):
            src_filename = os.path.join(path_to_transposed_report, fichier) #Departement du nord 73
            content = docx2python(src_filename, image_folder=image_folder)
            mesurecomment, voletcomment = get_comment(content, volet2mesures)
            dep_name = src_filename.split('_')[-1].split('.docx')[0]
            dico_commentaire_mesure[dep_name] = mesurecomment
            dico_commentaire_volet[dep_name] = voletcomment
    return dico_commentaire_mesure, dico_commentaire_volet


path_to_transposed_report = os.path.join(os.getcwd(), 'reports_word', 'transposed_reports')
dico_commentaire_mesure, dico_commentaire_volet = recuperation_commentaire_all(path_to_transposed_report)


L_to_dataframe = []
for dep in list(dico_commentaire_mesure.keys()):
    L_dep = [dep]
    for key in list(volet2mesures.keys()):
        for mesure in volet2mesures[key]:
            com_volet = dico_commentaire_volet[dep][encode_name(key)]
            try:
                com_mesure = dico_commentaire_mesure[dep][encode_name(mesure)]
                L_dep = [dep, encode_name(key), com_volet, encode_name(mesure), com_mesure]
            except: # Toutes les mesures n'ont pas de commentaires
                L_dep = [dep, encode_name(key), com_volet, encode_name(mesure), ""] 
            L_to_dataframe += [L_dep]

df = pd.DataFrame(L_to_dataframe)
df.columns = ["Département", "Volet", "Commentaire_volet", "Mesure", "Commentaire_mesure"]


volet2code_mesures = {
    'Ecologie': ["MPR2", "MPR4", "BOE1", "DVP1", "RBC3", "RBE1", "AEA1", "FFR1", "BPI1", "BPI2"],  #MPR et BPI x2
    'Compétitivité': ["IDF1", "IDF2", "IDF3", "PIT3", "SAC3", "FUM1", "SFC1", "SBF1"],
    'Cohésion': ["APP1", "PEJ1", "CIE1", "PEC1", "CDP1", "GJE1", "SCI1", "PTH1", "SIL1"],
}

liste_indic = ["MPR2", "MPR4", "BOE1", "DVP1", "RBC3", "RBE1", "AEA1", "FFR1", "BPI1", "BPI2",
               "IDF1", "IDF2", "IDF3", "PIT3", "SAC3", "FUM1", "SFC1", "SBF1",
               "APP1", "PEJ1", "CIE1", "PEC1", "CDP1", "GJE1", "SCI1", "PTH1", "SIL1"]
pp_dep = pd.read_csv("pp_dep.csv", sep=";", dtype={"reg":str, "dep":str})
pp_dep['Date'] = pp_dep.Date.apply(lambda x: re.sub(' +', ' ', x))
L_code = []
for code in liste_indic:
    try:
        print(code)
        mesure_comp = pp_dep[pp_dep.indic == code].mesure.iloc[0]
        mesure_dico = mesure_comp.replace(' ', '').replace('é', 'e').replace("'", "").replace("à", "a").replace("â","a").replace("(","").replace(")","").replace("è","e").replace(",", "").lower()
        if "relocalisation:soutienauxprojetsindustrielsdanslesterritoires" in mesure_dico:
            mesure_dico = "aapindustriesoutienauxprojetsindustrielsterritoires"
        if "relocalisation:securisationdesapprovisionnementscritiques" in mesure_dico:
            mesure_dico = "aapindustriesecurisationapprovisionnementscritiques"
        if "francenum" in mesure_dico:
            mesure_dico = "francenumaidealanumerisationdestpepmeeti"
        if "renforcementdessubventionsdebusinessfrancechequeexportchequevie" in mesure_dico:
            mesure_dico = "renforcementsubventionsbusinessfrance"
        if "ciejeunes" in mesure_dico:
            mesure_dico = "contratsinitiativesemploiciejeunes"
        if "pecjeunes" in mesure_dico:
            mesure_dico = "parcoursemploicompetencespecjeunes"
        indicateur = pp_dep[pp_dep.indic == code].indicateur.iloc[0]
        L_code += [[code, mesure_comp, mesure_dico, indicateur]]
    except:
        pass

df_to_merge = pd.DataFrame(L_code)
df_to_merge.columns = ["Quadrigramme", "mesure", "mesure_without_space", 'indicateurs']

df = pd.merge(df, df_to_merge, how='left', left_on='Mesure', right_on='mesure_without_space')
df = df.drop(["mesure", "mesure_without_space"], axis=1)
# Rajout de la date à laquelle les commentaires ont été fait
df['Date'] = "Mai 2021"
# Renommage d'un département pour la jointure future
df.Département = np.where(df.Département == "Val-d'Oise", "Val-D'Oise", df.Département)
# Récuperation des régions dans pp_dep
df_to_merge = pp_dep.filter(["departement", "region"])
df = pd.merge(df, df_to_merge, how='left', left_on="Département", right_on="departement")
df.drop_duplicates(keep = 'first', inplace=True)
df = df.drop(["departement"], axis=1)
df.to_csv("enrichissement_commentaireV0.csv", sep=";", index=False)