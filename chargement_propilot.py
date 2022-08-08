import json
import os
import fnmatch
from datetime import datetime
from urllib.request import urlopen
import urllib
import pandas as pd
import re

forbidden_period_value = ["Y", "Total"]

file_list = ["dim_activity",
             "dim_tree_nodes",
             "dim_top_levels",
             "dim_maturities",
             "dim_period",
             "dim_snapshots",
             "dim_effects",
             "dim_properties",
             "dim_states",
             "dim_structures",
             "fact_financials",
             "fact_property_values"]

month_dict = {"January": "Janvier",
              "February": "Février",
              "March": "Mars",
              "April": "Avril",
              "May": "Mai",
              "June": "Juin",
              "July": "Juillet",
              "August": "Août",
              "September": "Septembre",
              "October": "Octobre",
              "November": "Novembre",
              "December": "Décembre"}

dict_indicateur = {
    'Nombre de repas servis dans les restaurants universitaires au tarif à 1€': "Nombre de repas servis",
    'Montant cumulé de l’investissement total ainsi déclenché': "Montant cumumé de l'investissement total",
    'Nombre d’entreprises bénéficiaires du dispositif': "Nombre d'entreprises bénéficiaires",
    'Nombre de bâtiments Etat dont le marché de rénovation est notifié': 'Nombre de bâtiments dont le marché de rénovation est notifié',
    'Nombre de projets incluant une transformation de la ligne de production pour réduire son impact environnemental': 'Nombre de projets',
    'Nombre d\'exploitations certifiées "haute valeur environnementale"': 'Nombre d’exploitations certifiées',
    'Emissions de gaz à effet de serre évitées sur la durée de vie des équipements': 'Emissions de gaz à effet de serre évitées',
    'Nombre de bonus octroyés à des véhicules électriques et hybrides rechargeables': 'Nombre de bonus octroyés à des véhicules électriques',
    "Quantité de matières plastiques évitées ou dont le recyclage ou l'intégration a été soutenue": 'Quantité de matières plastiques évitées',
    # 'Montant total des travaux associés aux dossiers validés' : 'Montant total des travaux',
    'Nombre de nouveaux projets (nouvelle ligne, extension de ligne et pôle)': 'Nombre de nouveaux projets',
    'Montant de l’investissement total déclenché': 'Montant de l’investissement total',
    'Nombre de projets de tourisme durable financés': 'Nombre de projets de tourisme durable financés',
    'Nombre de projets de rénovation de cathédrales et de monuments nationaux initiés': 'Nombre de projets de rénovation',
    'Montant total investi pour la rénovation de monuments historiques appartenant aux collectivités territoriales': 'Montant total investi pour la rénovation',
    'Nombre de projets de rénovation de monuments historiques appartenant aux collectivités territoriales bénéficiaires initiés': 'Nombre de projets de rénovation',
    "Nombre de contrats de professionnalisation bénéficiaires de l'aide exceptionnelle": 'Nombre de contrats de professionnalisation',
    "Nombre de contrats d'apprentissage bénéficiaires de l'aide exceptionnelle": 'Nombre de contrats d’apprentissage',
    "Nombre d'aides à l'embauche des travailleurs handicapés": "Nombre d'aides à l'embauche des travailleurs handicapés",
    'Nombre de projets locaux soutenus  (rénovation, extension, création de lignes': 'Nombre de projets locaux soutenus',
    # 'Nombre de dossiers MaPrimeRénov validés': 'Nombre de dossiers MaPrimeRénov bénéficiaires ("particulier" et "copropriété")',
    # 'Nombre de dossiers MaPrimeRénov bénéficiaires ("particulier" et "copropriété")': 'Nombre de bénéficiaires',
    "Nombre de dossiers MaPrimeRénov' payés": "Nombre de dossiers payés",
    "Nombre d'entreprises bénéficiares": "Nombre d'entreprises bénéficiaires"
}

dict_mesures = {
        "Appels à projets dédiés à l'efficacité énergétique et à l'évolution des procédés en faveur de la décarbonation de l'industrie": "AAP Efficacité énergétique",
        "CIE jeunes": "Contrats Initiatives Emploi (CIE) Jeunes",
        'France Num': 'France Num : aide à la numérisation des TPE,PME,ETI',
        'Guichet efficacité énergétique dans industrie': 'Guichet efficacité énergétique',
        "Modernisation des filières automobiles et aéronautiques": "Modernisation des filières automobiles et aéronautiques",
        "PEC jeunes": "Parcours emploi compétences (PEC) Jeunes",
        'Relocalisation : soutien aux projets industriels dans les territoires': 'AAP Industrie : Soutien aux projets industriels territoires',
        'Relocalisation : sécurisation des approvisionnements critiques': 'AAP Industrie : Sécurisation approvisionnements critiques',
        'Renforcement des subventions de Business France (chèque export, chèque VIE)': 'Renforcement subventions Business France',
        "Soutien à la modernisation industrielle et renforcement des compétences dans la filière nucléaire": "AAP industrie : modernisation industrielle et renforcement des compétences dans la filière nucléaire",
        'Soutien à la recherche aéronautique civil': 'Soutien recherche aéronautique civil',
        'Rénovation bâtiments Etats': 'Rénovation des bâtiments Etats (marchés notifiés)'
    }


def import_json_to_dict(url):
    response = urllib.request.urlopen(url)
    my_dict = json.loads(response.read())
    return my_dict


def clean_str(s: str) -> str:
    """
    remplacement de caractères spéciaux par d'autre pour faciliter le traitement
    """
    d = {
        "’": "'",
        "\xa0": " ",
        "/": ","
    }
    for x in d:
        s = s.replace(x, d[x]).strip()
    return s


def format_date(raw_date: int) -> str:
    """
    Convertie un format (20210131, 2020) -> (2021-01-31, 2020)
    """
    str_date = str(raw_date)
    if str_date == 'nan':
        return raw_date

    if len(str_date) == 8:
        return str_date[:4] + '-' + str_date[4:6] + '-' + str_date[6:]
    else:
        return str_date


def extract_dep_code(expr: str) -> str:
    """
    Récupération du code de département en retirant les caractères non int de la variable expr, ex 'THD-D72' -> '72'
    """
    nums = re.findall(r'D\d+', expr)
    if expr.endswith('D2A'):
        return '2A'
    elif expr.endswith('D2B'):
        return '2B'
    elif expr.endswith('E00'):
        return '00'
    return nums[0][1:].zfill(2) if len(nums) > 0 else None


def clean_mesure_name(tree_node_name: str) -> str:
    """
    garde seulement la mesure et retire le département de three_node_name
    """
    raw_mesure = tree_node_name.split('/')[1].strip() if '/' in tree_node_name else tree_node_name
    # nettoyage de la colonne mesure, on enlève un point surnuméraire.
    mesure = re.sub('\.', "", raw_mesure)
    mesure = clean_str(mesure)
    return mesure


def get_df_sum_indicator(df_dep: pd.DataFrame,
                         indicators_to_sum: str,
                         new_indicator: str,
                         new_indic: str,
                         new_mesure: str):
    # TODO: doc func
    df_dep = df_dep.copy()
    df_temp = df_dep.loc[df_dep.indicateur.str.contains(indicators_to_sum, regex=True)].groupby(
        ["Date", "dep"]).sum().copy()
    df_temp["indicateur"] = new_indicator + " - " + new_indic
    df_temp["short_indic"] = new_indicator
    df_temp["mesure"] = new_mesure
    df_temp = (df_temp
               .merge(df_dep.drop(columns=["indicateur", "indic_id", "short_indic", "mesure", "short_mesure", "valeur"])
                      .drop_duplicates(["Date", "dep"]),
                      on=["Date", "dep"]))
    df_temp["short_mesure"] = new_mesure
    df_temp["indic_id"] = new_indic
    # df_temp.fillna("NaN", inplace=True)
    return df_temp


def recup_date(string: str) -> str:
    """
    retire les informations inutile présente dans la date '2020-12-31T00:00:00.0000000' -> '2020-12-31'
    """
    return string[:10]


def mkdir_ifnotexist(path: str):
    """
    créer le dossier 'export' si il n'existe pas
    """
    if not os.path.isdir(path):
        os.mkdir(path)


def load_propilot():
    taxo_dep_df = pd.read_csv('refs/taxo_deps.csv', dtype={'dep': str, 'reg': str})
    taxo_dep_df['dep'] = taxo_dep_df['dep'].apply(lambda x: x.zfill(2))
    taxo_dep_df['reg'] = taxo_dep_df['reg'].apply(lambda x: x.zfill(2))
    dep_list = list(taxo_dep_df['dep'].unique())

    taxo_reg_df = pd.read_csv('refs/taxo_regions.csv', dtype={'reg': str})
    taxo_reg_df['reg'] = taxo_reg_df['reg'].apply(lambda x: x.zfill(2))
    reg_list = list(taxo_reg_df['reg'].unique())

    data_dir_path = './data/'

    propilot_path = os.path.join("data")

    df_dict = {}

    for file in file_list:
        for file_csv in os.listdir(propilot_path):
            if fnmatch.fnmatch(file_csv, file + "*.csv"):
                print("File loaded : ", file_csv)
                df_dict[file] = pd.read_csv(os.path.join(propilot_path, file_csv), sep=";")

    df_dict['fact_financials']['period_id'] = df_dict['fact_financials']['period_id'].apply(lambda x: format_date(x))
    df_dict['dim_period']['period_id'] = df_dict['dim_period']['period_id'].apply(lambda x: format_date(x))

    # Retirer les lignes ayant un period quarter year commençant par 11 -> nonsense
    df_dict['dim_period'] = df_dict['dim_period'][
        ~df_dict['dim_period']['period_quarter_year'].str.startswith('11', na=False)]

    # Filtrer les valeurs nulles pour alléger fact_financials
    df_dict["fact_financials"] = df_dict["fact_financials"].loc[
        ~df_dict["fact_financials"].financials_cumulated_amount.isna()]

    df = (df_dict["fact_financials"]
          .merge(df_dict["dim_tree_nodes"], left_on="tree_node_id", right_on="tree_node_id")
          .merge(df_dict["dim_effects"], left_on="effect_id", right_on="effect_id")
          .merge(df_dict["dim_states"], left_on="state_id", right_on="state_id")
          .merge(df_dict["dim_period"], left_on="period_id", right_on="period_id", how='left')
          .merge(df_dict["dim_structures"], left_on="structure_id", right_on="structure_id"))

    df['dep_code'] = df['tree_node_code'].apply(lambda x: extract_dep_code(x))

    cols = ["tree_node_name", "structure_name", "effect_id", "state_id", "period_date", "period_month_tri",
            "period_month_year", "financials_cumulated_amount", "dep_code"]

    df = df[cols]
    df.rename(columns={"period_month_year": "Date", "financials_cumulated_amount": "valeur"}, inplace=True)

    df.rename(columns={"effect_id": "indicateur"}, inplace=True)
    df.indicateur = df.indicateur.str.strip()
    df["short_indic"] = df.indicateur.apply(lambda x: x.split("-")[0].strip())
    df["indic_id"] = df.indicateur.apply(lambda x: x.split("-")[-1].strip())

    df = df.loc[(~df.period_month_tri.isin(forbidden_period_value)) &
                (df.state_id == 'Valeur Actuelle') &
                (~df.valeur.isna())].copy()

    # Filtrer les lignes ayant une date ultérieure à la date d'aujourd'hui
    today = datetime.today().strftime('%Y-%m-%d')

    date_series = df['period_date'].apply(lambda x: x.split('T')[0])
    df['format_date'] = pd.to_datetime(date_series)

    df = df[df['format_date'] <= today]
    df = df.drop('format_date', axis=1)

    df["departement"] = df["tree_node_name"].apply(lambda x: x.split("/")[0].strip())
    df["mesure"] = df["tree_node_name"].apply(lambda x: x.split("/")[-1].strip())
    df.drop(columns=["tree_node_name"], inplace=True)

    # nettoyage de la colonne mesure, on enlève un point surnuméraire.
    df["mesure"].replace("\.", "", regex=True, inplace=True)
    df.mesure = df.mesure.apply(lambda x: clean_str(x))

    # traduit les mois dans la colonne date
    df.Date = df.Date.replace(month_dict, regex=True)

    # enrichit avec les noms de département/région
    df = df.merge(taxo_dep_df[["dep", "reg", "libelle"]],
                  how="left", left_on="dep_code", right_on="dep") \
        .merge(taxo_reg_df[["reg", "nccenr"]], how="left", left_on="reg", right_on="reg")

    df.rename(columns={"nccenr": "region"}, inplace=True)
    df.drop(['dep_code'], axis=1, inplace=True)  # Supprime code dep utilisé pour la jointure

    df_dep = df.loc[df.structure_name == "Département"].copy()
    df_dep.drop(columns=["structure_name"], inplace=True)

    dep_indic = set(df_dep.indicateur.unique())
    df_reg = df.loc[(df.structure_name == "Région")
                    & (~df.indicateur.isin(dep_indic))].copy()
    df_reg.drop(columns=["structure_name"], inplace=True)

    reg_indic = set(df_reg.indicateur.unique())
    df_nat = df.loc[(df.structure_name == "Mesure")
                    & (~df.indicateur.isin(reg_indic))
                    & (~df.indicateur.isin(dep_indic))].copy()
    df_nat.drop(columns=["structure_name"], inplace=True)

    df_dep["short_indic"] = df_dep.indicateur.apply(lambda x: x.split("-")[0].strip())
    df_dep.short_indic = df_dep.short_indic.apply(lambda x: clean_str(x))

    df_reg["short_indic"] = df_reg.indicateur.apply(lambda x: x.split("-")[0].strip())
    df_reg.short_indic = df_reg.short_indic.apply(lambda x: clean_str(x))

    df_nat2 = df_nat.copy()  # Opérations faites sur df_nat. df_nat utile pour autres choses.
    df_nat2["short_indic"] = df_nat2.indicateur.apply(lambda x: x.split("-")[0].strip())
    df_nat2.short_indic = df_nat2.short_indic.apply(lambda x: clean_str(x))

    df_dep.short_indic = df_dep.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)
    df_dep.short_indic = df_dep.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)
    df_dep.short_indic = df_dep.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)

    df_reg.short_indic = df_reg.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)
    df_reg.short_indic = df_reg.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)
    df_reg.short_indic = df_reg.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)

    df_nat2.short_indic = df_nat2.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)
    df_nat2.short_indic = df_nat2.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)
    df_nat2.short_indic = df_nat2.short_indic.apply(lambda x: dict_indicateur[x] if x in dict_indicateur else x)

    df.rename(columns={"effect_id": "indicateur"}, inplace=True)
    df.indicateur = df.indicateur.str.strip()

    df_dep["short_mesure"] = df_dep.mesure.apply(lambda x: dict_mesures[x] if x in dict_mesures else x)
    df_reg["short_mesure"] = df_dep.mesure.apply(lambda x: dict_mesures[x] if x in dict_mesures else x)
    df_nat2["short_mesure"] = df_dep.mesure.apply(lambda x: dict_mesures[x] if x in dict_mesures else x)

    df_temp = get_df_sum_indicator(
        df_dep,
        "PEE3|EEI1|CBC3",
        "Nombre d’entreprises ayant reçu l’aide",
        "SSS1",
        "Décarbonation de l'industrie (Appel à projets EE + Guichet EE + Chaleur bas carbone)")
    df_temp2 = get_df_sum_indicator(
        df_dep,
        "PEE2|CBC2",
        "Montant cumulé de l’investissement total ainsi déclenché",
        "SSS2",
        "Décarbonation de l'industrie (Appel à projets EE + Guichet EE + Chaleur bas carbone)")

    df_dep = pd.concat([df_dep, df_temp, df_temp2])

    df_dep.to_csv("pp_dep.csv", sep=";")

    df_nat.to_csv("pp_nat.csv", sep=";")

    df_nat2["maille"] = "Nationale"
    df_reg["maille"] = "Régionale"
    df_reg = df_reg.merge(taxo_reg_df[["nccenr", "reg"]], how='left', left_on='departement', right_on='nccenr')
    df_reg = df_reg.drop(["nccenr", 'reg_x'], axis=1)
    df_dep["maille"] = "Départementale"

    df_dep = df_dep.rename(columns={"departement": "localisation", 'dep': 'Code_Departement', 'reg': 'Code_Region'})
    df_reg = df_reg.rename(columns={"departement": "localisation", 'reg_y': 'Code_Region'})
    df_nat2 = df_nat2.rename(columns={"mesure": "localisation", "departement": "mesure"})
    df_nat2["localisation"] = "Nationale"

    df_all = pd.concat([df_dep, df_reg, df_nat2])
    df_all.reset_index(drop=True, inplace=True)

    columns = ['dep', 'reg', 'libelle', 'region', 'state_id', 'short_mesure']
    df_all.drop(columns, inplace=True, axis=1)
    df_all = df_all.rename(columns={'period_month_tri': 'abrev_mois'})

    df_all["period_date"] = df_all["period_date"].apply(recup_date)

    with open(os.path.join("refs", "volet2quadri.json")) as f:
        clef_indicateur2volet = json.load(f)

    ecologie2indic = clef_indicateur2volet["ecologie"]
    competitivite2indic = clef_indicateur2volet["competitivite"]
    cohesion2indic = clef_indicateur2volet["cohesion"]
    df1 = pd.DataFrame(ecologie2indic, columns=["code"])
    df1["volet"] = "Ecologie"

    df2 = pd.DataFrame(competitivite2indic, columns=["code"])
    df2["volet"] = "Compétitivité"

    df3 = pd.DataFrame(cohesion2indic, columns=["code"])
    df3["volet"] = "Cohésion"

    df = pd.concat([df1, df2, df3])

    df_all = df_all.merge(df, how="left", left_on="indic_id", right_on="code")

    df_mesure = pd.read_csv("refs/INDICATEURS.csv", sep=';')
    L = [["DVP1", "Transition écologique, transports, et logement"],
         ["ETE1", "Transition écologique, transports, et logement"],
         ["ETE2", "Transition écologique, transports, et logement"]]
    for mesure in df.code.unique():
        code = mesure[:-1]
        if code in df_mesure["Code-4"].unique():
            L += [[mesure, df_mesure[df_mesure["Code-4"] == code]["Ministères"].iloc[0]]]

    df_to_merge = pd.DataFrame(L, columns=["code", "Ministères"])
    df_all = df_all.merge(df_to_merge, how='left', left_on='code', right_on='code')

    df = df_all
    df.period_date = pd.to_datetime(df.period_date)
    df.set_index(["localisation", "period_date"], inplace=True)
    df = df[~df.index.duplicated()]  # this fix the ci by removing duplicate index in the multiIndex

    result = []
    indic_ids = df.indic_id.unique()

    for indic_id in indic_ids:
        temp = df.loc[(df.indic_id == indic_id)]

        new_index = pd.MultiIndex.from_product(temp.index.remove_unused_levels().levels)
        # print(list(temp.index))
        # print(list(temp.index.levels))
        new_temp = temp.reindex(new_index)

        temp2 = new_temp.groupby(level=0).apply(lambda x: x.reset_index(level=0, drop=True).asfreq("M"))

        for level in temp2.index.get_level_values(level=0).unique():
            # print(l)
            temp2.loc[level, "valeur"].fillna(method='ffill', inplace=True)
            temp2.loc[level, "valeur"].fillna(0, inplace=True)
            temp2.loc[level, "Date"] = temp2.loc[level, :].index.strftime("%B %Y")
            temp2.loc[level, :].fillna(method="ffill", inplace=True)
            temp2.loc[level, :].fillna(method="bfill", inplace=True)
        result.append(temp2)

    df_all = pd.concat(result)

    df_all.reset_index(inplace=True)

    mkdir_ifnotexist('exports')
    df_all.to_csv(os.path.join('exports', 'propilot.csv'), sep=';')


if __name__ == "__main__":
    load_propilot()
