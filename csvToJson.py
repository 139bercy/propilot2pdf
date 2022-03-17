import datetime

import pandas as pd
import json
import os


def convert_csv_to_json():

    col = ["indicateur",
           "period_date",
           "valeur",
           "mesure",
           "short_indic",
           "maille",
           "indic_id",
           "Code_Departement",
           "Code_Region"]
    df_propilot = pd.read_csv("exports/propilot.csv", usecols=col, sep=";")
    # avoid null values
    df_propilot = df_propilot[~df_propilot.indicateur.isna()]
    df_propilot['period_date'] = pd.to_datetime(df_propilot['period_date'])

    file_name = 'france-relance-data-tableau-de-bord.txt'
    clean_file(file_name)

    # Récupération des indicateurs uniques
    for indicateur in df_propilot.indicateur.unique():
        df_indicateur = df_propilot.loc[df_propilot.indicateur == indicateur]

        data = {"code": df_indicateur["indic_id"].iloc[0],
                "nom": df_indicateur["short_indic"].iloc[0],
                "unite": df_indicateur["short_indic"].iloc[0]}

        # France
        france = get_level(df_indicateur, "nat", "fra")

        data["france"] = [france]

        # Régions
        regions_data = []
        for region in df_indicateur.Code_Region.unique():
            df_indicateur_region = df_indicateur.loc[df_propilot.Code_Region == region]
            region_data = get_level(df_indicateur_region, "reg", region)
            regions_data.append(region_data)

        data["regions"] = regions_data

        # Départements
        departements_data = []
        for departement in df_indicateur.Code_Departement.unique():
            df_indicateur_departement = df_indicateur.loc[df_propilot.Code_Departement == departement]
            departement_data = get_level(df_indicateur_departement, "dep", departement)
            departements_data.append(departement_data)

        data["departements"] = departements_data

        append_to_file(data, file_name)


def evolVal(valI: float, valE: float) -> float:
    """
    retourne la valeur de l'évolution entre les deux valeurs d'entrées
    """
    if (valE != 0 and valI != valE):
        return valI - valE
    return 0


def evolPercent(ev: float, val: float) -> float:
    """
    retourne le pourcentage d'évolution entre les deux valeurs d'entrées
    """
    if val != 0:
        return round((ev / val)*100, 2)  # round((ev - val) / ev * 100) probablement faux mais pas utilisé plus tard
    return 0


def get_last_data(dff: pd.DataFrame) -> list[datetime.datetime, float]:
    """
    sommes des valeurs de la date la plus récente
    """
    most_recent_date = dff['period_date'].max()
    dfDate = dff.loc[dff.period_date == most_recent_date]
    return [most_recent_date, dfDate["valeur"].sum()]


def get_evolution(dff: pd.DataFrame, last_date: datetime.datetime, last_value: float) -> dict[int, float]:
    dfEvol = dff.copy()
    dfEvol.drop(dfEvol.loc[dfEvol['period_date'] == last_date].index, inplace=True)
    previous_last_data = get_last_data(dfEvol)
    evol = evolVal(last_value, previous_last_data[1])
    evol_percent = evolPercent(evol, previous_last_data[1])
    return [evol, evol_percent]


def get_data_history(dff: pd.DataFrame) -> list:
    """
    retourne la somme de toutes les valeurs pour une date et un indicateur pour une maille (région, département, france)
    """
    dates = dff.sort_values(by="period_date").period_date.unique()
    values = []
    for date in dates:
        value = dff.loc[dff.period_date == date]
        values.append({"date": date.astype(str), "value": value["valeur"].sum()})

    return values


def get_level(df: pd.DataFrame, level: str, code_level: str) -> dict:
    """
    dont know what it does for now
    """
    last_data = get_last_data(df)
    evolution = get_evolution(df, last_data[0], last_data[1])
    data_history = get_data_history(df)
    data_level = {"level": level,
                  "code_level": code_level,
                  "last_value": last_data[1],
                  "last_date": str(last_data[0]),
                  "evol": evolution[0],
                  "evol_percentage": evolution[1],
                  "evol_color": "red",
                  "values": data_history}
    return data_level


def clean_file(file_name: str):
    """
    Suppression du fichier généré précédemment
    """
    try:
        os.remove(file_name)
    except:
        pass


def append_to_file(data: dict, file_name: str):
    """
    creation du fichier txt contenant les valeurs du csv
    """
    with open(file_name, "a", encoding="utf8") as output_file:
        json.dump(data, output_file, ensure_ascii=False)
        output_file.write('\n')


if __name__ == "__main__":
    convert_csv_to_json()
