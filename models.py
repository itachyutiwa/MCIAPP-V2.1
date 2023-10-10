import pandas as pd
import numpy as np
import random
import datetime as dt
import os, xlsxwriter
import warnings
warnings.filterwarnings("ignore")
import dask.dataframe as dd
from pyspark.sql import SparkSession
from pyspark.sql.functions import col
import logging


def to_camel_case(column_name):
    # Supprimez les espaces et les underscores et divisez la chaîne en mots
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])
    if column_name.lower() in ["police"]:
        column_name = "NUM_POLICE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])
    #matricule_beneficiaire
    if column_name.lower() in ["matricule_beneficiaire","matricule beneficiaire"]:
        column_name = "MATRICULE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])
    
    if column_name.lower() in ["libelle_gestionnaire", "libelle gestionnaire"]:
        column_name = "GESTIONNAIRE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])
    
    if column_name.lower() in ["date reception", "date_reception","Date Recepetion"]:
        column_name = "DATE_RECEPTION"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])

    if column_name.lower() in ["date sortie", "date_sortie"]:
        column_name = "DATE_SORTIE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])
    #Date Debut Exercice
    if column_name.lower() in ["Date Debut Exercice", "DATE DEBUT EXERCICE","Date_Debut_Exercice", "DATE_DEBUT_EXERCICE"]:
        column_name = "DATE_DEBUT_EXERCICE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])
    
    if column_name.lower() in ["Date Fin Exercice", "DATE FIN EXERCICE","Date_Fin_Exercice", "DATE_FIN_EXERCICE"]:
        column_name = "DATE_FIN_EXERCICE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])
    
    #['numPolice', 'dateSoins', 'codeActe', 'codeAffectation', 'mtReclame', 'baseRemboursement', 'montantPaye']
    if column_name.lower() in ["CODE_ACTE", "code acte","Code Acte", "code_acte","Code_Acte"]:
        column_name = "CODE_ACTE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])
    
    if column_name.lower() in ["CODE AFFECTION", "code affection","code_affection", "Code Affection","Code_Affection"]:
        column_name = "CODE_AFFECTION"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])
    
    if column_name.lower() in ["MT RECLAME", "montant reclame","montant_reclame", "Montant Reclame","Montant_Reclame"]:
        column_name = "MT_RECLAME"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])
    
    if column_name.lower() in ["BASE REMBOURSEMENT", "base remboursement","base_remboursement", "Base Remboursement","Base_Remboursement"]:
        column_name = "BASE_REMBOURSEMENT"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                  for word in words[1:])
    
    if column_name.lower() in ["MONTANT PAYE", "Montant Paye","Montant_Paye", "montant paye","montant_paye"]:
        column_name = "MONTANT_PAYE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize() for word in words[1:])
    
    return camel_case_name

class Production:
    def __init__(self):
        return
# --------Conformez toutes les variables benef, coti et conso à celles de test.py

    def database(self, file_benef, file_cotisation, file_conso):
        self.file_benef = file_benef
        self.file_cotisation = file_cotisation
        self.file_conso = file_conso
        benef = pd.read_excel(file_benef)
        benef.columns = [to_camel_case(column) for column in benef.columns]

        cotisation = pd.read_excel(file_cotisation)
        cotisation.columns = [to_camel_case(column)
                              for column in cotisation.columns]

        consommation = pd.read_excel(file_conso)
        consommation.columns = [to_camel_case(
            column) for column in consommation.columns]

        # Nombre de police unique benef -- ok
        police_en_gestion = benef.groupby(["numPolice"])["matricule"].agg({
            "count"}).reset_index().shape[0]

        # nombre de police unique cotisation -- ok
        police_facturee = len(cotisation.groupby(["numPolice"]))

        # nombre de ligne cotisation sans doublons -- ok
        piece_de_facturation = cotisation.drop_duplicates().shape[0]

        # somme de beneficiaire par police unique -- ok
        beneficiaire = np.sum(benef.groupby(["numPolice"])[
                              "matricule"].agg({"count"}).reset_index()["count"])

        # somme de MT PRIME NET sans doublons cotisation -- ok
        valeur_facturee = np.sum(cotisation.groupby(["numPolice"])[
                                 "mtPrimeNet"].agg({"sum"}).reset_index()["sum"])
        return {
            "infos":
            {
                "police_en_gestion": police_en_gestion,
                "police_facturee": police_facturee,
                "piece_de_facturation": piece_de_facturation,
                "beneficiaire": beneficiaire,
                "valeur_facturee": valeur_facturee
            },

            "db":
                {
                    "benef": benef,
                    "cotisation": cotisation,
                    "consommation": consommation,
            }
        }

    # Vérifiez si  nombre de police unique benef égale au nombre de police unique cotisation
    # sinon vérifiez si les police de différence consomment

    def production(self, file_benef, file_cotisation, file_conso):
        rs = ""
        benef = self.database(file_benef, file_cotisation,
                              file_conso)["db"]["benef"]
        cotisation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["cotisation"]
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"]

        df1 = benef["numPolice"].dropna().apply(lambda x: str(
            x).replace(" ", "")).drop_duplicates().tolist()
        df2 = cotisation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()
        df3 = consommation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()
        result_1 = [id for id in df1 if id in df3 and id not in df2]
        if len(result_1) == 0:
            rs = rs
        elif len(result_1) != 0 :
            rs = f"{result_1} : bénéficiaire(s) ayant eu des prestations sans cotiser."
        return {
            "production": rs
        }

    # Vérifiez si  nombre de police unique benef égale au nombre de police unique cotisation --ok
    def affiliation_avnants(self, file_benef, file_cotisation, file_conso):
        rs = ""
        benef = self.database(file_benef, file_cotisation,
                              file_conso)["db"]["benef"]
        cotisation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["cotisation"]
        df1 = benef["numPolice"].dropna().apply(lambda x: str(
            x).replace(" ", "")).drop_duplicates().tolist()
        df2 = cotisation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()
        "Les NUM POLICE se trouvant dans beneficiaire  sans cotisation "
        result_1 = [id for id in df1 if id not in df2]

        if len(result_1) == 0:
            rs=rs
        elif len(result_1) != 0:
            f"{result_1} : bénéficiaire(s) sans cotisation."

        return {
            "polices": rs
        }

    # Choisir un nombre aléatoire parmi MT PRIME NET cotisation sans doublons --ok
    # Numéro de Quittance + Date émis policesion associé à soumettre  au courtier pour vérif
    def affiliation_pieces_justificatives(self, file_benef, file_cotisation, file_conso, n_alea=3):
        cotisation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["cotisation"]
        val_alea = list(set(cotisation.drop_duplicates()[
                        "numPolice"].dropna().apply(lambda x: str(x))))
        random.shuffle(list(set(val_alea)))
        liste, n = [], 0
        for i in val_alea:
            while len(liste) != n_alea:
                liste.append(i)
                n = n+1
        infos = {"Affiliation pièces justificatives": {}}
        for i in cotisation["numPolice"]:
            for j in liste:
                if i == j:
                    data = cotisation[cotisation["numPolice"]
                                      == i].drop_duplicates()
                    quittance = [i for i in data["numPolice"]]
                    date_emission = [i for i in data["dateEmission"]]
                    for i, j in zip(quittance, date_emission):
                        infos["Affiliation pièces justificatives"][f"Quittance N°{i}"] = f" {j}"

        return {
            "pieces_justif": list(infos['Affiliation pièces justificatives'].items())
        }

    # Choisir un nombre aléatoire de police  parmi cotisation sans doublons --ok
    def affiliation_facuration(self, file_benef, file_cotisation, file_conso, n_alea=3):

        cotisation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["cotisation"]
        val_alea = list(set(cotisation.drop_duplicates()[
                        "numPolice"].dropna().apply(lambda x: str(x))))
        random.shuffle(val_alea)
        alea = val_alea[:n_alea]
        return {
            "aff_facuration": alea
        }
    # Vérifiez les numéros police benf sans cotisation avec consommation --ok

    def recouvrement(self, file_benef, file_cotisation, file_conso):
        rs = ""
        benef = self.database(file_benef, file_cotisation,
                              file_conso)["db"]["benef"]
        cotisation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["cotisation"]
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"]
        df1 = benef["numPolice"].dropna().apply(lambda x: str(
            x).replace(" ", "")).drop_duplicates().tolist()
        df2 = cotisation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()
        df3 = consommation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()
        

        # 2. Les ID se trouvant dans BENEF et CONSO mais pas dans COTISA
        result_2 = [id for id in df1 if id in df3 and id not in df2]

        # 3. Les ID se trouvant dans CONSO mais pas dans BENEF et COTISA
        result_3 = [id for id in df3 if id not in df1 and id not in df2]
        if len(result_2) == 0 and len(result_3) == 0:
            rs = rs
        elif len(result_2) == 0 and len(result_3) != 0:
            rs = f"{list(set(result_3))} : prestation(s) sans bénéficiaire(s) ni cotisations( cas grave)."
        elif len(result_2) != 0 and len(result_3) == 0:
            rs = f"{list(set(result_2))} : bénéficiaire(s) sans cotisation ayant consommé(s)."
        elif len(result_2) != 0 and len(result_3) != 0:
            rs = f"{list(set(result_2))} : bénéficiaire(s) sans cotisation ayant consommé(s), {list(set(result_3))} : prestation(s) sans bénéficiaire(s) ni cotisations( cas grave)."

        return {
            "recouv":  rs
        }

    def facturation_frais_gestion(self, file_benef, file_cotisation, file_conso):
        return {
            "frais_gestion": "***"
        }

    def resultats(self, file_benef, file_cotisation, file_conso):
        
        production = self.production(file_benef, file_cotisation, file_conso)[
            "production"]
        affiliation_avnants = self.affiliation_avnants(
            file_benef, file_cotisation, file_conso)["polices"]
        affiliation_pieces_justificatives = self.affiliation_pieces_justificatives(
            file_benef, file_cotisation, file_conso, n_alea=3)["pieces_justif"]
        affiliation_facuration = self.affiliation_facuration(
            file_benef, file_cotisation, file_conso, n_alea=3)["aff_facuration"]
        recouvrement = self.recouvrement(
            file_benef, file_cotisation, file_conso)["recouv"]
        facturation_frais_gestion = self.facturation_frais_gestion(
            file_benef, file_cotisation, file_conso)["frais_gestion"]

        rs = {
            "result": [],
            "anorm": []
        }

        if len(production) != 0:
            rs["result"].append("KO, anomalie constatée")
            rs["anorm"].append(production)
        else:
            rs["result"].append("OK")
            rs["anorm"].append("RAS")

        if len(affiliation_avnants) != 0:
            rs["result"].append("KO, anomalie constatée")
            rs["anorm"].append(affiliation_avnants)
        else:
            rs["result"].append("OK")
            rs["anorm"].append("RAS")

        if len(affiliation_pieces_justificatives) != 0:
            rs["result"].append("En cours, infos filiale / courtier attendus"),
            rs["anorm"].append(
                "Attente des détails scannés des avenants spécifiés pour vérification des pièces constitutives.")
        else:
            rs["result"].append("En cours, infos filiale / courtier attendus"),
            rs["anorm"].append(
                "Attente des détails scannés des avenants spécifiés pour vérification des pièces constitutives.")
            

        if len(affiliation_facuration) != 0:
            rs["result"].append("En cours, infos filiale / courtier attendus"),
            rs["anorm"].append(
                "Absence de détail de facturation relatif aux bénéficiaires des polices.")
        else:
            rs["result"].append("En cours, infos filiale / courtier attendus"),
            rs["anorm"].append(
                "Absence de détail de facturation relatif aux bénéficiaires des polices.")
            

        if len(recouvrement) != 0:
            rs["result"].append("En cours, infos filiale / courtier attendus"),
            rs["anorm"].append(
                "Absence de l'état de recouvrement des factures de la période.")
        else:
            rs["result"].append("En cours, infos filiale / courtier attendus"),
            rs["anorm"].append(
                "Absence de l'état de recouvrement des factures de la période.")
            

        if len(facturation_frais_gestion) != 0:
            rs["result"].append("En cours, infos filiale / courtier attendus"),
            rs["anorm"].append(
                "Absence des états de facturation des honoraires de gestion")
        else:
            rs["result"].append("En cours, infos filiale / courtier attendus"),
            rs["anorm"].append(
                "Absence des états de facturation des honoraires de gestion")

        return rs


class Prestation:
    def __init__(self):
        return

    def database(self, file_benef, file_cotisation, file_conso):
        self.file_benef = file_benef
        self.file_cotisation = file_cotisation
        self.file_conso = file_conso
        benef = pd.read_excel(file_benef)
        benef.columns = [to_camel_case(column) for column in benef.columns]

        cotisation = pd.read_excel(file_cotisation)
        cotisation.columns = [to_camel_case(column)
                              for column in cotisation.columns]

        consommation = pd.read_excel(file_conso)
        consommation.columns = [to_camel_case(
            column) for column in consommation.columns]

        # Nombre de police unique de prestations --ok
        police_presta = consommation.groupby(
            ["numPolice"])["matricule"].agg({"count"}).reset_index().shape[0]

        # Nombre de beneficiaire des prestations
        beneficiaire_presta = len(set(consommation["matricule"]))

        # Dépense totale de toutes les prestations --ok
        depense_presta = np.sum(consommation.groupby(["numPolice"])[
                                "montantPaye"].agg({"sum"}).reset_index()["sum"])

        # Nombre de ligne de consommations --ok
        enregistrements = consommation.shape[0]

        return {
            "infos":
            {
                "police_presta": police_presta,
                "beneficiaire_presta": beneficiaire_presta,
                "depense_presta": depense_presta,
                "enregistrements": enregistrements,
            },
            "db":
            {
                "benef": benef,
                "cotisation": cotisation,
                "consommation": consommation,
            }
        }

    # Vérifiez si toutes les polices conso ont été facturées dans cotisation
    def conso_autorisees(self, file_benef, file_cotisation, file_conso):
        rs = ""
            
        benef = self.database(file_benef, file_cotisation,
                              file_conso)["db"]["benef"]
        cotisation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["cotisation"]
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"]
        df1 = benef["numPolice"].dropna().apply(lambda x: str(
            x).replace(" ", "")).drop_duplicates().tolist()
        df2 = cotisation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()
        df3 = consommation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()

        # 1. Les ID se trouvant dans df3 mais pas dans df2
        result_1 = [id for id in df1 if id not in df2 and id in df3]
        # 3. Les ID se trouvant dans df3 mais pas dans df1 et df2
        result_2 = [id for id in df3 if id not in df1 and id not in df2]

        if len(result_1) == 0 and len(result_2) ==0:
            rs =rs
        elif len(result_1) !=0 and len(result_2) == 0:
            rs = f"{list(set(result_1))} : bénéficiaire(s) sans cotisation ayant consommé(s)."
        elif len(result_1) ==0 and len(result_2) != 0:
            rs = f"{list(set(result_2))} : prestation(s) sans bénéficiaire(s) ni cotisations( cas grave)."
        elif len(result_1) !=0 and len(result_2) != 0:
            rs = f"{list(set(result_1))} : bénéficiaire(s) sans cotisation ayant consommé(s), {list(set(result_2))} : prestation(s) sans bénéficiaire(s) ni cotisations( cas grave)."
        return {
            "conso_auto": rs
        }

    # Verifier que la date émission cotisation <= date fin benef sinon relever les consommation associées
    def periode_couverture(self, file_benef, file_cotisation, file_conso, date_fin):
        rs = ""
        cotisation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["cotisation"].reset_index(drop=True)
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"].reset_index(drop=True)
        cotisation_filtre = cotisation[cotisation['dateEmission'] >= pd.to_datetime(
            date_fin, format='%d/%m/%Y', dayfirst=True)]["numPolice"].tolist()
        filtrage = consommation[consommation['numPolice'].isin(cotisation_filtre)]
        
        if len(list(filtrage)) == 0:
            rs= rs
        elif len(list(set(filtrage['numPolice']))) != 0:
            rs = f"{list(set(filtrage['numPolice']))} : prestations à une date d'émission après la période de contrôles."

        return {
            "periode_couv": rs 
        }

    # Verifiez que les enfants age <= 25 et les adultes <= 60 dans benef sinon relevez les conso associées
    def condition_age(self, file_benef, file_cotisation, file_conso, date_fin):
        rs = ""
        def age(x):
            if isinstance(x, str):
                return -5555555
            else:
                today = dt.datetime.strptime(str(date_fin), "%d/%m/%Y").date()
                return today.year - x.year - ((today.month, today.day) < (x.month, x.day))

        benef = self.database(file_benef, file_cotisation,
                              file_conso)["db"]["benef"]
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"]

        benef_sans_doublon = benef.drop_duplicates()
        benef_sans_doublon["dateNaissance"] = pd.to_datetime(
            benef_sans_doublon['dateNaissance'], format='infer', errors='coerce')
        benef_sans_doublon["age"] = benef_sans_doublon['dateNaissance'].apply(
            age)

        sans_doublons_conso = consommation.drop_duplicates()
        ENFANT_SUP_25 = benef_sans_doublon[benef_sans_doublon["statutAce"].isin(
            ['E']) & (benef_sans_doublon['age'] > 25)]
        ADULTES_SUP_60 = benef_sans_doublon[benef_sans_doublon["statutAce"].ne(
            'E') & (benef_sans_doublon['age'] > 60)]

        cond1 = sans_doublons_conso['matricule'].isin(
            ENFANT_SUP_25["matricule"].unique())
        ENFANT_SUP_25_CONSO = sans_doublons_conso[cond1]["numPolice"].apply(
            lambda x: str(x)).tolist()

        cond2 = sans_doublons_conso['matricule'].isin(
            ADULTES_SUP_60["matricule"].unique())
        ADULTES_SUP_60_CONSO = sans_doublons_conso[cond2]["numPolice"].apply(
            lambda x: str(x)).tolist()

        if len(ENFANT_SUP_25_CONSO) == 0 and len(ADULTES_SUP_60_CONSO) == 0:
            rs = rs
        elif len(ENFANT_SUP_25_CONSO) !=0 and len(ADULTES_SUP_60_CONSO) ==0:
            rs = f"{list(set(ENFANT_SUP_25_CONSO))} : prestations enfants de plus de 25 ans."
        elif len(ENFANT_SUP_25_CONSO) == 0 and len(ADULTES_SUP_60_CONSO) != 0:
            rs = f"{list(set(ADULTES_SUP_60_CONSO))} : prestations adultes de plus de 60 ans."
        elif len(ENFANT_SUP_25_CONSO) !=0 and len(ADULTES_SUP_60_CONSO) != 0:
            rs = f"{list(set(ENFANT_SUP_25_CONSO))} : prestations enfants de plus de 25 ans, {list(set(ADULTES_SUP_60_CONSO))} : prestations adultes de plus de 60 ans."
            
        return {
            "cndt_age": rs
        }

      # Vérifiez que les lignes consommation sont unique sinon relevez les conso associées
    def double_conso(self, file_benef, file_cotisation, file_conso):
        rs = ""
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"]
       
        colonnes_doublons = ['numPolice', 'matricule', 'dateSoins', 'codeActe', 'codeAffection', 'mtReclame', 'baseRemboursement', 'montantPaye']
        dble_conso = consommation[consommation.duplicated(subset=colonnes_doublons,keep=False)]

        liste_db_conso = consommation[consommation.duplicated(subset=colonnes_doublons,keep=False)]['numPolice']
        if len(list(dble_conso)) == 0:
            rs=rs
        elif len(list(dble_conso)) != 0:
            rs = f"{list(set(liste_db_conso))} : police(s) avec double prestations."
        return {
            "db_conso": rs
        }

    # Verifiez que date soins consommation - date emission cotisation <= 3 mois sinon relevez les conso associées
    def delai_contractuel(self, file_benef, file_cotisation, file_conso):
        rs = ""
        cotisation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["cotisation"]
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"]
        # Convertissez les colonnes de dates en objets datetime si elles ne le sont pas déjà
        cotisation['dateEmission'] = pd.to_datetime(cotisation['dateEmission'])
        consommation['dateSoins'] = pd.to_datetime(consommation['dateSoins'])

        # Calculez la différence en mois
        consommation['differenceEnMois'] = (
            abs(cotisation['dateEmission'] - consommation['dateSoins']).dt.days / 30)
        
        sansDateEmission = cotisation[cotisation['dateEmission'].isna()]["numPolice"].tolist()
        liste_ = set(list(consommation[consommation["numPolice"].isin(sansDateEmission)]["numPolice"]))
        # Filtrez les lignes où la différence est supérieure à 3 mois
        liste = set(
            list(consommation[consommation["differenceEnMois"] > 3]["numPolice"]))
        
        if len(list(liste)) == 0 and len(list(liste_)) ==0:
            rs=rs
        elif len(list(liste)) !=0 and len(list(liste_)) ==0:
            rs = f"{list(liste)} : prestations avec date de soins après plus de tois (03) mois de la date d'émission."
        elif len(list(liste)) ==0 and len(list(liste_)) !=0:
            rs = f"{list(liste_)} : prestations sans date émission."
        elif len(list(liste)) != 0 and len(list(liste_)) !=0:
            rs = f"{list(liste)} : prestations avec date de soins après plus de tois (03) mois de la date d'émission , {list(liste_)} : prestations sans date émission"
        

        return {
            "delai_cnt": rs
        }

    # Choisir quelques consommation élévées
    def depense_imortantes(self, file_benef, file_cotisation, file_conso, conso_ctx):
        rs=""
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"].dropna()
        consommation_impt = consommation[consommation["montantPaye"]>= int(conso_ctx)]
        # Trier par ordre décroissant de la colonne "MT" et sélectionner les n premiers éléments
        conso_sorted = consommation_impt.sort_values(by='montantPaye', ascending=False).drop_duplicates(
            subset=['montantPaye', 'numPolice'], keep='first')

        # Obtenir les valeurs de la colonne "NUM_POLICE" des n premiers éléments
        n_premiers_NUM_POLICE = set(conso_sorted['numPolice'].tolist())
        if len(list(n_premiers_NUM_POLICE)) == 0:
            rs=rs
        elif len(list(n_premiers_NUM_POLICE)) !=0:
            rs = f"{list(n_premiers_NUM_POLICE)} : quelques prestations importantes."
        return {
            "dp_mpt": rs
        }

    def resultats(self, file_benef, file_cotisation, file_conso, date_fin, conso_ctx=1000000):
        conso_autorisees = self.conso_autorisees(
            file_benef, file_cotisation, file_conso)["conso_auto"]
        periode_couverture = self.periode_couverture(
            file_benef, file_cotisation, file_conso, date_fin)["periode_couv"]
        condition_age = self.condition_age(
            file_benef, file_cotisation, file_conso, date_fin)["cndt_age"]
        double_conso = self.double_conso(
            file_benef, file_cotisation, file_conso)["db_conso"]
        delai_contractuel = self.delai_contractuel(
            file_benef, file_cotisation, file_conso)["delai_cnt"]
        depense_imortantes = self.depense_imortantes(
            file_benef, file_cotisation, file_conso,conso_ctx)["dp_mpt"]

        rs = {"result": [], "anorm": []}

        if len(conso_autorisees) != 0:
            rs["result"].append("KO, anomalie constatée")
            rs["anorm"].append(conso_autorisees)
        elif len(conso_autorisees) == 0:
            rs["result"].append("OK")
            rs["anorm"].append("NEANT")

        if len(periode_couverture) != 0:
            rs["result"].append("KO, anomalie constatée")
            rs["anorm"].append(periode_couverture)
        elif len(periode_couverture) == 0:
            rs["result"].append("OK")
            rs["anorm"].append("NEANT")

        if len(condition_age) != 0:
            rs["result"].append("KO, anomalie constatée"),
            rs["anorm"].append(condition_age)
        elif len(condition_age) == 0:
            rs["result"].append("OK")
            rs["anorm"].append("NEANT")

        if len(double_conso) != 0:
            rs["result"].append("KO, anomalie constatée"),
            rs["anorm"].append(double_conso)
        elif len(double_conso) == 0:
            rs["result"].append("OK")
            rs["anorm"].append("NEANT")

        if len(delai_contractuel) != 0:
            rs["result"].append("KO, anomalie constatée"),
            rs["anorm"].append(delai_contractuel)
        elif len(delai_contractuel) == 0:
            rs["result"].append("OK")
            rs["anorm"].append("NEANT")

        if len(depense_imortantes) != 0:
            rs["result"].append("KO, anomalie constatée"),
            rs["anorm"].append(depense_imortantes)
        elif len(depense_imortantes) == 0:
            rs["result"].append("OK")
            rs["anorm"].append("NEANT")

        return rs


class Resume:
    def __init__(self):
        return
    
    def database(self, file_benef, file_cotisation, file_conso):
        self.file_benef = file_benef
        self.file_cotisation = file_cotisation
        self.file_conso = file_conso

        try:
            benef = pd.read_excel(file_benef)
            benef.columns = [to_camel_case(column) for column in benef.columns]

            cotisation = pd.read_excel(file_cotisation)
            cotisation.columns = [to_camel_case(column)
                                for column in cotisation.columns]

            consommation = pd.read_excel(file_conso)
            consommation.columns = [to_camel_case(
                column) for column in consommation.columns]
        except:
            pass
        return {
            "db":
                {
                    "benef": benef,
                    "cotisation": cotisation,
                    "consommation": consommation,
            }}
        
    def ecart(self, file_benef, file_cotisation, file_conso, courtier,date_fin):
        try:
            nom_dossier = str(courtier)
            sous_dossier = "RESULTATS"
            dossier_complet = os.path.join(nom_dossier, sous_dossier)
            
            benef = self.database(file_benef, file_cotisation, file_conso)["db"]["benef"].reset_index(drop=True)
            cotisation = self.database(file_benef, file_cotisation, file_conso)["db"]["cotisation"].reset_index(drop=True)
            consommation = self.database(file_benef, file_cotisation, file_conso)["db"]["consommation"].reset_index(drop=True)

            # ... (le reste de votre code)

            # Votre code pour le traitement des données et la création des fichiers Excel

        except Exception as e:
            erreur_message = f"Une erreur s'est produite : {str(e)}"
            logging.error(erreur_message)
        
        benef = self.database(file_benef, file_cotisation, file_conso)["db"]["benef"].reset_index(drop=True)
        cotisation = self.database(file_benef, file_cotisation, file_conso)["db"]["cotisation"].reset_index(drop=True)
        consommation = self.database(file_benef, file_cotisation, file_conso)["db"]["consommation"].reset_index(drop=True)
        
        #1 - PRESTATIONS SANS COTISATION
        df1 = benef["numPolice"].dropna().apply(lambda x: str(
            x).replace(" ", "")).drop_duplicates().tolist()
        df2 = cotisation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()
        df3 = consommation["numPolice"].dropna().apply(
            lambda x: str(x).replace(" ", "")).drop_duplicates().tolist()
        presta_sans_cotisation = [id for id in df1 if id in df3 and id not in df2]
       
        #2 - PRESTATIONS HORS PERIODE COUVERTURE
        presta_hors_periode_couv = cotisation[cotisation['dateEmission'] >= pd.to_datetime(
            date_fin, format='%d/%m/%Y', dayfirst=True)]["numPolice"].tolist()
       
        #PRESTATIONS HORS CONDTION D'AGE()
        def age(x):
            if isinstance(x, str):
                return -5555555
            else:
                today = dt.datetime.strptime(str(date_fin), "%d/%m/%Y").date()
                return today.year - x.year - ((today.month, today.day) < (x.month, x.day))
        
        benef = self.database(file_benef, file_cotisation,
                            file_conso)["db"]["benef"]
        consommation = self.database(file_benef, file_cotisation, file_conso)[
            "db"]["consommation"]

        benef_sans_doublon = benef.drop_duplicates()
        benef_sans_doublon["dateNaissance"] = pd.to_datetime(
            benef_sans_doublon['dateNaissance'], format='infer', errors='coerce')
        benef_sans_doublon["age"] = benef_sans_doublon['dateNaissance'].apply(
            age)
        
         #3 CONSO BENEF SANS DATE DE NAISSANCE
        benef_sans_date_naiss = benef_sans_doublon[benef_sans_doublon['dateNaissance'].isna()]["numPolice"].apply(lambda x: str(x)).tolist()

        sans_doublons_conso = consommation.drop_duplicates()
        ENFANT_SUP_25 = benef_sans_doublon[benef_sans_doublon["statutAce"].isin(['E']) & (benef_sans_doublon['age'] > 25)]["numPolice"]
        ADULTES_SUP_60 = benef_sans_doublon[benef_sans_doublon["statutAce"].ne('E') & (benef_sans_doublon['age'] > 60)]["numPolice"]

        #4 CONSO ENFANTS DE PLUS DE 25 ANS
        cond_enf_sup_25 = sans_doublons_conso['numPolice'].isin(ENFANT_SUP_25)
        ENFANT_SUP_25_CONSO = sans_doublons_conso[cond_enf_sup_25]["numPolice"].apply(lambda x: str(x)).tolist()
       
        #5 CONSO ADULTES DE PLUS DE 60 ANS
        cond_adulte_sup_60 = sans_doublons_conso['numPolice'].isin(ADULTES_SUP_60)
        ADULTES_SUP_60_CONSO = sans_doublons_conso[cond_adulte_sup_60]["numPolice"].apply(lambda x: str(x)).tolist()
        
        #6 - DOUBLE CONSOMMATION - 
        #----numPolice, numMatricule, dateSoins, acte, codeAffectation, montantReclame, baseRemboursement, montantPaye
        colonnes_doublons = ['numPolice', 'matricule', 'dateSoins', 'codeActe', 'codeAffection', 'mtReclame', 'baseRemboursement', 'montantPaye']
        doublons_conso = consommation[consommation.duplicated(subset=colonnes_doublons,keep=False)]
        
        #------Formatage date emission & date soins
        cotisation['dateEmission'] = pd.to_datetime(cotisation['dateEmission'])
        consommation['dateSoins'] = pd.to_datetime(consommation['dateSoins'])

        #7 - PRESTATIONS SANS DATE EMISSION
        cotisation['dateEmission'] = pd.to_datetime(cotisation['dateEmission'])
        consommation['dateSoins'] = pd.to_datetime(consommation['dateSoins'])
        sansDateEmission = cotisation[cotisation['dateEmission'].isna()]["numPolice"].tolist()
        liste_sans_date_emission = set(list(consommation[consommation["numPolice"].isin(sansDateEmission)]["numPolice"]))

        #8 - PRESTATION HORS DELAI CONTRACTUEL
        consommation['differenceEnMois'] = (
            abs(cotisation['dateEmission'] - consommation['dateSoins']).dt.days / 30)
        liste_hors_delai_contractuel = set(
            list(consommation[consommation["differenceEnMois"] > 3]["numPolice"]))
        
        ##-----------
        # Convert pandas dataframes to dask dataframes
        consommation_dask = dd.from_pandas(consommation, npartitions=10)
        benef_dask = dd.from_pandas(benef, npartitions=10)
        

        #9 - PRESTA AVEC DATE SOINS HORS PERIODE: COTISATION x CONSO
        conso_x_cotisation = pd.merge(consommation, cotisation, on='numPolice', how='inner')
        conso_x_cotisation = conso_x_cotisation[['numPolice','dateSoins','dateDebutExercice', 'dateFinExercice']].dropna() 
        liste_presta_hors_periode = list(set(conso_x_cotisation[(conso_x_cotisation['dateSoins'].lt(conso_x_cotisation['dateDebutExercice'])) | (conso_x_cotisation['dateSoins'].gt(conso_x_cotisation['dateFinExercice']))]['numPolice'].apply(lambda x: str(x).replace(" ",""))))
        
        #10 - PRESTA AVEC DATE SOINS APRES DATE SORTIE: BENF x CONSO
        try:
            presta_x_benef = dd.merge(consommation_dask, benef_dask, on='numPolice', how='inner').compute()
            presta_x_benef = presta_x_benef[['numPolice','dateSoins','dateSortie']].dropna() 
            liste_presta_dateSoins_apres_dateSortie = list(set(presta_x_benef[(presta_x_benef['dateSoins'].gt(presta_x_benef['dateSortie']))]['numPolice'].apply(lambda x: str(x).replace(" ",""))))
        except:
            pass


        #11 - ERREUR DATE NAISSANCE: 
        ERREUR_AGE = benef_sans_doublon[(benef_sans_doublon["statutAce"] == "A") & ((benef_sans_doublon['age'] <= 25) | (benef_sans_doublon['age'] < 0)) | ((benef_sans_doublon["statutAce"] == 'E') & (benef_sans_doublon['age'] < 0))]
        cond = sans_doublons_conso['matricule'].isin(ERREUR_AGE["matricule"].unique())
        liste_errs= list(set(sans_doublons_conso[cond]["numPolice"].tolist()))

        df_erreur_age = sans_doublons_conso[sans_doublons_conso["numPolice"].isin(liste_errs)]
        df_erreur_age.loc[:, "age"] = pd.to_datetime(benef_sans_doublon['dateNaissance'], dayfirst=True).apply(age)
       
        #----------------FICHIER ECARTS 
        with pd.ExcelWriter('CONTROLE GESTION DELEGUE'+'/'+dossier_complet+'/'+'ECARTS.xlsx', engine='xlsxwriter') as writer:
            try:
                if len(presta_sans_cotisation) != 0:
                    df = consommation[consommation["numPolice"].isin(presta_sans_cotisation)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="PRESTA_SANS_COTISATION", index=False)
            except:
                #print('PRESTA_SANS_COTISATION vide')
                pass
            try:
                if len(presta_hors_periode_couv) != 0:
                    df = consommation[consommation["numPolice"].isin(presta_hors_periode_couv)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="PRESTA_HORS_PERIODE_CVT", index=False)
            except:
                #print('PRESTA_HORS_PERIODE_CVT vide')
                pass
            try:
                if len(ENFANT_SUP_25_CONSO) != 0:
                    df = consommation[consommation["numPolice"].isin(ENFANT_SUP_25_CONSO)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="PRESTA_ENFANTS_SUP_25_ANS", index=False)
            except:
                #print('PRESTA_ENFANTS_SUP_25_ANS vide')
                pass
            try:
                if len(ADULTES_SUP_60_CONSO) != 0:
                    df = consommation[consommation["numPolice"].isin(ADULTES_SUP_60_CONSO)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="PRESTA_ADULTES_SUP_60_ANS", index=False)
            except:
                #print('PRESTA_ADULTES_SUP_60_ANS vide')
                pass
            try:
                if len(benef_sans_date_naiss) != 0:
                    df = consommation[consommation["numPolice"].isin(benef_sans_date_naiss)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="SANS_DATE_NAISS", index=False)
            except:
                #print('PRESTA_SANS_DATE_NAISS vide')
                pass
            try:
                if len(doublons_conso) != 0:
                    df = consommation[consommation.duplicated(subset=colonnes_doublons,keep=False)].drop(columns=['unnamed:0'],axis=1)
                    df.to_excel(writer, sheet_name="DOUBLE_CONSOMMATION", index=False)
            except:
                #print('DOUBLE_CONSOMMATION vide')
                pass
            try:
                if len(liste_sans_date_emission) != 0:
                    df = consommation[consommation["numPolice"].isin(liste_sans_date_emission)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="SANS_DATE_EMISSION", index=False)
            except:
                #print('SANS_DATE_EMISSION vide')
                pass
            try:
                if len(liste_hors_delai_contractuel) !=0:
                    df = consommation[consommation["numPolice"].isin(liste_hors_delai_contractuel)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="HORS_DELAI_CNTCT", index=False)
            except:
                #print('PRESTA_HORS_DELAI_CONTRACTUEL vide')
                pass
            try:
                if len(liste_presta_hors_periode) != 0:
                    df = consommation[consommation["numPolice"].isin(liste_presta_hors_periode)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="HORS_PERIODE_CTRL", index=False)
            except:
                #print('PRESTA_HORS_PERIODE vide')
                pass
            try:
                if len(liste_presta_dateSoins_apres_dateSortie) != 0:
                    df = consommation[consommation["numPolice"].isin(liste_presta_dateSoins_apres_dateSortie)].drop(columns=['unnamed:0','differenceEnMois'],axis=1)
                    df.to_excel(writer, sheet_name="DSOINS_APR_DSORTIE", index=False)
            except:
                #print('DATE_SORTIE_APRES_DATE_SOINS vide')
                pass
            try:
                if len(liste_errs) != 0:
                    #df.loc[:, "age"] = pd.to_datetime(benef_sans_doublon['dateNaissance'], dayfirst=True).apply(age)
                    df = consommation[consommation["numPolice"].isin(liste_errs)]
                    df.to_excel(writer, sheet_name="ERREUR_DATE_NAISSANCE ", index=False)
                    
            except:
                #print('ERREUR_DATE_NAISSANCE vide')
                pass
            


        #-----------FICHIER DETAILS ECARTS
        with pd.ExcelWriter('CONTROLE GESTION DELEGUE'+'/'+dossier_complet + '/' + 'DETAILS_ECARTS.xlsx', engine='xlsxwriter') as writer_:
            try:
                # Si len(presta_sans_cotisation) n'est pas nul, procédez à la création de la feuille Excel
                if len(presta_sans_cotisation) != 0:
                    df = consommation[consommation["numPolice"].isin(presta_sans_cotisation)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="PRESTA_SANS_COTISATION", index=False)

            except Exception as e:
                #print('PRESTA_SANS_COTISATION vide')
                pass
            try:
                if len(presta_hors_periode_couv) != 0:
                    df = consommation[consommation["numPolice"].isin(presta_hors_periode_couv)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="PRESTA_HORS_PERIODE_CVT", index=False)                   
            except:
                #print('PRESTA_HORS_PERIODE_CVT vide')
                pass
            try:
                if len(ENFANT_SUP_25_CONSO) != 0:
                    df = consommation[consommation["numPolice"].isin(ENFANT_SUP_25_CONSO)].groupby(['numPolice']).agg({'mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="PRESTA_ENFANTS_SUP_25_ANS", index=False)
            except:
                #print('PRESTA_ENFANTS_SUP_25_ANS vide')
                pass
            try:
                if len(ADULTES_SUP_60_CONSO) != 0:
                    df = consommation[consommation["numPolice"].isin(ADULTES_SUP_60_CONSO)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="PRESTA_ADULTES_SUP_60_ANS", index=False)
                    
            except:
                #print('PRESTA_ADULTES_SUP_60_ANS vide')
                pass
            try:
                if len(benef_sans_date_naiss) != 0:
                    df = consommation[consommation["numPolice"].isin(benef_sans_date_naiss)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="SANS_DATE_NAISS", index=False)
            except:
                #print('PRESTA_SANS_DATE_NAISS vide')
                pass
            try:
                if len(doublons_conso) != 0:
                    df = doublons_conso.groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="DOUBLE_CONSOMMATION", index=False)
            except:
                #print('DOUBLE_CONSOMMATION vide')
                pass
            try:
                if len(liste_sans_date_emission) != 0:
                    df = consommation[consommation["numPolice"].isin(liste_sans_date_emission)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="SANS_DATE_EMISSION", index=False)
            except:
                #print('SANS_DATE_EMISSION vide')
                pass
            try:
                if len(liste_hors_delai_contractuel) !=0:
                    df = consommation[consommation["numPolice"].isin(liste_hors_delai_contractuel)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="HORS_DELAI_CNTCT", index=False)
            except:
                #print('PRESTA_HORS_DELAI_CONTRACTUEL vide')
                pass
            try:
                if len(liste_presta_hors_periode) != 0:
                    
                    df = consommation[consommation["numPolice"].apply(lambda x: str(x).replace(" ","")).isin(liste_presta_hors_periode)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="HORS_PERIODE_CTRL", index=False)
            except:
                #print('PRESTA_HORS_PERIODE vide')
                pass
            try:
                if len(liste_presta_dateSoins_apres_dateSortie) != 0:
                    df = consommation[consommation["numPolice"].isin(liste_presta_dateSoins_apres_dateSortie)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="DSOINS_APR_DSORTIE", index=False)
            except:
                #print('DATE_SORTIE_APRES_DATE_SOINS vide')
                pass
            try:
                if len(liste_errs) != 0:
                    df = consommation[consommation["numPolice"].isin(liste_errs)].groupby(['numPolice']).agg({'matricule':'count','mtReclame': 'sum', 'mtExclu': 'sum','baseRemboursement':'sum', 'ticketModerateur': 'sum', 'montantPaye': 'sum'})
                    totals = df.sum(numeric_only=True) 
                    df.loc['TOTAL'] = totals
                    df.reset_index(inplace=True)
                    df.reset_index().to_excel(writer_, sheet_name="ERREUR_DATE_NAISSANCE", index=False)
                    
            except:
                #print('AGE_ERRONES vide')
                pass
       
           
            
     
            