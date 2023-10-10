import locale
locale.setlocale(locale.LC_ALL, 'fr_CI.UTF-8')
from imports import *
from functions import *

base_directory = 'CONTROLE GESTION DELEGUE'
if not os.path.exists(base_directory):
    os.makedirs(base_directory, exist_ok=True)

# Global variables to store the file paths
selected_beneficiaire_path = ""
selected_emission_path = ""
selected_consommation_path = ""

def upload_beneficiaire_file():
    global selected_beneficiaire_path
    file_path = filedialog.askopenfilename()
    if file_path:
        selected_beneficiaire_path = file_path
        beneficiaire_entry.delete(0, tk.END)
        beneficiaire_entry.insert(0, selected_beneficiaire_path)


def upload_emission_file():
    global selected_emission_path
    file_path = filedialog.askopenfilename()
    if file_path:
        selected_emission_path = file_path
        emission_entry.delete(0, tk.END)
        emission_entry.insert(0, selected_emission_path)


def upload_consommation_file():
    global selected_consommation_path
    file_path = filedialog.askopenfilename()
    if file_path:
        selected_consommation_path = file_path
        consommation_entry.delete(0, tk.END)
        consommation_entry.insert(0, selected_consommation_path)


def age(x):
    if isinstance(x, str):
        return -5555555
    else:
        today = dt.datetime.strptime(
            str(end_date_entry.get()), "%d/%m/%Y").date()
        return today.year - x.year - ((today.month, today.day) < (x.month, x.day))


def toggle_date_fields_benef():
    if benef_checkbox_var.get() == 1:
        anterieure_entry_benef_sortie.config(state="normal")
        posterieure_entry_benef_sortie.config(state="normal")
    else:
        anterieure_entry_benef_sortie.config(state="disabled")
        posterieure_entry_benef_sortie.config(state="disabled")


def toggle_date_fields_emission():
    if emission_checkbox_var.get() == 1:
        anterieure_date_emission_entry.config(state="normal")
        posterieure_date_emission_entry.config(state="normal")
    else:
        anterieure_date_emission_entry.config(state="disabled")
        posterieure_date_emission_entry.config(state="disabled")


def toggle_date_fields_conso():
    if conso_checkbox_var.get() == 1:
        anterieure_dt_soins_conso.config(state="normal")
        posterieure_dt_soins_conso.config(state="normal")
        anterieure_date_recept_conso.config(state="normal")
        posterieure_date_recept_conso.config(state="normal")

    else:
        anterieure_dt_soins_conso.config(state="disabled")
        posterieure_dt_soins_conso.config(state="disabled")
        anterieure_date_recept_conso.config(state="disabled")
        posterieure_date_recept_conso.config(state="disabled")


# Create the main window
window = tk.Tk()
window.title("*"*60 + "CONTROLE DELEGUE MCI-CARE COTE D'IVOIRE"+"*"*60)
window.resizable(False, False)
# Create a ttkbootstrap style
style = Style(theme="superhero")

# Create Label, Entry field, and Upload button for BENEFICIAIRE
beneficiaire_label = ttk.Label(window, text="BENEFICIAIRE:")
beneficiaire_label.grid(row=0, column=0, padx=10, pady=5)

beneficiaire_entry = ttk.Entry(window, width=50)
beneficiaire_entry.grid(row=0, column=1, padx=10, pady=5)

beneficiaire_upload_button = ttk.Button(
    window, text="Chargez fichier BENEFICIAIRE", command=upload_beneficiaire_file)
beneficiaire_upload_button.grid(row=0, column=2, padx=10, pady=10)

# Create Label, Entry field, and Upload button for EMISSION
emission_label = ttk.Label(window, text="EMISSION:")
emission_label.grid(row=1, column=0, padx=10, pady=5)

emission_entry = ttk.Entry(window, width=50)
emission_entry.grid(row=1, column=1, padx=10, pady=5)

emission_upload_button = ttk.Button(
    window, text="Chargez fichier EMISSION", command=upload_emission_file)
emission_upload_button.grid(row=1, column=2, padx=10, pady=10)

# Create Label, Entry field, and Upload button for CONSOMMATION
consommation_label = ttk.Label(window, text="CONSOMMATION:")
consommation_label.grid(row=2, column=0, padx=10, pady=5)

consommation_entry = ttk.Entry(window, width=50)
consommation_entry.grid(row=2, column=1, padx=10, pady=5)

consommation_upload_button = ttk.Button(
    window, text="Chargez fichier CONSOMMATION", command=upload_consommation_file)
consommation_upload_button.grid(row=2, column=2, padx=10, pady=10)

# Label and Entry for Start Date
start_label = ttk.Label(window, text="Date début:")
start_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
start_date_entry = ttk.Entry(window)
start_date_entry.grid(row=3, column=1, padx=10, pady=10, sticky="w")

# Label and Entry for End Date
end_label = ttk.Label(window, text="Date de fin:")
end_label.grid(row=3, column=2, padx=10, pady=10, sticky="w")
end_date_entry = ttk.Entry(window)
end_date_entry.grid(row=3, column=3, padx=10, pady=10, sticky="w")

# Label and Entry for Seuil coûteux
seuil_label = ttk.Label(window, text="Seuil coûteux:")
seuil_label.grid(row=4, column=0, padx=10, pady=10, sticky="w")
seuil_entry = ttk.Entry(window)
seuil_entry.grid(row=4, column=1, padx=10, pady=10, sticky="w")

# Label and Entry for COUTIER
COUTIER_label = ttk.Label(window, text="LIBELLE COURTIER - NOM PAYS:")
COUTIER_label.grid(row=4, column=2, padx=10, pady=10, sticky="w")
COUTIER_entry = ttk.Entry(window)
COUTIER_entry.grid(row=4, column=3, padx=10, pady=10, sticky="w")


# Début

# Add a label for RESTRICTIONS
restrictions_label = ttk.Label(
    window, text="RESTRICTIONS", font=("Helvetica", 12, "bold"))
restrictions_label.grid(row=5, column=0, columnspan=3,
                        padx=10, pady=10, sticky="w")

# Create IntVar variables for the checkboxes
benef_checkbox_var = tk.IntVar()
emission_checkbox_var = tk.IntVar()
conso_checkbox_var = tk.IntVar()


# Create checkboxes for BENEF, EMISSION, and CONSO
benef_checkbox = ttk.Checkbutton(
    window, text="BENEF.", variable=benef_checkbox_var, command=toggle_date_fields_benef)
benef_checkbox.grid(row=6, column=0, padx=10, pady=5, sticky="w")

emission_checkbox = ttk.Checkbutton(
    window, text="EMISSION.", variable=emission_checkbox_var, command=toggle_date_fields_emission)
emission_checkbox.grid(row=6, column=1, padx=10, pady=5, sticky="w")

conso_checkbox = ttk.Checkbutton(
    window, text="CONSO.", variable=conso_checkbox_var, command=toggle_date_fields_conso)
conso_checkbox.grid(row=6, column=2, padx=10, pady=5, sticky="w")


# Créez une étiquette pour afficher le texte "BENEFICIAIRE"
benef_label = Label(window, text="BENEFICIAIRE", font=("Arial", 10, "bold"))
benef_label.grid(row=7, column=2, columnspan=3, padx=10, pady=10, sticky="w")


date_sortie_label_benf = ttk.Label(window, text="Date sortie ")
anterieure_entry_benef_sortie = ttk.Entry(window, state="disabled")
posterieure_entry_benef_sortie = ttk.Entry(window, state="disabled")
anterieure_entry_benef_sortie.grid(
    row=8, column=2, padx=10, pady=5, sticky="w")
posterieure_entry_benef_sortie.grid(
    row=8, column=4, padx=10, pady=5, sticky="w")
antérieur_entry_date_sortie = ttk.Label(window, text="Antérieur à :")  # Début
antérieur_entry_date_sortie.grid(row=8, column=1, padx=10, pady=5, sticky="e")
posterieur_entry_date_sortie = ttk.Label(window, text="Postérieur à :")
posterieur_entry_date_sortie.grid(
    row=8, column=3, padx=10, pady=5, sticky="e")  # Fin
date_sortie_label_benf.grid(row=8, column=0, padx=10, pady=5, sticky="w")


# Créez une étiquette pour afficher le texte "BENEFICIAIRE"
emission_label = Label(window, text="EMISSION", font=("Arial", 10, "bold"))
emission_label.grid(row=11, column=2, columnspan=3,
                    padx=10, pady=10, sticky="w")


date_emission_label = ttk.Label(window, text="Date émission ")
anterieure_date_emission_entry = ttk.Entry(window, state="disabled")
posterieure_date_emission_entry = ttk.Entry(window, state="disabled")
date_emission_label.grid(row=12, column=0, padx=10, pady=5, sticky="w")
anterieure_date_emission_entry.grid(
    row=12, column=2, padx=10, pady=5, sticky="w")
posterieure_date_emission_entry.grid(
    row=12, column=4, padx=10, pady=5, sticky="w")
antérieur_date_emission = ttk.Label(window, text="Antérieur à :")  # Début
antérieur_date_emission .grid(row=12, column=1, padx=10, pady=5, sticky="e")
posterieur_date_emission = ttk.Label(window, text="Postérieur à :")
posterieur_date_emission .grid(
    row=12, column=3, padx=10, pady=5, sticky="e")  # Fin label

# Début
# Créez une étiquette pour afficher le texte "BENEFICIAIRE"
conso_label = Label(window, text="CONSO", font=("Arial", 10, "bold"))
conso_label.grid(row=15, column=2, columnspan=3, padx=10, pady=10, sticky="w")

# Initially disable the date fields
date_soins_label_conso = ttk.Label(window, text="Date soins ")
anterieure_dt_soins_conso = ttk.Entry(window, state="disabled")
posterieure_dt_soins_conso = ttk.Entry(window, state="disabled")
date_soins_label_conso.grid(row=16, column=0, padx=10, pady=5, sticky="w")
anterieure_dt_soins_conso.grid(row=16, column=2, padx=10, pady=5, sticky="w")
posterieure_dt_soins_conso.grid(row=16, column=4, padx=10, pady=5, sticky="w")
antérieur_date_soins = ttk.Label(window, text="Antérieur à :")  # Début
antérieur_date_soins.grid(row=16, column=1, padx=10, pady=5, sticky="e")
posterieur_date_soins = ttk.Label(window, text="Postérieur à :")
posterieur_date_soins.grid(row=16, column=3, padx=10,
                           pady=5, sticky="e")  # Fin label


date_recept_conso = ttk.Label(window, text="Date réception ")
anterieure_date_recept_conso = ttk.Entry(window, state="disabled")
posterieure_date_recept_conso = ttk.Entry(window, state="disabled")
date_recept_conso.grid(row=17, column=0, padx=10, pady=5, sticky="w")
anterieure_date_recept_conso.grid(
    row=17, column=2, padx=10, pady=5, sticky="w")
posterieure_date_recept_conso.grid(
    row=17, column=4, padx=10, pady=5, sticky="w")
antérieur_date_rgl = ttk.Label(window, text="Antérieur à :")  # Début
antérieur_date_rgl.grid(row=17, column=1, padx=10, pady=5, sticky="e")
posterieur_date_rgl = ttk.Label(window, text="Postérieur à :")
posterieur_date_rgl.grid(row=17, column=3, padx=10,
                         pady=5, sticky="e")  # Fin label


# ---------------- LES FONCTIONS al@ssanlwag0pu

        
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
    # matricule_beneficiaire
    if column_name.lower() in ["matricule_beneficiaire", "matricule beneficiaire"]:
        column_name = "MATRICULE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    if column_name.lower() in ["libelle_gestionnaire", "libelle gestionnaire"]:
        column_name = "GESTIONNAIRE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    if column_name.lower() in ["date reception", "date_reception", "Date Recepetion"]:
        column_name = "DATE_RECEPTION"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    if column_name.lower() in ["date sortie", "date_sortie"]:
        column_name = "DATE_SORTIE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])
    # Date Debut Exercice
    if column_name.lower() in ["Date Debut Exercice", "DATE DEBUT EXERCICE", "Date_Debut_Exercice", "DATE_DEBUT_EXERCICE"]:
        column_name = "DATE_DEBUT_EXERCICE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    if column_name.lower() in ["Date Fin Exercice", "DATE FIN EXERCICE", "Date_Fin_Exercice", "DATE_FIN_EXERCICE"]:
        column_name = "DATE_FIN_EXERCICE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    # ['numPolice', 'dateSoins', 'codeActe', 'codeAffectation', 'mtReclame', 'baseRemboursement', 'montantPaye']
    if column_name.lower() in ["CODE_ACTE", "code acte", "Code Acte", "code_acte", "Code_Acte"]:
        column_name = "CODE_ACTE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    if column_name.lower() in ["CODE AFFECTION", "code affection", "code_affection", "Code Affection", "Code_Affection"]:
        column_name = "CODE_AFFECTION"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    if column_name.lower() in ["MT RECLAME", "montant reclame", "montant_reclame", "Montant Reclame", "Montant_Reclame"]:
        column_name = "MT_RECLAME"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    if column_name.lower() in ["BASE REMBOURSEMENT", "base remboursement", "base_remboursement", "Base Remboursement", "Base_Remboursement"]:
        column_name = "BASE_REMBOURSEMENT"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    if column_name.lower() in ["MONTANT PAYE", "Montant Paye", "Montant_Paye", "montant paye", "montant_paye"]:
        column_name = "MONTANT_PAYE"
    words = column_name.replace(" ", "_").replace("_", " ").split()
    camel_case_name = words[0].lower() + ''.join(word.capitalize()
                                                 for word in words[1:])

    return camel_case_name



def critere_benef(file, ante=pd.to_datetime(str(start_date_entry.get()), dayfirst=True), post=pd.to_datetime(str(end_date_entry.get())), dayfirst=True):
    benef = pd.read_excel(file)
    benef.columns = [to_camel_case(column) for column in benef.columns]

    benef["dateSortie"] = pd.to_datetime(
        benef["dateSortie"], format='%d/%m/%Y', errors='coerce', dayfirst=True)
    benef["dateSortie"] = benef["dateSortie"].apply(
        lambda x: pd.NaT if x > post else x)
    benef = benef[(benef["dateSortie"].lt(pd.to_datetime(str(post), dayfirst=True))) & (
        benef["dateSortie"].gt(pd.to_datetime(str(ante), dayfirst=True))) | (benef["dateSortie"].isna())]
    benef_doublon = benef[benef.duplicated(keep=False)]
    benef_sans_doublon = benef.drop_duplicates()

    police_benef = benef_sans_doublon.groupby(
        ["numPolice"])["numPolice"].agg({"count"}).reset_index()
    benef_sans_doublon.loc[:, "age"] = pd.to_datetime(
        benef_sans_doublon['dateNaissance'], dayfirst=True).apply(age)

    benef_sans_date_naiss = benef_sans_doublon[benef_sans_doublon['dateNaissance'].isna(
    )]

    if benef_sans_doublon.shape[0] != 0:
        ENFANT_SUP_25 = benef_sans_doublon[benef_sans_doublon["statutAce"].isin(
            ['E']) & (benef_sans_doublon['age'] > 25)]
        ADULTES_SUP_60 = benef_sans_doublon[benef_sans_doublon["statutAce"].ne(
            'E') & (benef_sans_doublon['age'] > 60)]
    else:
        ENFANT_SUP_25 = benef_sans_doublon
        ADULTES_SUP_60 = benef_sans_doublon

    df_erreur_age = benef_sans_doublon[(benef_sans_doublon["statutAce"] == "A") & ((benef_sans_doublon['age'] <= 25) | (
        benef_sans_doublon['age'] < 0)) | ((benef_sans_doublon["statutAce"] == 'E') & (benef_sans_doublon['age'] < 0))]
    df_erreur_age.loc[:, "age"] = pd.to_datetime(
        benef_sans_doublon['dateNaissance'], dayfirst=True).apply(age)
    liste_errs = list(set(df_erreur_age["numPolice"]))

    if len(benef_sans_doublon) != 0:

        path_sans_doublon = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "SANS_DOUBLONS_BENEF.xlsx")
        benef_sans_doublon.to_excel(
            path_sans_doublon, sheet_name="SANS_DOUBLONS_BENEF")

    if len(benef_doublon) != 0:
        path_doublon = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "DOUBLONS_BENEF.xlsx")
        benef_doublon.to_excel(path_doublon, sheet_name="DOUBLONS_BENEF")

    if len(police_benef) != 0:
        path_police_benef = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "POLICE_BENEF.xlsx")
        police_benef.to_excel(path_police_benef, sheet_name="POLICE_BENEF")

    if len(ENFANT_SUP_25) != 0:
        path_ENFANT_SUP_25 = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "ENFANTS_SUP_25_ANS.xlsx")
        ENFANT_SUP_25.to_excel(
            path_ENFANT_SUP_25, sheet_name="ENFANTS_SUP_25_ANS")

    if len(ADULTES_SUP_60) != 0:
        path_ADULTES_SUP_60 = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "ADULTES_SUP_60_ANS.xlsx")
        ADULTES_SUP_60.to_excel(path_ADULTES_SUP_60,
                                sheet_name="ADULTES_SUP_60_ANS")

    if len(benef_sans_date_naiss) != 0:
        path_benef_sans_date_naiss = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "BENEF_SANS_DATE_NAISS.xlsx")
        benef_sans_date_naiss.to_excel(
            path_benef_sans_date_naiss, sheet_name="BENEF_SANS_DATE_NAISS")

    if len(df_erreur_age) != 0:
        path_df_erreur_age = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "BENEF_ERREUR_DATE_NAISS.xlsx")
        df_erreur_age[df_erreur_age["numPolice"].isin(liste_errs)].to_excel(
            path_df_erreur_age, sheet_name="BENEF_ERREUR_DATE_NAISS")

    # ------------Controle production
    Obj = Production().database(selected_beneficiaire_path,
                                selected_emission_path, selected_consommation_path)['infos']
    aff_avnt = Production().affiliation_avnants(selected_beneficiaire_path,
                                                selected_emission_path, selected_consommation_path)
    aff_pieces_justif = Production().affiliation_pieces_justificatives(
        selected_beneficiaire_path, selected_emission_path, selected_consommation_path, n_alea=3)
    aff_facuration = Production().affiliation_facuration(selected_beneficiaire_path,
                                                         selected_emission_path, selected_consommation_path, n_alea=4)
    recouvrement = Production().recouvrement(selected_beneficiaire_path,
                                             selected_emission_path, selected_consommation_path)
    facturation_frais_gestion = Production().facturation_frais_gestion(
        selected_beneficiaire_path, selected_emission_path, selected_consommation_path)
    rs = Production().resultats(selected_beneficiaire_path,
                                selected_emission_path, selected_consommation_path)

    # "Affiliations (pièces justificatives)","Affiliations (Facturation)"
    df_production = pd.DataFrame({
        "Volet": ["Production", "Affiliations (avenants)", "Affiliations (pièces justificatives)", "Affiliations (Facturation)", "Recouvrement", "Facturation frais de Gestion"],
        "Périmètre contrôlé": [
            f"{Obj['police_en_gestion']} polices en gestion, {Obj['police_facturee']} polices facturées comportant {Obj['piece_de_facturation']} pièces de facturation pour la période contrôlée. Contrôle de la production",
            f"{Obj['police_facturee']} polices facturées comportant {Obj['piece_de_facturation']} pièces de facturation pour la période contrôlée. Contrôle des avenants: {aff_avnt['polices']}",
            f"{Obj['police_facturee']} polices facturées comportant {Obj['piece_de_facturation']} pièces de facturation pour la période contrôlée. Contrôle des pièces justificatives des avenants: {aff_pieces_justif['pieces_justif']}",
            f"{Obj['police_en_gestion']} polices en gestion pour {Obj['piece_de_facturation']} pièces de facturation, {Obj['police_facturee']} polices avec un éffectif de total de {Obj['beneficiaire']} bénéficiaires, valeur facturée {Obj['valeur_facturee']} F CFA. Contrôle de la facturation: {aff_facuration['aff_facuration']}",
            f"{Obj['police_en_gestion']} polices en gestion pour {Obj['piece_de_facturation']} pièces de facturation, {Obj['police_facturee']} polices avec un éffectif de total de {Obj['beneficiaire']} bénéficiaires, valeur facturée {Obj['valeur_facturee']} F CFA. Contrôle du recouvrement: {recouvrement['recouv']}",
            f"{Obj['police_en_gestion']} polices en gestion pour {Obj['piece_de_facturation']} pièces de facturation, {Obj['police_facturee']} polices avec un éffectif de total de {Obj['beneficiaire']} bénéficiaires, valeur facturée {Obj['valeur_facturee']} F CFA. Contrôle du recouvrement: {facturation_frais_gestion['frais_gestion']}",
        ],
        "Résultat": rs["result"],
        "Anomalie": rs["anorm"]
    })
    if len(df_production) != 0:
        path_df_production = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/RESULTATS", "PRODUCTION.xlsx")
        df_production.to_excel(path_df_production, sheet_name="PRODUCTION")
    return


def critere_emission(file, post=pd.to_datetime(str(end_date_entry.get()), dayfirst=True)):
    emission = pd.read_excel(file)
    emission.columns = [to_camel_case(column) for column in emission.columns]

    doublons_emission = emission[emission.duplicated(keep=False)]
    sans_doublons_emission = emission.drop_duplicates()
    police_emission = sans_doublons_emission.groupby(
        ["numPolice"])["numPolice"].agg({"count"}).reset_index()

    if len(doublons_emission) != 0:
        path_doublon = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "DOUBLONS_COTISATION.xlsx")
        doublons_emission.to_excel(
            path_doublon, sheet_name="DOUBLONS_COTISATION")

    if len(sans_doublons_emission) != 0:
        path_sans_doublon = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "SANS_DOUBLONS_COTISATION.xlsx")
        sans_doublons_emission.to_excel(
            path_sans_doublon, sheet_name="SANS_DOUBLONS_COTISATION")

    if len(police_emission) != 0:
        path_police_emission = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/PRODUCTION", "POLICE_COTISATION.xlsx")
        police_emission.to_excel(path_police_emission,
                                 sheet_name="POLICE_COTISATION")
    return


def critere_conso(file):
    conso = pd.read_excel(file)
    conso.columns = [to_camel_case(column) for column in conso.columns]

    colonnes_doublons = ['numPolice', 'matricule', 'dateSoins', 'codeActe',
                         'codeAffection', 'mtReclame', 'baseRemboursement', 'montantPaye']
    doublons_conso = conso[conso.duplicated(
        subset=colonnes_doublons, keep=False)]
    sans_doublons_conso = conso.drop_duplicates()
    police_conso = conso.groupby(['numPolice'])['montantPaye'].agg(
        {"count", "sum"}).sort_values(by=["count"], ascending=False).reset_index()

    if len(doublons_conso) != 0:
        path_doublon = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/CONSOMMATION", "DOUBLONS_CONSO.xlsx")
        doublons_conso.to_excel(path_doublon, sheet_name="DOUBLONS_CONSO")

    if len(sans_doublons_conso) != 0:
        path_sans_doublons_conso = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/CONSOMMATION", "SANS_DOUBLONS_CONSO.xlsx")
        sans_doublons_conso.to_excel(
            path_sans_doublons_conso, sheet_name="SANS_DOUBLONS_CONSO")

    if len(police_conso) != 0:
        path_police_conso = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/CONSOMMATION", "POLICE_CONSO.xlsx")
        police_conso.to_excel(path_police_conso, sheet_name="POLICE_CONSO")

    # ------------Controle production
    Obj = Prestation().database(selected_beneficiaire_path,
                                selected_emission_path, selected_consommation_path)['infos']
    rs = Prestation().resultats(selected_beneficiaire_path, selected_emission_path,
                                selected_consommation_path, date_fin=end_date_entry.get(), conso_ctx=int(seuil_entry.get()))
    # print(len( rs["result"]), len( rs["anorm"]))
    if (len(rs["result"]) == len(rs["anorm"]) == 6):
        df_prestation = pd.DataFrame({
            "Volet": ["Consommations autorisées", "Période de couverture", "Conditions d'âge", "Double consommations", "Délai contractuel", "Dépenses importantes"],
            "Objet du contrôle": [
                f"{Obj['police_presta']} polices avec des dépenses  d'un total de {Obj['depense_presta']} concernant {Obj['beneficiaire_presta']} bénéficiaires. Contrôle des consommations autorisées",
                f"{Obj['police_presta']} polices avec des dépenses  d'un total de {Obj['depense_presta']} concernant {Obj['beneficiaire_presta']} bénéficiaires. Contrôle du respect de la période de couverture contractuelle.",
                f"{Obj['police_presta']} polices avec des dépenses  d'un total de {Obj['depense_presta']} concernant {Obj['beneficiaire_presta']} bénéficiaires. Contrôle du respect desconditions d'âge.",
                f"{Obj['police_presta']} polices avec des dépenses  d'un total de {Obj['depense_presta']} concernant {Obj['beneficiaire_presta']} bénéficiaires, {Obj['enregistrements']}  enregistrements. Contrôle du respect des délais d'attente  contractuels.",
                f"{Obj['police_presta']} polices avec des dépenses  d'un total de {Obj['depense_presta']} concernant {Obj['beneficiaire_presta']} bénéficiaires, {Obj['enregistrements']}  enregistrements. Contrôle du respect de la procédure de gestion des dépenses importantes.",
                f"{Obj['police_presta']} polices avec des dépenses  d'un total de {Obj['depense_presta']} concernant {Obj['beneficiaire_presta']} bénéficiaires, {Obj['enregistrements']}  enregistrements. Contrôle du respect de la procédure de gestion des dépenses fréquentes.",
            ],
            "Résultat": rs["result"],
            "Anomalie": rs["anorm"]

        })
        if len(df_prestation) != 0:
            path_df_prestation = os.path.join(
                f"{base_directory}/{str(COUTIER_entry.get())}/RESULTATS", "PRESTATION.xlsx")
            df_prestation.to_excel(path_df_prestation, sheet_name="PRESTATION")
    else:
        # print("Attention: La taille des listes doit être égale.")
        pass
    return


def resume(file_benef, file_emission, file_conso, seuil_ctx, ante=pd.to_datetime(str(start_date_entry.get())), post=pd.to_datetime(str(end_date_entry.get()))):
    benef = pd.read_excel(file_benef)
    benef.columns = [to_camel_case(column) for column in benef.columns]

    benef["dateSortie"] = pd.to_datetime(
        benef["dateSortie"], format='%d/%m/%Y', errors='coerce', dayfirst=True)
    benef["dateSortie"] = benef["dateSortie"].apply(
        lambda x: pd.NaT if x >= post else x)
    benef = benef[(benef["dateSortie"] <= pd.to_datetime(str(post), dayfirst=True)) & (
        benef["dateSortie"] >= pd.to_datetime(str(ante), dayfirst=True)) | (benef["dateSortie"].isna())]

    emission = pd.read_excel(file_emission)
    emission.columns = [to_camel_case(column) for column in emission.columns]

    emission = emission[(emission["dateEmission"] <= pd.to_datetime(
        str(anterieure_date_emission_entry.get()), dayfirst=True))]
    emission["dateEmission"] = pd.to_datetime(
        emission["dateEmission"], format='%d/%m/%Y', errors='coerce', dayfirst=True)

    conso = pd.read_excel(file_conso)
    conso.columns = [to_camel_case(column) for column in conso.columns]

    benef_sans_doublon = benef.drop_duplicates()
    benef_sans_doublon.loc[:, "age"] = pd.to_datetime(
        benef_sans_doublon['dateNaissance'], dayfirst=True).apply(age)

    benef_sans_date_naiss = benef_sans_doublon[benef_sans_doublon['dateNaissance'].isna(
    )]
    sans_doublons_emission = emission.drop_duplicates()
    sans_doublons_conso = conso.drop_duplicates()

    ENFANT_SUP_25 = benef_sans_doublon[benef_sans_doublon["statutAce"].isin(
        ['E']) & (benef_sans_doublon['age'] > 25)]
    ADULTES_SUP_60 = benef_sans_doublon[benef_sans_doublon["statutAce"].ne(
        'E') & (benef_sans_doublon['age'] > 60)]

    liste_police_anormale = set(sans_doublons_conso["numPolice"]).difference(
        set(sans_doublons_emission["numPolice"]))
    cond1 = sans_doublons_conso['numPolice'].isin(liste_police_anormale) & pd.to_datetime(
        sans_doublons_conso["dateSoins"], dayfirst=True).gt(pd.to_datetime(str(start_date_entry.get()), dayfirst=True))
    police_anomale = sans_doublons_conso[cond1].astype(
        {'montantPaye': 'float'})
    police_anomale_resume = police_anomale.groupby(
        ['numPolice'])['montantPaye'].agg(['count', 'sum']).reset_index()

    cond2 = sans_doublons_conso['matricule'].isin(
        ADULTES_SUP_60["matricule"].unique())
    ADULTES_SUP_60_CONSO = sans_doublons_conso[cond2]

    cond3 = sans_doublons_conso['matricule'].isin(
        ENFANT_SUP_25["matricule"].unique())
    ENFANT_SUP_25_CONSO = sans_doublons_conso[cond3]

    if len(benef_sans_date_naiss) != 0:
        try:
            condition = sans_doublons_conso['matricule'].isin(
                benef_sans_date_naiss["matricule"].unique())
            benef_sans_date_naiss = sans_doublons_conso[condition]
        except:
            # print("benef_sans_date_naiss vide")
            pass

    cond4 = pd.to_datetime(sans_doublons_conso["dateSoins"], dayfirst=True).lt(pd.to_datetime(str(anterieure_dt_soins_conso.get()), dayfirst=True)) & pd.to_datetime(
        sans_doublons_conso["dateReception"], dayfirst=True).gt(pd.to_datetime(str(posterieure_date_recept_conso.get()), dayfirst=True))
    PRESTATIONS_TARDIVES = sans_doublons_conso[cond4]

    cond5 = sans_doublons_conso['montantPaye'].astype(
        "float").gt(float(seuil_ctx))
    CONSO_COUTEUSES = sans_doublons_conso[cond5]

    if len(CONSO_COUTEUSES) != 0:
        path_CONSO_COUTEUSES = os.path.join(
            f"{base_directory}/{str(COUTIER_entry.get())}/CONSOMMATION", "CONSO_COUTEUSES.xlsx")
        CONSO_COUTEUSES.to_excel(path_CONSO_COUTEUSES)

    Resume().ecart(selected_beneficiaire_path, selected_emission_path,
                   selected_consommation_path, COUTIER_entry.get(), end_date_entry.get())
    return


def start_progress():
    createDirs(base_directory,COUTIER_entry)
    progress_bar['value'] = 0
    critere_benef(selected_beneficiaire_path, pd.to_datetime(str(start_date_entry.get(
    )), dayfirst=True), pd.to_datetime(str(end_date_entry.get()), dayfirst=True))
    critere_emission(selected_emission_path, pd.to_datetime(
        str(end_date_entry.get()), dayfirst=True))
    critere_conso(selected_consommation_path)
    resume(selected_beneficiaire_path, selected_emission_path, selected_consommation_path, seuil_entry.get(), ante=pd.to_datetime(
        str(start_date_entry.get()), dayfirst=True), post=pd.to_datetime(str(end_date_entry.get()), dayfirst=True))
    update_progress()
    
    return


def update_progress():
    liste = [1, 2, 3, 4, 5, 6, 7, 8, 9]
    current_value = progress_bar["value"]
    if current_value < 100:
        increment = random.choice(liste)
        if current_value + increment > 100:
            increment = 100 - current_value
        progress_bar['value'] += increment
        current_value = progress_bar['value']
        progress_label.config(text=f"{current_value}%")
        window.after(1500, update_progress)
    else:
        progress_label.config(text="Terminé")


def reset_progress():
    global current_value
    if progress_label.cget("text") == "Terminé" and current_value == 100:
        current_value = 0
        progress_bar['value'] = 0
        progress_label.config(text="0%")


# ----------------------------------------------------------------------

# Add KeyRelease events to the entry widgets
start_date_entry.bind("<KeyRelease>", lambda event: validateButton(start_button, start_date_entry.get(), end_date_entry.get(), seuil_entry.get(), COUTIER_entry.get(),beneficiaire_entry.get() ,emission_entry.get() ,consommation_entry.get()))
end_date_entry.bind("<KeyRelease>", lambda event: validateButton(start_button, start_date_entry.get(), end_date_entry.get(), seuil_entry.get(), COUTIER_entry.get(),beneficiaire_entry.get() ,emission_entry.get() ,consommation_entry.get()))
seuil_entry.bind("<KeyRelease>", lambda event: validateButton(start_button, start_date_entry.get(), end_date_entry.get(), seuil_entry.get(), COUTIER_entry.get(),beneficiaire_entry.get() ,emission_entry.get() ,consommation_entry.get()))
COUTIER_entry.bind("<KeyRelease>", lambda event: validateButton(start_button, start_date_entry.get(), end_date_entry.get(), seuil_entry.get(), COUTIER_entry.get(),beneficiaire_entry.get() ,emission_entry.get() ,consommation_entry.get()))

# Add a button to start the control process
start_button = ttk.Button(window, text="COMMENCER LE CONTROLE",
                          style="danger.Outline.TButton", command=start_progress, state="disabled")
start_button.grid(row=18, column=1, columnspan=3,
                  pady=10, ipadx=2, sticky="we")

beneficiaire_entry.bind("<FocusOut>", handle_beneficiaire_selection(start_button,start_date_entry,end_date_entry,seuil_entry,COUTIER_entry,beneficiaire_entry,emission_entry,consommation_entry))
emission_entry.bind("<FocusOut>", handle_emission_selection(start_button,start_date_entry,end_date_entry,seuil_entry,COUTIER_entry,beneficiaire_entry,emission_entry,consommation_entry))
consommation_entry.bind("<FocusOut>", handle_consommation_selection(start_button,start_date_entry,end_date_entry,seuil_entry,COUTIER_entry,beneficiaire_entry,emission_entry,consommation_entry))

# Create a progress bar
progress_bar = ttk.Progressbar(window, orient="horizontal",
                               length=200, mode="determinate", takefocus=True, maximum=100)
progress_bar.grid(row=19, column=1, columnspan=3,
                  pady=10, padx=10, sticky="we")

# Add a label to display the percentage
progress_label = ttk.Label(window, text="0%")
progress_label.grid(row=20, column=2, columnspan=3,
                    pady=10, padx=10, sticky="we")
# Start the Tkinter main loop

window.mainloop()
