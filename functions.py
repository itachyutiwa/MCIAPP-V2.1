from tkinter.simpledialog import askstring
from tkinter import messagebox
import os
import tkinter as tk

def validateButton(start_btn,entry_date,end_date,seuil_entry,courtier_entry,beneficiaire_entry ,emission_entry , consommation_entry):
    def entryState():
        if str(entry_date) == "" or str(end_date) == "" or str(seuil_entry) == "" or str(courtier_entry) == "" or str(beneficiaire_entry) == "" or str(emission_entry) == "" or str(consommation_entry) == "":
            return False
        elif str(entry_date) != "" and str(end_date) != "" and str(seuil_entry) != "" and str(courtier_entry) != "" and str(beneficiaire_entry) != "" and str(emission_entry)  != "" and str(consommation_entry) != "":
            return True

    if entryState() == True:
        start_btn.config(state="normal")
    elif entryState() == False :
        start_btn.config(state="disabled")

def handle_beneficiaire_selection(start_button,start_date_entry,end_date_entry,seuil_entry,COUTIER_entry,beneficiaire_entry,emission_entry,consommation_entry):
    validateButton(
        start_button,
        start_date_entry.get(),
        end_date_entry.get(),
        seuil_entry.get(),
        COUTIER_entry.get(),
        beneficiaire_entry.get(),
        emission_entry.get(),
        consommation_entry.get(),
    )

def handle_emission_selection(start_button,start_date_entry,end_date_entry,seuil_entry,COUTIER_entry,beneficiaire_entry,emission_entry,consommation_entry):
    validateButton(
        start_button,
        start_date_entry.get(),
        end_date_entry.get(),
        seuil_entry.get(),
        COUTIER_entry.get(),
        beneficiaire_entry.get(),
        emission_entry.get(),
        consommation_entry.get(),
    )

def handle_consommation_selection(start_button,start_date_entry,end_date_entry,seuil_entry,COUTIER_entry,beneficiaire_entry,emission_entry,consommation_entry):
    validateButton(
        start_button,
        start_date_entry.get(),
        end_date_entry.get(),
        seuil_entry.get(),
        COUTIER_entry.get(),
        beneficiaire_entry.get(),
        emission_entry.get(),
        consommation_entry.get(),
    )

def createDirs(base_directory,COUTIER_entry):
    courtier_name = str(COUTIER_entry.get())

    def getCourtier():
        return str(COUTIER_entry.get())

    courtier_path = os.path.join(base_directory, getCourtier())
    if COUTIER_entry is not None and not os.path.exists(courtier_path):
        # Create the entire directory path for the courtier
        os.makedirs(courtier_path)

        # Create the subdirectories inside the courtier directory
        dossier_production_path = os.path.join(courtier_path, "PRODUCTION")
        os.makedirs(dossier_production_path)

        dossier_consommation_path = os.path.join(courtier_path, "CONSOMMATION")
        os.makedirs(dossier_consommation_path)

        dossier_resultats_path = os.path.join(courtier_path, "RESULTATS")
        os.makedirs(dossier_resultats_path)
        if not os.path.exists(dossier_resultats_path):
            os.makedirs(dossier_resultats_path)
    elif os.path.exists(os.path.join(base_directory, str(COUTIER_entry.get()))) and COUTIER_entry is not None:
        # Afficher une boîte de dialogue de type "warning"
        messagebox.showwarning(
            "Avertissement", f"Le dossier '{courtier_name}' existe déjà.")

        # Demander à l'utilisateur de saisir un nouveau nom de dossier
        new_courtier_name = askstring(
            "Nouveau Nom", f"Entrez un nouveau nom à la place de '{courtier_name}':")
        COUTIER_entry.delete(0, tk.END)
        COUTIER_entry.insert(0, new_courtier_name)
        dossier_production_path = os.path.join(
            base_directory, str(new_courtier_name), "PRODUCTION")
        if not os.path.exists(dossier_production_path):
            os.makedirs(dossier_production_path)

        dossier_consommation_path = os.path.join(
            base_directory, str(new_courtier_name), "CONSOMMATION")
        if not os.path.exists(dossier_consommation_path):
            os.makedirs(dossier_consommation_path)

        dossier_resultats_path = os.path.join(
            base_directory, str(new_courtier_name), "RESULTATS")
        if not os.path.exists(dossier_resultats_path):
            os.makedirs(dossier_resultats_path)
     

