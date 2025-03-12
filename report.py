import requests
import pandas as pd
import json
from datetime import datetime, timedelta
from workalendar.europe import France
import re

# 🔧 Configuration GitLab
GITLAB_URL = "https://gitlab.com"
PRIVATE_TOKEN = "***"
PROJECT_ID = 136  # Postman

# Désactivation des avertissements SSL (Hyper nécessaire problème de certification)
requests.packages.urllib3.disable_warnings()

# 📅 Gestion des jours fériés et jours ouvrés
cal = France()
aujourd_hui = datetime.now()
premier_jour_mois_actuel = aujourd_hui.replace(day=1)

def avancer_jours(date, jours):
    """Avance une date en excluant les week-ends et jours fériés. => Idées Laurent"""
    while jours > 0:
        date -= timedelta(days=1)
        if date.weekday() < 5 and not cal.is_holiday(date):
            jours -= 1
    return date.strftime("%d-%m-%Y")

# 🎯 Fonctions pour extraire les informations des issues
# Demande d'extraction de priorités
"""def extraire_priorite(labels):
    for label in labels:
        if label.startswith("Priorité::"):
            return label.split("::")[-1]
    return "Non définie"""
def extraire_priorite(labels):
    for label in labels:
        if label == "Priorité::0":
            return "P0"
        elif label == "Priorité::1":
            return "P1"
        elif label == "Priorité::2":
            return "P2"
        elif label == "Priorité::3":
            return "P3"
    return "Non définie"

# Demande de traduction
def renommer_statut(state):
    return "Ouvert" if state == "opened" else "Fermé"
# Demande d'extraction de labels
def extraire_statut(labels):
    statut_map = {
        "workflow::To Do": "A faire",
        "workflow::Doing": "En cours",
        "workflow::Shipped": "Terminé",
        "workflow::Stand-by": "En attente",
        "workflow::Ready": "Prêt",
        "EN ATTENTE DE VALIDATION": "Attente de validation",
        "EN ATTENTE ": "Attente "
    }
    return next((statut_map[label] for label in labels if label in statut_map), "Non défini")
# Demande d'extraction de USECASE
def extraire_usecase(tache):
    match = re.match(r'\[(.*?)\]', tache)
    return match.group(1) if match else "Non défini"
# Demande d'extraction de la phase
def extraire_phase(labels):
    if "MCO" in labels:
        return "MCO"
    elif "PILOTE" in labels:
        return "PILOTE"
    return "Non défini"
# Demande d'extraction du type
def extraire_type(labels):
    if "OPERATIONNEL" in labels:
        return "OPÉRATIONNEL"
    elif "Bogue" in labels:
        return "BUG"
    elif "EVOLUTION" in labels:
        return "ÉVOLUTION"
    return "Non défini"
# Demande d'extraction de l'état
def determiner_responsable(etat):
    if etat in ["Attente ", "Attente de validation"]:
        return "Externe"
    return "Interne" if etat != "Non défini" else "Non défini"

# 🔍 Récupération des événements de label via API
def recuperer_label_events(issue_iid):
    """
    Récupère les événements de labels pour une issue via l'API GitLab.
    Identifie la date de prise en charge et la date de résolution.
    #pour plus d'info sur la raison derrière cette méthode : https://stackoverflow.com/questions/74125721/gitlab-how-to-get-the-label-added-date-by-gitlab-api
    # ou  https://docs.gitlab.com/api/resource_label_events/
    """
    url = f"{GITLAB_URL}/api/v4/projects/{PROJECT_ID}/issues/{issue_iid}/resource_label_events"
    headers = {"PRIVATE-TOKEN": PRIVATE_TOKEN}

    try:
        response = requests.get(url, headers=headers, verify=False)
        response.raise_for_status()
        events = response.json()

        date_prise_en_charge = None
        date_resolution = None

        for event in events:
            label_name = event.get("label", {}).get("name")
            action = event.get("action")
            event_date = event.get("created_at")

            if label_name in ["workflow::Doing"] and action == "add":
                date_prise_en_charge = event_date
            if label_name in ["EN ATTENTE DE VALIDATION", "EN ATTENTE "] and action == "add":
                date_resolution = event_date

        return date_prise_en_charge, date_resolution

    except requests.exceptions.RequestException as e:
        print(f"⚠️ Erreur lors de la récupération des événements pour l'issue {issue_iid}: {e}")
        return None, None

# 📥 Récupération des issues du projet
def recuperer_issues():
    """
    Récupère toutes les issues du projet GitLab et extrait les informations nécessaires.
    """
    url = f"{GITLAB_URL}/api/v4/projects/{PROJECT_ID}/issues"
    headers = {"PRIVATE-TOKEN": PRIVATE_TOKEN}

    try:
        response = requests.get(url, headers=headers, verify=False, params={"per_page": 100})
        response.raise_for_status()
        issues = response.json()

        issues_data = []

        premier_jour_mois_precedent = premier_jour_mois_actuel - timedelta(days=1)
        date_limite = premier_jour_mois_precedent.replace(day=1)
        print(f"📅 Date limite pour les issues fermées : {date_limite}")

        for issue in issues:
            if "WFRESIL" in issue.get("labels", []) or issue.get("confidential", False):
                continue

            created_at = datetime.strptime(issue["created_at"], "%Y-%m-%dT%H:%M:%S.%fZ")
            updated_at = datetime.strptime(issue["updated_at"], "%Y-%m-%dT%H:%M:%S.%fZ")
            closed_at = issue.get("closed_at")
            if isinstance(issue, dict):
                labels = issue.get('labels', [])
                issue_state = issue.get('state', '')
                issue_closed_at = issue.get('closed_at', None)
                issue_created_at = issue.get('created_at', '')
                issue_updated_at = issue.get('updated_at', '')

            if closed_at:
                closed_at = datetime.strptime(closed_at, "%Y-%m-%dT%H:%M:%S.%fZ")
                if closed_at < date_limite:
                    continue

            time_estimate_days = issue.get("time_stats", {}).get("time_estimate", 0) / 86400
            # Gestion de la date de fin
            if issue_state == "closed" and issue_closed_at:
                date_fin = datetime.strptime(issue_closed_at, "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d-%m-%Y")
            elif 'time_stats' in issue and issue['time_stats'].get('time_estimate', 0) > 0:
                time_estimate_seconds = issue['time_stats']['time_estimate']
                time_estimate_days = time_estimate_seconds / 86400  # Secondes en jours
                date_fin = avancer_jours(
                    datetime.strptime(issue_created_at, "%Y-%m-%dT%H:%M:%S.%fZ") + timedelta(days=time_estimate_days),
                    0)
            else:
                date_fin = ""
            date_prise_en_charge, date_resolution = recuperer_label_events(issue["iid"])
            date_prise_en_charge = date_prise_en_charge or "Non définie"
            date_resolution = date_resolution or "Non définie"
            priorite_label = extraire_priorite(labels)

            labels = issue.get("labels", [])
            issue_info = {
                "LIBELLE": issue["title"],
                "USECASE": extraire_usecase(issue["title"]),
                "PHASE": extraire_phase(labels),
                "TYPE": extraire_type(labels),
                "STATUT": renommer_statut(issue["state"]),
                "RESPONSABLE": determiner_responsable(extraire_statut(labels)),
                "PRIORITE": priorite_label,
                "ETAT": extraire_statut(labels),
                "AUTEUR": issue["author"]["name"],
                "DATE OUVERTURE": created_at.strftime("%d-%m-%Y"),
                "DERNIERE MISE A JOUR": updated_at.strftime("%d-%m-%Y"),
                "DATE FIN": date_fin or "",
                "DATE DE PRISE EN CHARGE": date_prise_en_charge,
                "DATE DE RÉSOLUTION": date_resolution,
                "REFERENCE_ID": issue["iid"],
                "URL": f"{GITLAB_URL}/caciis/sfr-obsope-audit/-/issues/{issue['iid']}",
                "LABELS": ", ".join(labels)
            }
            issues_data.append(issue_info)

        return issues_data

    except requests.exceptions.RequestException as e:
        print(f"❌ Erreur lors de la récupération des issues : {e}")
        return []

# 📤 Exportation en Excel
issues_data = recuperer_issues()
df = pd.DataFrame(issues_data)
for col in ["DATE OUVERTURE", "DERNIERE MISE A JOUR", "DATE FIN", "DATE DE PRISE EN CHARGE", "DATE DE RÉSOLUTION"]:
    df[col] = pd.to_datetime(df[col], format="%Y-%m-%dT%H:%M:%S.%fZ", errors="coerce").dt.strftime("%d-%m-%Y")


# 📤 Exportation en Excel
def exporter_en_excel(issues_data):
    df = pd.DataFrame(issues_data)

    # 🗓️ Correction du format des dates
    for col in ["DATE DE PRISE EN CHARGE", "DATE DE RÉSOLUTION"]:
        df[col] = pd.to_datetime(df[col], format="%Y-%m-%dT%H:%M:%S.%fZ", errors="coerce").dt.strftime("%d-%m-%Y")

    # 🔍 Nom du fichier
    fichier_excel = f"issues_export_Nect_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

    with pd.ExcelWriter(fichier_excel, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Issues", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Issues"]

        # 🎨 Appliquer un style d'en-tête
        header_format = workbook.add_format(
            {'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'align': 'center'})

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 20)  # Ajuste la largeur des colonnes

        # 🧊 Geler la première ligne
        worksheet.freeze_panes(1, 0)

        # 🎨 Mise en forme conditionnelle pour le statut
        status_colors = {
            "A faire": "#EBF0FA",
            "En cours": "#FFF5CC",
            "Terminé": "#F7E8A4",
            "En attente": "#D8C3E8",
            "Prêt": "#E1EEC5",
            "Attente de validation": "#A3D9A5",
            "Attente SFR": "#E8B4A5"
        }

        # Trouver l'index de la colonne "ETAT"
        col_etat_index = df.columns.get_loc("ETAT")
        col_letter = chr(65 + col_etat_index)  # Convertit l'index en lettre Excel (A=65)

        # 📊 Appliquer les couleurs sur la colonne "ETAT"
        for status, color in status_colors.items():
            format_statut = workbook.add_format({'bg_color': color})
            worksheet.conditional_format(f'{col_letter}2:{col_letter}{len(df) + 1}', {
                'type': 'text',
                'criteria': 'containing',
                'value': status,
                'format': format_statut
            })

    print(f"✅ Fichier généré : {fichier_excel}")


# Appel de la fonction après récupération des issues
issues_data = recuperer_issues()
exporter_en_excel(issues_data)







