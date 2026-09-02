# -*- coding: utf-8 -*-
"""
db.py — Couche de données pour ALLUCO Planning Laquage.

Stockage 100% fichiers locaux (JSON / CSV), pas de base de données externe,
cohérent avec le reste d'ALLUCO. Ce module expose des fonctions simples de
lecture / écriture ainsi qu'un jeu de données de démonstration généré au
premier lancement si aucun fichier n'existe encore.

Structure des fichiers (dossier data/) :
    - couleurs.json      : référentiel des couleurs (RAL, famille, clarté)
    - transitions.json    : coûts / temps de nettoyage entre couleurs
    - commandes.json      : commandes à planifier / planifiées
    - planning.json       : affectation planning (jour -> lignes)
    - parametres.json     : paramètres généraux (capacité, coefficients...)
    - historique.json     : semaines validées
"""

from __future__ import annotations

import json
import os
import random
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

# --------------------------------------------------------------------------
# Chemins
# --------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

F_COULEURS = DATA_DIR / "couleurs.json"
F_TRANSITIONS = DATA_DIR / "transitions.json"
F_COMMANDES = DATA_DIR / "commandes.json"
F_PLANNING = DATA_DIR / "planning.json"
F_PARAMETRES = DATA_DIR / "parametres.json"
F_HISTORIQUE = DATA_DIR / "historique.json"

JOURS_SEMAINE = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]
JOURS_SEMAINE_COURT = ["LUN", "MAR", "MER", "JEU", "VEN", "SAM"]

STATUTS = ["À PLANIFIER", "PLANIFIÉ", "EN PRÉPARATION", "EN LAQUAGE", "TERMINÉ", "BLOQUÉ"]
PRIORITES = ["NORMAL", "HAUTE", "URGENTE", "CRITIQUE"]


# --------------------------------------------------------------------------
# Utilitaires génériques JSON
# --------------------------------------------------------------------------

def _load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return default


def _save_json(path: Path, data: Any) -> None:
    tmp_path = path.with_suffix(".tmp")
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp_path.replace(path)


# --------------------------------------------------------------------------
# Référentiel couleurs
# --------------------------------------------------------------------------

DEFAULT_COULEURS = [
    {"couleur": "NOIR", "ral": "-", "famille": "Très foncé", "clarte": 1},
    {"couleur": "DARK", "ral": "-", "famille": "Très foncé", "clarte": 1},
    {"couleur": "ACAJOU FONCÉ", "ral": "-", "famille": "Foncé", "clarte": 2},
    {"couleur": "GRIS FONCÉ", "ral": "-", "famille": "Foncé", "clarte": 2},
    {"couleur": "ACAJOU", "ral": "-", "famille": "Moyen", "clarte": 3},
    {"couleur": "GRIS", "ral": "-", "famille": "Moyen", "clarte": 3},
    {"couleur": "GRISG", "ral": "-", "famille": "Moyen", "clarte": 3},
    {"couleur": "ACAJOU CLAIR", "ral": "-", "famille": "Clair", "clarte": 4},
    {"couleur": "GRIS CLAIR", "ral": "-", "famille": "Clair", "clarte": 4},
    {"couleur": "R9016", "ral": "9016", "famille": "Très clair", "clarte": 5},
    {"couleur": "BLC", "ral": "-", "famille": "Très clair", "clarte": 5},
]


def get_couleurs() -> List[Dict[str, Any]]:
    return _load_json(F_COULEURS, DEFAULT_COULEURS)


def save_couleurs(couleurs: List[Dict[str, Any]]) -> None:
    _save_json(F_COULEURS, couleurs)


def get_couleur_info(nom: str) -> Optional[Dict[str, Any]]:
    for c in get_couleurs():
        if c["couleur"] == nom:
            return c
    return None


def get_clarte(nom: str) -> int:
    info = get_couleur_info(nom)
    return info["clarte"] if info else 3


# --------------------------------------------------------------------------
# Transitions (coût / temps de nettoyage entre couleurs)
# --------------------------------------------------------------------------

def get_transitions() -> List[Dict[str, Any]]:
    return _load_json(F_TRANSITIONS, [])


def save_transitions(transitions: List[Dict[str, Any]]) -> None:
    _save_json(F_TRANSITIONS, transitions)


def get_transition_info(depart: str, arrivee: str) -> Dict[str, Any]:
    """Renvoie le coût et le temps de nettoyage entre deux couleurs.

    Si la transition n'est pas explicitement définie dans transitions.json,
    elle est calculée à partir de l'écart de niveau de clarté (règle par
    défaut, configurable ultérieurement dans Paramètres > Transitions).
    """
    for t in get_transitions():
        if t["depart"] == depart and t["arrivee"] == arrivee:
            return t
    ecart = abs(get_clarte(depart) - get_clarte(arrivee))
    cout = ecart * 4
    temps_min = 10 if ecart <= 1 else (20 if ecart == 2 else 35 if ecart == 3 else 45)
    return {"depart": depart, "arrivee": arrivee, "cout": cout, "temps_nettoyage": temps_min}


def transition_alerte(depart: str, arrivee: str) -> Optional[Dict[str, Any]]:
    """Retourne un dict d'alerte si la transition est déconseillée (grand
    écart de clarté), sinon None."""
    ecart = abs(get_clarte(depart) - get_clarte(arrivee))
    if ecart >= 3:
        couleurs = get_couleurs()
        # Suggestion : chercher une couleur intermédiaire
        c1, c2 = get_clarte(depart), get_clarte(arrivee)
        lo, hi = min(c1, c2), max(c1, c2)
        intermediaires = [
            c["couleur"] for c in couleurs
            if lo < c["clarte"] < hi and c["couleur"] not in (depart, arrivee)
        ]
        return {
            "depart": depart,
            "arrivee": arrivee,
            "ecart": ecart,
            "suggestion": intermediaires,
        }
    return None


# --------------------------------------------------------------------------
# Paramètres
# --------------------------------------------------------------------------

DEFAULT_PARAMETRES = {
    "capacite_quotidienne_h": 15.0,
    "temps_par_balancelle_min": 6.5,
    "coefficient_poudre_kg_par_kg": 0.052,
    "jours_travailles": JOURS_SEMAINE,
    "horaire_debut": "07:00",
    "horaire_fin": "17:00",
}


def get_parametres() -> Dict[str, Any]:
    p = _load_json(F_PARAMETRES, {})
    merged = {**DEFAULT_PARAMETRES, **p}
    return merged


def save_parametres(parametres: Dict[str, Any]) -> None:
    _save_json(F_PARAMETRES, parametres)


# --------------------------------------------------------------------------
# Commandes
# --------------------------------------------------------------------------

def get_commandes() -> List[Dict[str, Any]]:
    return _load_json(F_COMMANDES, [])


def save_commandes(commandes: List[Dict[str, Any]]) -> None:
    _save_json(F_COMMANDES, commandes)


def get_commande(commande_id: str) -> Optional[Dict[str, Any]]:
    for c in get_commandes():
        if c["id"] == commande_id:
            return c
    return None


def update_commande_statut(commande_id: str, statut: str) -> None:
    commandes = get_commandes()
    for c in commandes:
        if c["id"] == commande_id:
            c["statut"] = statut
    save_commandes(commandes)


# --------------------------------------------------------------------------
# Planning
# --------------------------------------------------------------------------

def get_planning() -> Dict[str, List[Dict[str, Any]]]:
    """Renvoie {jour: [lignes de planning]}."""
    return _load_json(F_PLANNING, {j: [] for j in JOURS_SEMAINE})


def save_planning(planning: Dict[str, List[Dict[str, Any]]]) -> None:
    _save_json(F_PLANNING, planning)


def get_planning_jour(jour: str) -> List[Dict[str, Any]]:
    return get_planning().get(jour, [])


def save_planning_jour(jour: str, lignes: List[Dict[str, Any]]) -> None:
    planning = get_planning()
    planning[jour] = lignes
    save_planning(planning)


def calculer_charge_jour(lignes: List[Dict[str, Any]]) -> Dict[str, float]:
    """Calcule charge production, changements couleur et charge totale (h)."""
    params = get_parametres()
    temps_bal_h = params["temps_par_balancelle_min"] / 60.0
    charge_production = sum(l.get("bal", 0) * temps_bal_h for l in lignes)

    ordre = [l["couleur"] for l in sorted(lignes, key=lambda x: x.get("ordre", 0))]
    changements_h = 0.0
    for i in range(1, len(ordre)):
        if ordre[i] != ordre[i - 1]:
            t = get_transition_info(ordre[i - 1], ordre[i])
            changements_h += t["temps_nettoyage"] / 60.0

    return {
        "charge_production": round(charge_production, 2),
        "changements_couleur": round(changements_h, 2),
        "charge_totale": round(charge_production + changements_h, 2),
        "capacite": params["capacite_quotidienne_h"],
    }


def sequence_couleurs_jour(lignes: List[Dict[str, Any]]) -> List[str]:
    """Séquence de couleurs (sans doublons consécutifs) dans l'ordre du jour."""
    ordre = [l["couleur"] for l in sorted(lignes, key=lambda x: x.get("ordre", 0))]
    seq: List[str] = []
    for c in ordre:
        if not seq or seq[-1] != c:
            seq.append(c)
    return seq


def optimiser_ordre_couleurs(couleurs: List[str]) -> List[str]:
    """Tri glouton des couleurs à produire par clarté croissante (ou
    décroissante selon le point de départ), pour minimiser le coût total
    de transition. Simple et explicable pour un responsable d'atelier."""
    couleurs_uniques = list(dict.fromkeys(couleurs))
    return sorted(couleurs_uniques, key=lambda c: get_clarte(c))


def cout_total_sequence(sequence: List[str]) -> Dict[str, float]:
    cout = 0
    temps_min = 0
    for i in range(1, len(sequence)):
        t = get_transition_info(sequence[i - 1], sequence[i])
        cout += t["cout"]
        temps_min += t["temps_nettoyage"]
    return {"cout": cout, "temps_min": temps_min}


# --------------------------------------------------------------------------
# Historique
# --------------------------------------------------------------------------

def get_historique() -> List[Dict[str, Any]]:
    return _load_json(F_HISTORIQUE, [])


def save_historique(historique: List[Dict[str, Any]]) -> None:
    _save_json(F_HISTORIQUE, historique)


def archiver_semaine(semaine: int, annee: int, statut: str = "VALIDÉ") -> None:
    planning = get_planning()
    total_h = sum(calculer_charge_jour(planning.get(j, []))["charge_totale"] for j in JOURS_SEMAINE)
    couleurs = {l["couleur"] for j in JOURS_SEMAINE for l in planning.get(j, [])}
    commandes = get_commandes()

    historique = get_historique()
    historique.insert(0, {
        "semaine": semaine,
        "annee": annee,
        "statut": statut,
        "heures": round(total_h, 2),
        "nb_couleurs": len(couleurs),
        "nb_commandes": len(commandes),
        "date_validation": datetime.now().isoformat(timespec="seconds"),
    })
    save_historique(historique)


# --------------------------------------------------------------------------
# Jeu de données de démonstration (généré une seule fois)
# --------------------------------------------------------------------------

def _demo_commandes() -> List[Dict[str, Any]]:
    clients = ["Sofal", "Alutec", "Fenal", "Menuis Pro", "Delta Alu", "Tunal", "Alu Concept"]
    articles = ["Profilé PA", "Profilé PB", "Volet roulant", "Chassis", "Coulissant", "Barre L", "Cornière"]
    couleurs_demo = [c["couleur"] for c in DEFAULT_COULEURS]
    commandes = []
    for i in range(1, 43):
        commandes.append({
            "id": f"CMD-{1000 + i}",
            "commande": f"CMD-{1000 + i}",
            "client": random.choice(clients),
            "article": random.choice(articles),
            "couleur": random.choice(couleurs_demo),
            "reste": random.randint(50, 800),
            "stock": random.randint(0, 200),
            "of": f"OF-{2200 + i}",
            "priorite": random.choices(PRIORITES, weights=[55, 25, 15, 5])[0],
            "statut": random.choices(STATUTS, weights=[30, 30, 15, 15, 8, 2])[0],
            "date": (date.today() + timedelta(days=random.randint(-3, 10))).isoformat(),
        })
    return commandes


def _demo_planning() -> Dict[str, List[Dict[str, Any]]]:
    commandes = get_commandes()
    couleurs_demo = [c["couleur"] for c in DEFAULT_COULEURS]
    planning: Dict[str, List[Dict[str, Any]]] = {j: [] for j in JOURS_SEMAINE}
    idx = 0
    for j in JOURS_SEMAINE:
        n = random.randint(5, 8)
        ordre = 1
        jour_couleurs = sorted(random.sample(couleurs_demo, k=min(3, len(couleurs_demo))),
                                key=lambda c: next(x["clarte"] for x in DEFAULT_COULEURS if x["couleur"] == c))
        for _ in range(n):
            if idx >= len(commandes):
                break
            cmd = commandes[idx]
            couleur = random.choice(jour_couleurs)
            bal = random.randint(10, 40)
            poids_t = round(bal * random.uniform(0.018, 0.028), 3)
            planning[j].append({
                "ordre": ordre,
                "commande_id": cmd["id"],
                "commande": cmd["commande"],
                "client": cmd["client"],
                "article": cmd["article"],
                "couleur": couleur,
                "reste": cmd["reste"],
                "lancement": (date.today()).isoformat(),
                "re_laquage": "",
                "poids_t": poids_t,
                "poudre_kg": round(poids_t * 1000 * DEFAULT_PARAMETRES["coefficient_poudre_kg_par_kg"], 1),
                "barre_par_bal": random.randint(8, 16),
                "bal": bal,
                "temps_h": round(bal * DEFAULT_PARAMETRES["temps_par_balancelle_min"] / 60.0, 2),
                "verrouille": False,
            })
            ordre += 1
            idx += 1
    return planning


def init_demo_data(force: bool = False) -> None:
    """Génère un jeu de données de démonstration cohérent si les fichiers
    n'existent pas encore (ou si force=True)."""
    if force or not F_COULEURS.exists():
        save_couleurs(DEFAULT_COULEURS)
    if force or not F_PARAMETRES.exists():
        save_parametres(DEFAULT_PARAMETRES)
    if force or not F_TRANSITIONS.exists():
        save_transitions([])
    if force or not F_COMMANDES.exists():
        random.seed(36)
        save_commandes(_demo_commandes())
    if force or not F_PLANNING.exists():
        random.seed(36)
        save_planning(_demo_planning())
    if force or not F_HISTORIQUE.exists():
        save_historique([
            {"semaine": 35, "annee": 2026, "statut": "VALIDÉ", "heures": 84.6,
             "nb_couleurs": 7, "nb_commandes": 39, "date_validation": "2026-08-29T17:05:00"},
            {"semaine": 34, "annee": 2026, "statut": "VALIDÉ", "heures": 81.2,
             "nb_couleurs": 6, "nb_commandes": 35, "date_validation": "2026-08-22T17:10:00"},
        ])
