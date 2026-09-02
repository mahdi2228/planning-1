ALLUCO — Planning IA Agentique v3
================================

OBJECTIF
- Input direct : "Base Commandes Client encours confirmées par mois.xlsx"
- Historique optionnel : un ancien "Planning Sxx.xlsx"
- Output : six feuilles Planning Lundi -> Planning SAMEDI, avec les 28 colonnes métier ALLUCO.

INSTALLATION
1) python -m pip install -r requirements.txt
2) streamlit run app.py

UTILISATION
1) Déposer la base commandes.
2) Déposer un ancien planning (recommandé) : l'application apprend PoidsUn, Barre/bal, Stock brut, Nuance,
   coefficient poudre et minutes/balancelle.
3) Choisir semaine, capacité et objectif IA.
4) Cliquer "Générer le meilleur planning IA".
5) Vérifier les six jours puis télécharger l'Excel.

MOTEUR AGENTIQUE
- Agent Données : nettoie et filtre les commandes encours.
- Agent Référentiel : apprend les paramètres techniques de l'historique.
- Agent Quantités : calcule Lancement, PoidsT, Poudre, Nbre Bal, tps.
- Agent Priorités : combine délai, ancienneté, OF, réservation brut et disponibilité.
- Agent Planificateur : OR-Tools CP-SAT affecte les jobs aux 6 jours sous contraintes.
- Agent Couleurs : regroupe/séquence les couleurs.
- Agent Critique : contrôle capacité, inconnues et backlog prioritaire.
- Export : Excel métier identique au format Planning S36 fourni.

POINTS IMPORTANTS CORRIGÉS PAR RAPPORT À L'ANCIENNE VERSION
- Une commande peut avoir plusieurs lignes/articles/OF : aucune fusion sur NumCommande.
- Lancement est une QUANTITÉ, pas une date.
- tps = Nbre Bal × minutes/bal ; le Planning S36 fourni apprend 4 min/bal.
- Poudre = PoidsT × coefficient appris ; le Planning S36 apprend environ 0,052.
- Article/int et Couleur sont dérivés du dernier tiret de Article.
- Le bouton natif d'ouverture/fermeture de la sidebar Streamlit n'est jamais masqué.

TEST SANS STREAMLIT
python app.py --self-test --input "Base Commandes Client encours confirmées par mois.xlsx" --history "Planning S36.xlsx"

GÉNÉRATION CLI
python app.py --generate --input "Base Commandes Client encours confirmées par mois.xlsx" --history "Planning S36.xlsx" --year 2026 --week 36 --output "Planning_IA_S36.xlsx"
