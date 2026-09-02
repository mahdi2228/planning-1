# -*- coding: utf-8 -*-
"""
ALLUCO — Planning & Pilotage Laquage
Application Streamlit, light mode, stockage local (JSON/CSV via db.py).
"""

import os
from datetime import date

import pandas as pd
import streamlit as st

import db

# ==========================================================================
# CONFIG PAGE
# ==========================================================================

st.set_page_config(
    page_title="ALLUCO | Planning Laquage",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="expanded",
)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo.png")

# ==========================================================================
# PALETTE / CSS
# ==========================================================================

COLOR_BG = "#F7F8FA"
COLOR_CARD = "#FFFFFF"
COLOR_TEXT = "#172033"
COLOR_TEXT_SECONDARY = "#667085"
COLOR_BORDER = "#E5E7EB"
COLOR_PRIMARY = "#2563EB"
COLOR_SUCCESS = "#16A34A"
COLOR_WARNING = "#F59E0B"
COLOR_ERROR = "#DC2626"
COLOR_INFO = "#2563EB"

CUSTOM_CSS = f"""
<style>
    .stApp {{
        background-color: {COLOR_BG};
    }}
    #MainMenu, footer {{visibility: hidden;}}

    section[data-testid="stSidebar"] {{
        background-color: {COLOR_CARD};
        border-right: 1px solid {COLOR_BORDER};
    }}
    section[data-testid="stSidebar"] .block-container {{
        padding-top: 1.2rem;
    }}

    h1, h2, h3, h4, h5, h6, p, span, label, div {{
        color: {COLOR_TEXT};
    }}

    .alluco-header {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 0.9rem 1.2rem;
        background-color: {COLOR_CARD};
        border: 1px solid {COLOR_BORDER};
        border-radius: 14px;
        margin-bottom: 1.4rem;
    }}
    .alluco-header-title {{
        font-size: 1.15rem;
        font-weight: 700;
        color: {COLOR_TEXT};
    }}
    .alluco-header-sub {{
        font-size: 0.82rem;
        color: {COLOR_TEXT_SECONDARY};
    }}

    .alluco-card {{
        background-color: {COLOR_CARD};
        border: 1px solid {COLOR_BORDER};
        border-radius: 14px;
        padding: 1.1rem 1.3rem;
        box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
        margin-bottom: 1rem;
    }}

    .alluco-kpi-label {{
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.04em;
        color: {COLOR_TEXT_SECONDARY};
        text-transform: uppercase;
        margin-bottom: 0.25rem;
    }}
    .alluco-kpi-value {{
        font-size: 1.55rem;
        font-weight: 700;
        color: {COLOR_TEXT};
    }}

    .badge {{
        display: inline-block;
        padding: 0.18rem 0.65rem;
        border-radius: 999px;
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.02em;
    }}
    .badge-normal {{ background-color: #F1F2F4; color: {COLOR_TEXT_SECONDARY}; }}
    .badge-haute {{ background-color: #DBEAFE; color: #1D4ED8; }}
    .badge-urgente {{ background-color: #FEF3C7; color: #B45309; }}
    .badge-critique {{ background-color: #FEE2E2; color: #B91C1C; }}
    .badge-success {{ background-color: #DCFCE7; color: #15803D; }}
    .badge-draft {{ background-color: #FEF3C7; color: #B45309; }}
    .badge-valide {{ background-color: #DCFCE7; color: #15803D; }}

    .transition-flow {{
        font-size: 0.85rem;
        color: {COLOR_TEXT_SECONDARY};
        text-align: center;
        line-height: 1.3;
    }}

    .alert-card {{
        background-color: #FEF2F2;
        border: 1px solid #FCA5A5;
        border-radius: 12px;
        padding: 0.9rem 1.1rem;
        margin: 0.6rem 0;
    }}
    .alert-title {{
        color: {COLOR_ERROR};
        font-weight: 700;
        font-size: 0.85rem;
        margin-bottom: 0.3rem;
    }}

    .db-indicator {{
        font-size: 0.78rem;
        color: {COLOR_SUCCESS};
        display: flex;
        align-items: center;
        gap: 0.35rem;
    }}
    .db-dot {{
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background-color: {COLOR_SUCCESS};
        display: inline-block;
    }}

    div[data-testid="stMetricValue"] {{
        color: {COLOR_TEXT};
    }}

    .stButton > button {{
        border-radius: 8px;
        border: 1px solid {COLOR_BORDER};
    }}
    .stButton > button[kind="primary"] {{
        background-color: {COLOR_PRIMARY};
        border-color: {COLOR_PRIMARY};
    }}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ==========================================================================
# INITIALISATION DES DONNÉES
# ==========================================================================

db.init_demo_data()

if "semaine" not in st.session_state:
    st.session_state.semaine = 36
if "annee" not in st.session_state:
    st.session_state.annee = 2026
if "planning_statut" not in st.session_state:
    st.session_state.planning_statut = "BROUILLON"
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"
if "proposition_optim" not in st.session_state:
    st.session_state.proposition_optim = None


# ==========================================================================
# HELPERS UI
# ==========================================================================

def badge_priorite(priorite: str) -> str:
    mapping = {
        "NORMAL": "badge-normal",
        "HAUTE": "badge-haute",
        "URGENTE": "badge-urgente",
        "CRITIQUE": "badge-critique",
    }
    cls = mapping.get(priorite, "badge-normal")
    return f'<span class="badge {cls}">{priorite}</span>'


def kpi_card(label: str, value: str):
    st.markdown(
        f"""
        <div class="alluco-card">
            <div class="alluco-kpi-label">{label}</div>
            <div class="alluco-kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def charge_color(pct: float) -> str:
    if pct >= 100:
        return COLOR_ERROR
    if pct >= 90:
        return COLOR_WARNING
    return COLOR_SUCCESS


def render_page_header(titre: str, sous_titre: str = ""):
    col_logo, col_title, col_meta = st.columns([1, 5, 3])
    with col_logo:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=48)
        else:
            st.markdown("### ALLUCO")
    with col_title:
        st.markdown(f"<div class='alluco-header-title'>{titre}</div>", unsafe_allow_html=True)
        if sous_titre:
            st.markdown(f"<div class='alluco-header-sub'>{sous_titre}</div>", unsafe_allow_html=True)
    with col_meta:
        m1, m2, m3 = st.columns([1, 1, 1])
        with m1:
            st.session_state.semaine = st.number_input(
                "Semaine", min_value=1, max_value=53, value=st.session_state.semaine, label_visibility="collapsed"
            )
        with m2:
            st.session_state.annee = st.number_input(
                "Année", min_value=2024, max_value=2030, value=st.session_state.annee, label_visibility="collapsed"
            )
        with m3:
            if st.button("🔄 Actualiser", use_container_width=True):
                st.rerun()
    st.markdown("<hr style='margin-top:0.2rem;margin-bottom:1.2rem;border-color:#E5E7EB;'>", unsafe_allow_html=True)


# ==========================================================================
# SIDEBAR
# ==========================================================================

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)
    else:
        st.markdown("## ALLUCO")

    st.markdown("**Planning Laquage**")
    st.write("")

    menu_items = {
        "🏠 Dashboard": "Dashboard",
        "📦 Commandes": "Commandes",
        "📅 Planning": "Planning",
        "⚡ Optimisation": "Optimisation",
        "🎨 Analyse couleurs": "Analyse couleurs",
        "📊 Analyse production": "Analyse production",
        "✅ Suivi": "Suivi",
        "🕘 Historique": "Historique",
        "⚙️ Paramètres": "Paramètres",
    }

    choix = st.radio(
        "Navigation",
        list(menu_items.keys()),
        index=list(menu_items.values()).index(st.session_state.page),
        label_visibility="collapsed",
    )
    st.session_state.page = menu_items[choix]

    st.markdown("---")
    st.markdown(
        f"""
        <div style="font-size:0.78rem;color:{COLOR_TEXT_SECONDARY};">Semaine actuelle</div>
        <div style="font-weight:700;margin-bottom:0.5rem;">S{st.session_state.semaine} - {st.session_state.annee}</div>
        <div class="db-indicator"><span class="db-dot"></span> Base de données connectée</div>
        """,
        unsafe_allow_html=True,
    )

page = st.session_state.page

# ==========================================================================
# PAGE : DASHBOARD
# ==========================================================================

if page == "Dashboard":
    render_page_header("Planning Laquage", f"Semaine {st.session_state.semaine} - Vue globale de production")

    commandes = db.get_commandes()
    planning = db.get_planning()
    couleurs_ref = db.get_couleurs()
    params = db.get_parametres()

    total_bal = sum(l.get("bal", 0) for j in db.JOURS_SEMAINE for l in planning.get(j, []))
    total_poids_kg = sum(l.get("poids_t", 0) * 1000 for j in db.JOURS_SEMAINE for l in planning.get(j, []))
    total_poudre = sum(l.get("poudre_kg", 0) for j in db.JOURS_SEMAINE for l in planning.get(j, []))
    total_pcs = sum(l.get("reste", 0) for j in db.JOURS_SEMAINE for l in planning.get(j, []))
    total_h = sum(db.calculer_charge_jour(planning.get(j, []))["charge_totale"] for j in db.JOURS_SEMAINE)

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        kpi_card("Commandes", str(len(commandes)))
    with c2:
        kpi_card("À laquer", f"{total_pcs:,} pcs".replace(",", " "))
    with c3:
        kpi_card("Poids", f"{total_poids_kg:,.0f} kg".replace(",", " "))
    with c4:
        kpi_card("Poudre", f"{total_poudre:,.0f} kg".replace(",", " "))
    with c5:
        kpi_card("Balancelles", f"{total_bal:,}".replace(",", " "))
    with c6:
        kpi_card("Charge", f"{total_h:.1f} h")

    urgentes = [c for c in commandes if c["priorite"] in ("URGENTE", "CRITIQUE")]
    retards = [c for c in commandes if c["statut"] == "BLOQUÉ"]
    nb_couleurs_semaine = len({l["couleur"] for j in db.JOURS_SEMAINE for l in planning.get(j, [])})
    taux_moyen = 0.0
    jours_avec_charge = [db.calculer_charge_jour(planning.get(j, [])) for j in db.JOURS_SEMAINE]
    if jours_avec_charge:
        taux_moyen = sum(
            (j["charge_totale"] / j["capacite"] * 100) if j["capacite"] else 0 for j in jours_avec_charge
        ) / len(jours_avec_charge)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Urgentes", str(len(urgentes)))
    with c2:
        kpi_card("Retards / Bloquées", str(len(retards)))
    with c3:
        kpi_card("Couleurs", str(nb_couleurs_semaine))
    with c4:
        kpi_card("Taux de charge moyen", f"{taux_moyen:.0f} %")

    st.markdown("#### Planning de la semaine")
    cols = st.columns(6)
    for i, jour in enumerate(db.JOURS_SEMAINE):
        lignes = planning.get(jour, [])
        charge = db.calculer_charge_jour(lignes)
        pct = (charge["charge_totale"] / charge["capacite"] * 100) if charge["capacite"] else 0
        pct = min(pct, 100)
        couleurs_jour = db.sequence_couleurs_jour(lignes)
        bal_jour = sum(l.get("bal", 0) for l in lignes)
        poids_jour = sum(l.get("poids_t", 0) for l in lignes) * 1000
        color = charge_color(pct)
        with cols[i]:
            st.markdown(
                f"""
                <div class="alluco-card" style="text-align:center;">
                    <div style="font-weight:700;font-size:0.82rem;letter-spacing:0.03em;">{jour}</div>
                    <div style="font-size:1.3rem;font-weight:700;margin:0.3rem 0;">{charge['charge_totale']:.2f} h</div>
                    <div style="background:#F1F2F4;border-radius:6px;height:8px;overflow:hidden;margin-bottom:0.3rem;">
                        <div style="background:{color};width:{pct:.0f}%;height:100%;"></div>
                    </div>
                    <div style="font-size:0.72rem;color:{COLOR_TEXT_SECONDARY};margin-bottom:0.4rem;">{pct:.0f} %</div>
                    <div style="font-size:0.75rem;color:{COLOR_TEXT_SECONDARY};">{' • '.join(couleurs_jour) if couleurs_jour else '—'}</div>
                    <div style="font-size:0.72rem;color:{COLOR_TEXT_SECONDARY};margin-top:0.4rem;">{bal_jour} balancelles</div>
                    <div style="font-size:0.72rem;color:{COLOR_TEXT_SECONDARY};">{poids_jour:,.0f} kg</div>
                </div>
                """.replace(",", " "),
                unsafe_allow_html=True,
            )

# ==========================================================================
# PAGE : COMMANDES
# ==========================================================================

elif page == "Commandes":
    render_page_header("Commandes à planifier", f"Semaine {st.session_state.semaine}")

    commandes = db.get_commandes()
    df = pd.DataFrame(commandes)

    search = st.text_input("🔎 Rechercher commande, client, article, OF...", "")

    fc1, fc2, fc3, fc4, fc5 = st.columns(5)
    with fc1:
        f_client = st.selectbox("Client", ["Tous"] + sorted(df["client"].unique().tolist()) if not df.empty else ["Tous"])
    with fc2:
        f_couleur = st.selectbox("Couleur", ["Toutes"] + sorted(df["couleur"].unique().tolist()) if not df.empty else ["Toutes"])
    with fc3:
        f_article = st.selectbox("Article", ["Tous"] + sorted(df["article"].unique().tolist()) if not df.empty else ["Tous"])
    with fc4:
        f_priorite = st.selectbox("Priorité", ["Toutes"] + db.PRIORITES)
    with fc5:
        f_statut = st.selectbox("Statut", ["Tous"] + db.STATUTS)

    df_filtre = df.copy()
    if not df_filtre.empty:
        if search:
            s = search.lower()
            df_filtre = df_filtre[
                df_filtre.apply(lambda r: s in str(r["commande"]).lower() or s in str(r["client"]).lower()
                                 or s in str(r["article"]).lower() or s in str(r["of"]).lower(), axis=1)
            ]
        if f_client != "Tous":
            df_filtre = df_filtre[df_filtre["client"] == f_client]
        if f_couleur != "Toutes":
            df_filtre = df_filtre[df_filtre["couleur"] == f_couleur]
        if f_article != "Tous":
            df_filtre = df_filtre[df_filtre["article"] == f_article]
        if f_priorite != "Toutes":
            df_filtre = df_filtre[df_filtre["priorite"] == f_priorite]
        if f_statut != "Tous":
            df_filtre = df_filtre[df_filtre["statut"] == f_statut]

    st.markdown(f"<div class='alluco-header-sub'>{len(df_filtre)} commande(s)</div>", unsafe_allow_html=True)

    if not df_filtre.empty:
        df_display = df_filtre.copy()
        df_display.insert(0, "✓", False)
        edited = st.data_editor(
            df_display[["✓", "commande", "client", "article", "couleur", "reste", "stock", "of", "priorite", "statut"]],
            column_config={
                "✓": st.column_config.CheckboxColumn(""),
                "commande": "Commande",
                "client": "Client",
                "article": "Article",
                "couleur": "Couleur",
                "reste": st.column_config.NumberColumn("Reste", format="%d"),
                "stock": st.column_config.NumberColumn("Stock", format="%d"),
                "of": "OF",
                "priorite": "Priorité",
                "statut": "Statut",
            },
            disabled=["commande", "client", "article", "couleur", "reste", "stock", "of", "priorite", "statut"],
            hide_index=True,
            use_container_width=True,
        )
        selection = edited[edited["✓"]]["commande"].tolist()
        st.button(
            f"➕ Ajouter au planning ({len(selection)})",
            type="primary",
            disabled=len(selection) == 0,
        )
    else:
        st.info("Aucune commande ne correspond à ces filtres.")

# ==========================================================================
# PAGE : PLANNING
# ==========================================================================

elif page == "Planning":
    statut_badge = "badge-valide" if st.session_state.planning_statut == "VALIDÉ" else "badge-draft"
    col_h, col_actions = st.columns([3, 2])
    with col_h:
        st.markdown(
            f"### Planning S{st.session_state.semaine} "
            f"<span class='badge {statut_badge}'>{st.session_state.planning_statut}</span>",
            unsafe_allow_html=True,
        )
    with col_actions:
        a1, a2, a3 = st.columns(3)
        with a1:
            if st.button("🔄 Actualiser", use_container_width=True):
                st.rerun()
        with a2:
            if st.button("⚡ Optimiser", use_container_width=True):
                st.session_state.page = "Optimisation"
                st.rerun()
        with a3:
            if st.button("✅ Valider planning", type="primary", use_container_width=True):
                st.session_state.planning_statut = "VALIDÉ"
                db.archiver_semaine(st.session_state.semaine, st.session_state.annee)
                st.success("Planning validé et archivé dans l'historique.")

    st.write("")

    tabs = st.tabs([j.capitalize() for j in db.JOURS_SEMAINE])
    planning = db.get_planning()

    for tab, jour in zip(tabs, db.JOURS_SEMAINE):
        with tab:
            lignes = planning.get(jour, [])
            charge = db.calculer_charge_jour(lignes)
            jour_date = date.today().strftime("%d/%m/%Y")

            st.markdown(
                f"""
                <div class="alluco-card">
                    <div style="font-weight:700;margin-bottom:0.5rem;">{jour} - {jour_date}</div>
                    <table style="width:100%;font-size:0.88rem;">
                        <tr><td style="color:{COLOR_TEXT_SECONDARY};">Charge production</td>
                            <td style="text-align:right;font-weight:600;">{charge['charge_production']:.2f} h</td></tr>
                        <tr><td style="color:{COLOR_TEXT_SECONDARY};">Changements couleur</td>
                            <td style="text-align:right;font-weight:600;">{charge['changements_couleur']:.2f} h</td></tr>
                        <tr><td style="color:{COLOR_TEXT_SECONDARY};">Charge totale</td>
                            <td style="text-align:right;font-weight:700;">{charge['charge_totale']:.2f} h</td></tr>
                        <tr><td style="color:{COLOR_TEXT_SECONDARY};">Capacité</td>
                            <td style="text-align:right;">{charge['capacite']:.2f} h</td></tr>
                    </table>
                </div>
                """,
                unsafe_allow_html=True,
            )

            sequence = db.sequence_couleurs_jour(lignes)
            if sequence:
                st.markdown(
                    "<div class='transition-flow'>" + " &nbsp;→&nbsp; ".join(sequence) + "</div>",
                    unsafe_allow_html=True,
                )
                for i in range(1, len(sequence)):
                    alerte = db.transition_alerte(sequence[i - 1], sequence[i])
                    if alerte:
                        suggestion = " → ".join([sequence[i - 1]] + alerte["suggestion"] + [sequence[i]]) \
                            if alerte["suggestion"] else f"{sequence[i-1]} → {sequence[i]}"
                        st.markdown(
                            f"""
                            <div class="alert-card">
                                <div class="alert-title">⚠ TRANSITION NON RECOMMANDÉE</div>
                                <div style="font-size:0.85rem;">{sequence[i-1]} → {sequence[i]}</div>
                                <div style="font-size:0.78rem;color:{COLOR_TEXT_SECONDARY};margin-top:0.2rem;">
                                    Écart de clarté important.
                                </div>
                                <div style="font-size:0.78rem;margin-top:0.3rem;">Suggestion : {suggestion}</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

            st.write("")

            # Groupement par couleur
            couleurs_jour = db.sequence_couleurs_jour(lignes)
            for idx_c, couleur in enumerate(couleurs_jour):
                lignes_couleur = [l for l in lignes if l["couleur"] == couleur]
                info = db.get_couleur_info(couleur) or {"famille": "-"}
                with st.expander(f"{couleur}  ·  {info.get('famille', '-')}", expanded=(idx_c == 0)):
                    df_c = pd.DataFrame(lignes_couleur)
                    if not df_c.empty:
                        df_show = df_c[[
                            "ordre", "commande", "client", "article", "reste", "lancement",
                            "re_laquage", "poids_t", "poudre_kg", "barre_par_bal", "bal", "temps_h",
                        ]].rename(columns={
                            "ordre": "Ordre", "commande": "Commande", "client": "Client", "article": "Article",
                            "reste": "Reste", "lancement": "Lancement", "re_laquage": "Re-laquage",
                            "poids_t": "Poids T", "poudre_kg": "Poudre", "barre_par_bal": "Barre/Bal",
                            "bal": "Bal", "temps_h": "Temps",
                        })
                        st.dataframe(df_show, hide_index=True, use_container_width=True)

                        tot_pcs = sum(l["reste"] for l in lignes_couleur)
                        tot_kg = sum(l["poids_t"] for l in lignes_couleur) * 1000
                        tot_poudre = sum(l["poudre_kg"] for l in lignes_couleur)
                        tot_bal = sum(l["bal"] for l in lignes_couleur)
                        tot_h = sum(l["temps_h"] for l in lignes_couleur)
                        st.markdown(
                            f"""
                            <div style="font-size:0.82rem;color:{COLOR_TEXT_SECONDARY};margin-top:0.4rem;">
                                Total {couleur} — {tot_pcs:,} pcs · {tot_kg:,.0f} kg · {tot_poudre:.1f} kg poudre
                                · {tot_bal} balancelles · {tot_h:.2f} h
                            </div>
                            """.replace(",", " "),
                            unsafe_allow_html=True,
                        )

                    if idx_c < len(couleurs_jour) - 1:
                        suivante = couleurs_jour[idx_c + 1]
                        t_info = db.get_transition_info(couleur, suivante)
                        st.markdown(
                            f"<div class='transition-flow' style='margin-top:0.6rem;'>{couleur}<br>↓ {t_info['temps_nettoyage']} min ↓<br>{suivante}</div>",
                            unsafe_allow_html=True,
                        )

            # Mode manuel
            with st.expander("🔧 Mode manuel — modifier les lignes"):
                for l in lignes:
                    cols = st.columns([2, 2, 2, 2, 1, 1])
                    cols[0].write(f"**{l['commande']}** — {l['client']}")
                    cols[1].write(l["couleur"])
                    nouveau_jour = cols[2].selectbox(
                        "Jour", db.JOURS_SEMAINE, index=db.JOURS_SEMAINE.index(jour),
                        key=f"jour_{jour}_{l['commande_id']}", label_visibility="collapsed",
                    )
                    nouvel_ordre = cols[3].number_input(
                        "Ordre", min_value=1, value=l.get("ordre", 1),
                        key=f"ordre_{jour}_{l['commande_id']}", label_visibility="collapsed",
                    )
                    verrou = cols[4].checkbox("🔒", value=l.get("verrouille", False), key=f"lock_{jour}_{l['commande_id']}")
                    supprimer = cols[5].button("🗑️", key=f"del_{jour}_{l['commande_id']}")

                    if supprimer:
                        lignes = [x for x in lignes if x["commande_id"] != l["commande_id"]]
                        db.save_planning_jour(jour, lignes)
                        st.rerun()
                    elif nouveau_jour != jour or nouvel_ordre != l.get("ordre") or verrou != l.get("verrouille"):
                        if nouveau_jour != jour:
                            lignes = [x for x in lignes if x["commande_id"] != l["commande_id"]]
                            db.save_planning_jour(jour, lignes)
                            l["ordre"] = nouvel_ordre
                            l["verrouille"] = verrou
                            autre = db.get_planning_jour(nouveau_jour)
                            autre.append(l)
                            db.save_planning_jour(nouveau_jour, autre)
                        else:
                            l["ordre"] = nouvel_ordre
                            l["verrouille"] = verrou
                            db.save_planning_jour(jour, lignes)
                        st.rerun()

# ==========================================================================
# PAGE : OPTIMISATION
# ==========================================================================

elif page == "Optimisation":
    render_page_header("Optimisation couleurs", f"Semaine {st.session_state.semaine}")

    st.markdown("#### Optimisation couleurs")

    jour_choisi = st.selectbox("Journée à optimiser", db.JOURS_SEMAINE, key="jour_optim")
    lignes = db.get_planning_jour(jour_choisi)
    sequence_actuelle = db.sequence_couleurs_jour(lignes)

    st.markdown("##### Couleur actuelle")
    st.write(sequence_actuelle[0] if sequence_actuelle else "—")

    st.markdown("##### Couleurs à produire")
    st.write(" · ".join(sequence_actuelle) if sequence_actuelle else "—")

    sequence_optimisee = db.optimiser_ordre_couleurs(sequence_actuelle)

    st.markdown("##### Ordre recommandé")
    st.markdown(
        "<div class='transition-flow'>" + " ↓<br>".join(sequence_optimisee) + "</div>",
        unsafe_allow_html=True,
    )

    st.write("")
    st.markdown("#### Comparaison AVANT / PROPOSITION")

    col_a, col_b = st.columns(2)
    cout_actuel = db.cout_total_sequence(sequence_actuelle)
    cout_optim = db.cout_total_sequence(sequence_optimisee)

    with col_a:
        st.markdown("**Planning actuel**")
        st.markdown(
            f"<div class='alluco-card'>{'<br>'.join(sequence_actuelle) if sequence_actuelle else '—'}</div>",
            unsafe_allow_html=True,
        )
        kpi_card("Coût transition", str(cout_actuel["cout"]))
        h, m = divmod(cout_actuel["temps_min"], 60)
        kpi_card("Temps nettoyage", f"{h} h {m:02d}" if h else f"{m} min")

    with col_b:
        st.markdown("**Planning optimisé**")
        st.markdown(
            f"<div class='alluco-card'>{'<br>'.join(sequence_optimisee) if sequence_optimisee else '—'}</div>",
            unsafe_allow_html=True,
        )
        kpi_card("Coût transition", str(cout_optim["cout"]))
        h, m = divmod(cout_optim["temps_min"], 60)
        kpi_card("Temps nettoyage", f"{h} h {m:02d}" if h else f"{m} min")

    delta_cout = cout_optim["cout"] - cout_actuel["cout"]
    delta_temps = cout_optim["temps_min"] - cout_actuel["temps_min"]
    dh, dm = divmod(abs(delta_temps), 60)
    st.markdown(
        f"""
        <div class="alluco-card" style="text-align:center;">
            <span style="color:{COLOR_SUCCESS if delta_cout <= 0 else COLOR_ERROR};font-weight:700;">
                {delta_cout:+d} coût transition
            </span>
            &nbsp;&nbsp;
            <span style="color:{COLOR_SUCCESS if delta_temps <= 0 else COLOR_ERROR};font-weight:700;">
                {'-' if delta_temps <= 0 else '+'}{dh} h {dm:02d} nettoyage
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    b1, b2 = st.columns(2)
    with b1:
        if st.button("✅ Appliquer optimisation", type="primary", use_container_width=True):
            ordre_map = {c: i for i, c in enumerate(sequence_optimisee)}
            nouvelles_lignes = sorted(
                lignes,
                key=lambda l: (ordre_map.get(l["couleur"], 999), l.get("ordre", 0)),
            )
            for i, l in enumerate(nouvelles_lignes, start=1):
                if not l.get("verrouille"):
                    l["ordre"] = i
            db.save_planning_jour(jour_choisi, nouvelles_lignes)
            st.success(f"Optimisation appliquée au planning de {jour_choisi}.")
    with b2:
        st.button("❌ Annuler", use_container_width=True)

# ==========================================================================
# PAGE : ANALYSE COULEURS
# ==========================================================================

elif page == "Analyse couleurs":
    render_page_header("Analyse couleurs", f"Semaine {st.session_state.semaine}")

    planning = db.get_planning()
    toutes_lignes = [l for j in db.JOURS_SEMAINE for l in planning.get(j, [])]
    couleurs_semaine = {l["couleur"] for l in toutes_lignes}

    nb_changements = sum(
        len(db.sequence_couleurs_jour(planning.get(j, []))) - 1
        for j in db.JOURS_SEMAINE if len(db.sequence_couleurs_jour(planning.get(j, []))) > 0
    )

    nb_optimales, nb_surveiller, nb_mauvaises = 0, 0, 0
    for j in db.JOURS_SEMAINE:
        seq = db.sequence_couleurs_jour(planning.get(j, []))
        for i in range(1, len(seq)):
            ecart = abs(db.get_clarte(seq[i - 1]) - db.get_clarte(seq[i]))
            if ecart <= 1:
                nb_optimales += 1
            elif ecart == 2:
                nb_surveiller += 1
            else:
                nb_mauvaises += 1

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        kpi_card("Couleurs cette semaine", str(len(couleurs_semaine)))
    with c2:
        kpi_card("Changements", str(nb_changements))
    with c3:
        kpi_card("Transitions optimales", str(nb_optimales))
    with c4:
        kpi_card("À surveiller", str(nb_surveiller))
    with c5:
        kpi_card("Mauvaises", str(nb_mauvaises))

    st.write("")
    st.markdown("#### Répartition par couleur")

    stats = {}
    for l in toutes_lignes:
        c = l["couleur"]
        s = stats.setdefault(c, {"qte": 0, "poids": 0.0, "poudre": 0.0, "bal": 0, "temps": 0.0})
        s["qte"] += l.get("reste", 0)
        s["poids"] += l.get("poids_t", 0) * 1000
        s["poudre"] += l.get("poudre_kg", 0)
        s["bal"] += l.get("bal", 0)
        s["temps"] += l.get("temps_h", 0)

    if stats:
        max_qte = max(s["qte"] for s in stats.values()) or 1
        for couleur, s in sorted(stats.items(), key=lambda x: -x[1]["qte"]):
            pct = s["qte"] / max_qte * 100
            col_label, col_bar = st.columns([1, 4])
            col_label.write(couleur)
            col_bar.markdown(
                f"""
                <div style="background:#F1F2F4;border-radius:6px;height:18px;overflow:hidden;">
                    <div style="background:{COLOR_PRIMARY};width:{pct:.0f}%;height:100%;
                        display:flex;align-items:center;padding-left:6px;color:white;font-size:0.72rem;">
                        {s['qte']:,}
                    </div>
                </div>
                """.replace(",", " "),
                unsafe_allow_html=True,
            )

        df_stats = pd.DataFrame([
            {"Couleur": c, "Qte": s["qte"], "Poids": round(s["poids"]), "Poudre": round(s["poudre"], 1),
             "Balancelles": s["bal"], "Temps": round(s["temps"], 2)}
            for c, s in stats.items()
        ]).sort_values("Qte", ascending=False)
        st.dataframe(df_stats, hide_index=True, use_container_width=True)
    else:
        st.info("Aucune donnée de planning pour cette semaine.")

# ==========================================================================
# PAGE : ANALYSE PRODUCTION
# ==========================================================================

elif page == "Analyse production":
    render_page_header("Analyse production", f"Semaine {st.session_state.semaine}")

    planning = db.get_planning()

    st.markdown("#### Charge par jour")
    df_charge = pd.DataFrame([
        {"Jour": j[:3], "Charge (h)": db.calculer_charge_jour(planning.get(j, []))["charge_totale"]}
        for j in db.JOURS_SEMAINE
    ]).set_index("Jour")
    st.bar_chart(df_charge)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Quantité par jour")
        df_qte = pd.DataFrame([
            {"Jour": j[:3], "Pièces": sum(l.get("reste", 0) for l in planning.get(j, []))}
            for j in db.JOURS_SEMAINE
        ]).set_index("Jour")
        st.bar_chart(df_qte)
    with col2:
        st.markdown("#### Poudre nécessaire")
        df_poudre = pd.DataFrame([
            {"Jour": j[:3], "Poudre (kg)": sum(l.get("poudre_kg", 0) for l in planning.get(j, []))}
            for j in db.JOURS_SEMAINE
        ]).set_index("Jour")
        st.bar_chart(df_poudre)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("#### Nombre de balancelles")
        df_bal = pd.DataFrame([
            {"Jour": j[:3], "Balancelles": sum(l.get("bal", 0) for l in planning.get(j, []))}
            for j in db.JOURS_SEMAINE
        ]).set_index("Jour")
        st.bar_chart(df_bal)
    with col4:
        st.markdown("#### Surcharge / retards")
        params = db.get_parametres()
        surcharges = [j for j in db.JOURS_SEMAINE
                      if db.calculer_charge_jour(planning.get(j, []))["charge_totale"] > params["capacite_quotidienne_h"]]
        commandes = db.get_commandes()
        retards = [c for c in commandes if c["statut"] == "BLOQUÉ"]
        kpi_card("Jours en surcharge", str(len(surcharges)))
        kpi_card("Commandes en retard", str(len(retards)))

# ==========================================================================
# PAGE : SUIVI
# ==========================================================================

elif page == "Suivi":
    render_page_header("Suivi", f"Semaine {st.session_state.semaine}")

    planning = db.get_planning()
    lignes_toutes = []
    for j in db.JOURS_SEMAINE:
        for l in planning.get(j, []):
            cmd = db.get_commande(l["commande_id"]) or {}
            lignes_toutes.append({
                "Commande": l["commande"],
                "Article": l["article"],
                "Couleur": l["couleur"],
                "Qte": l["reste"],
                "État": cmd.get("statut", "À PLANIFIER"),
                "_id": l["commande_id"],
            })

    if lignes_toutes:
        df_suivi = pd.DataFrame(lignes_toutes)
        edited = st.data_editor(
            df_suivi,
            column_config={
                "État": st.column_config.SelectboxColumn("État", options=db.STATUTS),
                "_id": None,
            },
            disabled=["Commande", "Article", "Couleur", "Qte"],
            hide_index=True,
            use_container_width=True,
            key="suivi_editor",
        )
        for _, row in edited.iterrows():
            db.update_commande_statut(row["_id"], row["État"])
    else:
        st.info("Aucune opération planifiée pour cette semaine.")

# ==========================================================================
# PAGE : HISTORIQUE
# ==========================================================================

elif page == "Historique":
    render_page_header("Historique planning", "")

    historique = db.get_historique()

    f1, f2, f3, f4, f5, f6 = st.columns(6)
    with f1:
        annees = ["Toutes"] + sorted({str(h["annee"]) for h in historique}, reverse=True)
        f_annee = st.selectbox("Année", annees)
    with f2:
        semaines = ["Toutes"] + sorted({str(h["semaine"]) for h in historique}, reverse=True)
        f_semaine = st.selectbox("Semaine", semaines)
    with f3:
        st.selectbox("Jour", ["Tous"])
    with f4:
        st.selectbox("Client", ["Tous"])
    with f5:
        st.selectbox("Article", ["Tous"])
    with f6:
        st.selectbox("Couleur", ["Toutes"])

    st.write("")
    for h in historique:
        if f_annee != "Toutes" and str(h["annee"]) != f_annee:
            continue
        if f_semaine != "Toutes" and str(h["semaine"]) != f_semaine:
            continue
        badge_cls = "badge-valide" if h["statut"] == "VALIDÉ" else "badge-draft"
        col_info, col_btn = st.columns([5, 1])
        with col_info:
            st.markdown(
                f"""
                <div class="alluco-card">
                    <b>S{h['semaine']} - {h['annee']}</b>
                    &nbsp; <span class="badge {badge_cls}">{h['statut']}</span><br>
                    <span style="color:{COLOR_TEXT_SECONDARY};font-size:0.85rem;">
                        {h['heures']} heures · {h['nb_couleurs']} couleurs · {h['nb_commandes']} commandes
                    </span>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with col_btn:
            st.write("")
            st.button("Ouvrir", key=f"open_{h['semaine']}_{h['annee']}")

# ==========================================================================
# PAGE : PARAMÈTRES
# ==========================================================================

elif page == "Paramètres":
    render_page_header("Paramètres", "")

    tab_prod, tab_couleurs, tab_transitions, tab_planning = st.tabs(
        ["Production", "Couleurs", "Transitions", "Planning"]
    )

    with tab_prod:
        params = db.get_parametres()
        capacite = st.number_input("Capacité quotidienne (h)", value=float(params["capacite_quotidienne_h"]), step=0.5)
        temps_bal = st.number_input("Temps / balancelle (min)", value=float(params["temps_par_balancelle_min"]), step=0.5)
        coeff_poudre = st.number_input(
            "Coefficient poudre (kg poudre / kg pièce)", value=float(params["coefficient_poudre_kg_par_kg"]),
            step=0.001, format="%.3f",
        )
        if st.button("💾 Enregistrer", key="save_prod"):
            params.update({
                "capacite_quotidienne_h": capacite,
                "temps_par_balancelle_min": temps_bal,
                "coefficient_poudre_kg_par_kg": coeff_poudre,
            })
            db.save_parametres(params)
            st.success("Paramètres de production enregistrés.")

    with tab_couleurs:
        couleurs = db.get_couleurs()
        df_couleurs = pd.DataFrame(couleurs).rename(columns={
            "couleur": "Couleur", "ral": "RAL", "famille": "Famille", "clarte": "Clarté",
        })
        edited = st.data_editor(
            df_couleurs,
            num_rows="dynamic",
            hide_index=True,
            use_container_width=True,
            column_config={
                "Clarté": st.column_config.NumberColumn("Clarté", min_value=1, max_value=5, step=1),
            },
        )
        if st.button("➕ Ajouter couleur / 💾 Enregistrer", key="save_couleurs"):
            nouvelles = edited.rename(columns={
                "Couleur": "couleur", "RAL": "ral", "Famille": "famille", "Clarté": "clarte",
            }).to_dict("records")
            db.save_couleurs(nouvelles)
            st.success("Référentiel couleurs mis à jour.")

    with tab_transitions:
        transitions = db.get_transitions()
        couleurs_noms = [c["couleur"] for c in db.get_couleurs()]

        if transitions:
            df_trans = pd.DataFrame(transitions).rename(columns={
                "depart": "Couleur départ", "arrivee": "Couleur arrivée",
                "cout": "Coût", "temps_nettoyage": "Temps nettoyage (min)",
            })
        else:
            df_trans = pd.DataFrame(columns=["Couleur départ", "Couleur arrivée", "Coût", "Temps nettoyage (min)"])

        edited_trans = st.data_editor(
            df_trans,
            num_rows="dynamic",
            hide_index=True,
            use_container_width=True,
            column_config={
                "Couleur départ": st.column_config.SelectboxColumn(options=couleurs_noms),
                "Couleur arrivée": st.column_config.SelectboxColumn(options=couleurs_noms),
            },
        )
        st.caption("Les transitions non définies ici sont calculées automatiquement à partir de l'écart de clarté.")
        if st.button("💾 Enregistrer", key="save_transitions"):
            nouvelles = edited_trans.rename(columns={
                "Couleur départ": "depart", "Couleur arrivée": "arrivee",
                "Coût": "cout", "Temps nettoyage (min)": "temps_nettoyage",
            }).dropna(subset=["depart", "arrivee"]).to_dict("records")
            db.save_transitions(nouvelles)
            st.success("Transitions enregistrées.")

    with tab_planning:
        params = db.get_parametres()
        jours_travailles = st.multiselect("Jours travaillés", db.JOURS_SEMAINE, default=params["jours_travailles"])
        h1, h2 = st.columns(2)
        with h1:
            horaire_debut = st.text_input("Horaire début", value=params["horaire_debut"])
        with h2:
            horaire_fin = st.text_input("Horaire fin", value=params["horaire_fin"])
        if st.button("💾 Enregistrer", key="save_planning_params"):
            params.update({
                "jours_travailles": jours_travailles,
                "horaire_debut": horaire_debut,
                "horaire_fin": horaire_fin,
            })
            db.save_parametres(params)
            st.success("Paramètres planning enregistrés.")
