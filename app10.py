            st.info("Aucune Base active n'a encore été chargée par l'administrateur.")
            return
        _client_portal(source)
        return

    if not _admin_gate():
        return

    with st.sidebar:
        if st.button("Se déconnecter"):
            st.session_state["admin_ok"] = False; st.rerun()
        st.divider()
        uploaded = st.file_uploader("Base AX / extraction v0", type=["xlsx"])
        if uploaded is not None:
            data = uploaded.getvalue()
            if not active_data or sha256_bytes(data) != sha256_bytes(active_data):
                blob_set("active_source", data); kv_set("active_source_name", uploaded.name)
                st.session_state.pop("last_result", None)
                st.success("Base activée.")
                active_data = data
                source = load_source_workbook(data)
        with st.expander("Référentiel Base.xlsx (optionnel)"):
            ref_upload = st.file_uploader("Mettre à jour le référentiel", type=["xlsx"], key="ref")
            if ref_upload is not None and st.button("Activer ce référentiel"):
                ref = reference_from_excel(ref_upload.getvalue(), label=ref_upload.name)
                if ref is None:
                    st.error("Référentiel non reconnu.")
                else:
                    kv_set("reference_override", {"version": ref.version, "colors": ref.colors, "articles": ref.articles, "stock": ref.stock})
                    st.success(f"Référentiel activé: {len(ref.articles)} articles / {len(ref.colors)} couleurs.")
                    st.rerun()

    if source is None:
        st.title("Planning IA")
        st.info("Charge une Base AX / extraction v0 dans la barre latérale.")
        return

    cfg = _config_ui(_default_cfg())
    master = load_reference()
    source_name = kv_get("active_source_name", "Base active")
    st.markdown(f"### Planning IA · `{source_name}`")
    st.caption(f"Feuille détectée: {source.attrs.get('source_sheet','—')} · mode {source.attrs.get('source_mode','—')} · référentiel {master.version}")

    c1, c2 = st.columns([1, 4])
    with c1:
        generate = st.button("Générer / Régénérer", type="primary", use_container_width=True)
    if generate or "last_result" not in st.session_state:
        with st.spinner("Agents: lecture → préparation → quantités → stock → campagnes → audit..."):
            try:
                st.session_state["last_result"] = generate_agentic_plan(source, cfg, master)
            except Exception as exc:
                st.error(f"Génération impossible: {exc}")
                return
    result = st.session_state["last_result"]
    # Si paramètres changent, recalcul explicite via le bouton, ce qui évite les reruns coûteux.
    _render_dashboard(result)

    st.divider()
    c1, c2, c3 = st.columns(3)
    excel = export_planning_excel(result)
    c1.download_button("Télécharger Excel", excel, file_name=f"Planning_IA_S{result['config'].week}_{result['config'].year}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    if REPORTLAB_AVAILABLE:
        pdf = export_planning_pdf(result)
        c2.download_button("Télécharger PDF", pdf, file_name=f"Planning_IA_S{result['config'].week}_{result['config'].year}.pdf", mime="application/pdf", use_container_width=True)
    if not result["hard_errors"]:
        if c3.button("Publier au portail client", type="primary", use_container_width=True):
            kv_set("published_plan", published_payload(result, source.attrs.get("source_sha256", "")))
            st.success("Planning publié.")
    else:
        c3.warning("Publication bloquée par l'audit.")

    tabs = st.tabs(["Analyse agents", "Audit", "Exclusions", "Décisions"])
    with tabs[0]:
        for agent, status, msg in result["steps"]:
            st.markdown(f"**{agent} — {status}**  \n{msg}")
        st.caption(f"Temps calcul: {result['elapsed_s']:.3f} s")
    with tabs[1]:
        st.dataframe(result["audit"], hide_index=True, use_container_width=True)
    with tabs[2]:
        if result["excluded"].empty:
            st.success("Aucune exclusion.")
        else:
            st.dataframe(result["excluded"], hide_index=True, use_container_width=True, height=500)
            counts = result["excluded"]["reason_code"].value_counts().rename_axis("reason_code").reset_index(name="lignes")
            st.dataframe(counts, hide_index=True, use_container_width=True)
    with tabs[3]:
        st.dataframe(explanation_table(result), hide_index=True, use_container_width=True, height=500)


# =============================================================================
# 14. CLI / TESTS
# =============================================================================
def cli_generate(source_path: str, output_path: str, year: Optional[int] = None, week: Optional[int] = None, reference_path: Optional[str] = None) -> Dict[str, Any]:
    data = Path(source_path).read_bytes()
    source = load_source_workbook(data)
    y, w, _, _ = next_planning_period()
    if year is not None:
        y = int(year)
    if week is not None:
        w = int(week)
    cfg = PlannerConfig(year=y, week=w)
    master = embedded_reference()
    if reference_path:
        override = reference_from_excel(Path(reference_path).read_bytes(), label=Path(reference_path).name)
        master = merge_reference(master, override)
    result = generate_agentic_plan(source, cfg, master)
    Path(output_path).write_bytes(export_planning_excel(result))
    return result


def self_test() -> None:
    ref = embedded_reference()
    assert len(ref.colors) >= 50
    assert len(ref.articles) >= 400
    assert abs(DEFAULT_MIN_PER_BAL - 4.0) < 1e-9
    assert len(OUTPUT_COLUMNS) == 31
    y, w, start, end = next_planning_period(date(2026, 10, 2))
    assert (y, w, start, end) == (2026, 41, date(2026, 10, 5), date(2026, 10, 9))
    print("[OK] self-test", APP_VERSION, len(ref.articles), "articles", len(ref.colors), "couleurs")


def main() -> None:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--generate", action="store_true")
    parser.add_argument("--input")
    parser.add_argument("--output")
    parser.add_argument("--reference")
    parser.add_argument("--year", type=int)
    parser.add_argument("--week", type=int)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--hash-password", action="store_true")
    args, _ = parser.parse_known_args()
    if args.hash_password:
        import getpass
        pwd = getpass.getpass("Mot de passe: ")
        print(make_password_hash(pwd)); return
    if args.self_test:
        self_test(); return
    if args.generate:
        if not args.input or not args.output:
            raise SystemExit("--input et --output sont obligatoires")
        r = cli_generate(args.input, args.output, args.year, args.week, args.reference)
        print("Planning généré:", args.output)
        print("Validation moteur:", r["validation_pct"], "%")
        print("Lignes préparées:", len(r["prepared"]), "Backlog:", len(r["backlog"]), "Report S+1:", len(r["report"]))
        return
    if st is None:
        print("Lancez avec: streamlit run app.py")
        print("Dépendances: pip install streamlit pandas openpyxl numpy reportlab")
        return
    render_ui()


if __name__ == "__main__":
    main()
