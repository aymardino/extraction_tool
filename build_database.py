"""
build_database.py — Reconstruit la base relationnelle de l'observatoire AISESA
à partir du classeur Excel (source de saisie).

Usage:
    python build_database.py Models_Africa_v2.xlsm enermod.db

C'est aussi le workflow "remplacer le SQL" : tu édites dans Excel, tu relances
ce script, la base est régénérée proprement (UTF-8, clés primaires/étrangères).

Pour passer à Postgres (Supabase) plus tard : remplace la connexion sqlite3 par
une connexion SQLAlchemy (create_engine("postgresql+psycopg2://...")) et utilise
df.to_sql(...) — le reste du script est identique.
"""
import sys, re, sqlite3
import pandas as pd

SHEETS = ['studies.csv', 'study_tools', 'study_countries', 'study_pools',
          'countries.csv', 'power_pools.csv', 'tools.csv']

# Table SQL <- feuille Excel
TABLE = {'studies.csv': 'studies', 'study_tools': 'study_tools',
         'study_countries': 'study_countries', 'study_pools': 'study_pools',
         'countries.csv': 'countries', 'power_pools.csv': 'power_pools',
         'tools.csv': 'tools'}

# Corrections de noms de colonnes (fautes + espaces -> snake_case)
FIX = {'strenghts': 'strengths', 'grey_litterature': 'grey_literature',
       'tool_category': 'type',  # nouveau nom -> nom DB historique
       'type of study': 'study_type', 'financing mechanism': 'financing_mechanism',
       'authors affiliation': 'authors_affiliation',
       'mathematical approach': 'mathematical_approach', 'key result': 'key_result'}


def clean_col(c):
    c = str(c).strip().lower()
    c = FIX.get(c, c)
    return re.sub(r'[^\w]+', '_', c).strip('_')


def load(path):
    # keep_default_na=False : sinon "NA" (code ISO de la Namibie) devient NaN.
    # On lit toutes les feuilles puis on aliase 'studies' -> 'studies.csv' pour
    # accepter le format de l'extraction tool (qui produit 'studies' sans .csv).
    xls = pd.read_excel(path, sheet_name=None, dtype=str, engine='openpyxl',
                        keep_default_na=False, na_values=[])
    # Alias : si l'export produit 'studies' au lieu de 'studies.csv'
    if 'studies' in xls and 'studies.csv' not in xls:
        xls['studies.csv'] = xls.pop('studies')
    # Vérifier qu'on a toutes les feuilles attendues
    missing = [s for s in SHEETS if s not in xls]
    if missing:
        raise SystemExit(f"Feuilles manquantes dans {path}: {missing}\n"
                         f"Trouvées: {list(xls.keys())}")
    # Ne garder que les feuilles qu'on connaît
    xls = {k: v for k, v in xls.items() if k in SHEETS}
    out = {}
    for name, df in xls.items():
        df.columns = [clean_col(c) for c in df.columns]
        df = df.loc[:, [c for c in df.columns if c and not c.startswith('unnamed')]]
        for col in df.columns:                       # trim des cellules texte
            df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
        # lignes entièrement vides (toutes cellules == "") -> supprimées
        df = df[~df.apply(lambda r: all(str(v).strip() == '' for v in r), axis=1)]
        out[name] = df
    # Tables maîtres : supprimer les lignes sans clé primaire
    for sheet, pk in [('countries.csv', 'iso_code'), ('tools.csv', 'tool_name'),
                      ('power_pools.csv', 'pool_code')]:
        out[sheet] = out[sheet][out[sheet][pk].astype(str).str.strip() != '']
    # study_id en entier (PK propre)
    for s in ['studies.csv', 'study_tools', 'study_countries', 'study_pools']:
        out[s]['study_id'] = pd.to_numeric(out[s]['study_id'], errors='coerce').astype('Int64')
        out[s] = out[s].dropna(subset=['study_id'])
    out['studies.csv']['year'] = pd.to_numeric(out['studies.csv']['year'], errors='coerce').astype('Int64')
    return out


# Définition du schéma : clés primaires + clés étrangères (déclarées, non bloquantes)
DDL = """
CREATE TABLE countries (
    iso_code TEXT PRIMARY KEY, iso3 TEXT, country_name TEXT, power_pool TEXT,
    region TEXT, language TEXT, nb_models_applied INTEGER, nb_models_national INTEGER,
    has_institutional_capacity TEXT, data_availability TEXT,
    electrification_rate REAL, has_ndc TEXT, has_lts TEXT
);
CREATE TABLE power_pools (
    pool_code TEXT PRIMARY KEY, pool_name TEXT, member_countries TEXT,
    nb_models_applied INTEGER
);
CREATE TABLE tools (
    tool_name TEXT PRIMARY KEY, full_name TEXT, type TEXT, origin_country TEXT,
    origin_institution TEXT, license TEXT, cost_usd TEXT, free_for_developing TEXT,
    learning_curve TEXT, doc_quality TEXT, training_available TEXT, community TEXT,
    programming_required TEXT, helpdesk TEXT, os TEXT, nb_studies_in_inventory INTEGER,
    best_for TEXT, clean_cooking_capable TEXT, financing_modelling TEXT,
    strengths TEXT, weaknesses TEXT, typical_results TEXT, often_linked TEXT
);
CREATE TABLE studies (
    study_id INTEGER PRIMARY KEY, authors TEXT, model_name TEXT, extraction_level TEXT,
    study_category TEXT, year INTEGER, study_type TEXT, scale TEXT, countries TEXT,
    nb_countries_covered TEXT, power_pool TEXT, study_objective TEXT, key_result TEXT,
    time_horizon TEXT, time_horizon_start TEXT, time_horizon_end TEXT, approach TEXT,
    method TEXT, mathematical_approach TEXT, hydro TEXT, solar TEXT, wind TEXT,
    biomass TEXT, nuclear TEXT, geothermal TEXT, fossil TEXT, h2 TEXT, coal TEXT,
    other_technologies TEXT, sector TEXT, open_source TEXT, data_requirements TEXT,
    frequency_of_use TEXT, sdg_7 TEXT, sdg_13 TEXT, ndc_mention TEXT, agenda_2063 TEXT,
    aisesa_theme TEXT, informal_economy TEXT, biomass_charcoal TEXT, clean_cooking TEXT,
    power_reliability TEXT, urbanization TEXT, strengths TEXT, weaknesses TEXT,
    authors_affiliation TEXT, author_origin TEXT, local_ownership TEXT,
    institutional_users TEXT, grey_literature TEXT, cost_of_capital TEXT,
    financing_modelling TEXT, financing_mechanism TEXT, full_title TEXT,
    link_doi TEXT, contact TEXT, other_references TEXT
);
CREATE TABLE study_tools (
    study_id INTEGER, tool_name TEXT, type TEXT,
    PRIMARY KEY (study_id, tool_name),
    FOREIGN KEY (study_id) REFERENCES studies(study_id)
);
CREATE TABLE study_countries (
    study_id INTEGER, iso_code TEXT, scale TEXT, iso3 TEXT, country TEXT,
    PRIMARY KEY (study_id, iso_code),
    FOREIGN KEY (study_id) REFERENCES studies(study_id),
    FOREIGN KEY (iso_code) REFERENCES countries(iso_code)
);
CREATE TABLE study_pools (
    study_id INTEGER, pool_code TEXT,
    PRIMARY KEY (study_id, pool_code),
    FOREIGN KEY (study_id) REFERENCES studies(study_id)
);
"""


def build(data, db_path):
    con = sqlite3.connect(db_path)
    con.executescript("DROP TABLE IF EXISTS study_tools; DROP TABLE IF EXISTS study_countries;"
                      "DROP TABLE IF EXISTS study_pools; DROP TABLE IF EXISTS studies;"
                      "DROP TABLE IF EXISTS tools; DROP TABLE IF EXISTS power_pools;"
                      "DROP TABLE IF EXISTS countries;")
    con.executescript(DDL)
    # Ordre d'insertion : parents avant enfants
    order = ['countries.csv', 'power_pools.csv', 'tools.csv', 'studies.csv',
             'study_tools', 'study_countries', 'study_pools']
    for name in order:
        df = data[name]
        cols = pd.read_sql(f"PRAGMA table_info({TABLE[name]})", con)['name'].tolist()
        df = df[[c for c in df.columns if c in cols]]   # garder colonnes connues du schéma
        df.to_sql(TABLE[name], con, if_exists='append', index=False)
    con.commit()
    return con


def integrity_report(con):
    """Rapporte les violations de clés étrangères sans bloquer (données vivantes)."""
    rows = con.execute("PRAGMA foreign_key_check").fetchall()
    if not rows:
        print("  Intégrité référentielle : OK (aucun orphelin)")
        return
    from collections import Counter
    by_table = Counter(r[0] for r in rows)
    print("  ⚠ Orphelins détectés (FK déclarées mais non bloquantes) :")
    for tbl, n in by_table.items():
        print(f"     - {tbl}: {n} ligne(s) sans parent")
    # Détail des outils orphelins (le cas le plus fréquent)
    orphan_tools = con.execute(
        "SELECT DISTINCT st.tool_name FROM study_tools st "
        "LEFT JOIN tools t ON t.tool_name = st.tool_name WHERE t.tool_name IS NULL"
    ).fetchall()
    if orphan_tools:
        print("     outils absents du master tools :",
              ", ".join(sorted(r[0] for r in orphan_tools)))


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else 'Models_Africa_v2.xlsm'
    dst = sys.argv[2] if len(sys.argv) > 2 else 'enermod.db'
    data = load(src)
    con = build(data, dst)
    print(f"Base créée : {dst}")
    for t in ['countries', 'power_pools', 'tools', 'studies',
              'study_tools', 'study_countries', 'study_pools']:
        n = con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        print(f"  {t:18s}: {n} lignes")
    integrity_report(con)
    con.close()
