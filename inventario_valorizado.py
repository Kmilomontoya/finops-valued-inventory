"""
INVENTARIO MENSUAL VALORIZADO — Azure FOCUS
Azure / FOCUS

═══════════════════════════════════════════════════════════════════════════
  PIPELINE (ejecución única, sin menú)

    L1  LIMPIEZA        : lee el inventario crudo del tenant, descarta filas
                          vacías, limpia valores y deduplica por RESOURCE ID.
    L2  HOMOLOGACIÓN    : unifica columnas-variante (APLICATION/Aplicativo…
                          → APPLICATION ; ENVIROMENT/Ambiente → ENVIRONMENT)
                          y homologa valores a su forma canónica. Las
                          etiquetas canónicas a incluir se FILTRAN en el .env
                          (CANONICAL_TAGS).
    L3  VALORIZACIÓN    : lee el dataset FOCUS del mes (1..N archivos CSV),
                          agrega el costo por ResourceId y hace LEFT JOIN
                          inventario ← costo. El inventario manda: el archivo
                          final tiene exactamente tantas filas como el
                          inventario limpio.
    →   Genera  inventario_valorizado_<MES>.xlsx

  REGLAS
    · Reservation / Snapshot (configurable) se excluyen del costeo y se
      apartan en su propia hoja.
    · Recursos FOCUS sin correspondencia en el inventario → hoja Huérfanos.
    · Prorrateos {APP1:50, APP2:50} se conservan como UNA fila en el detalle;
      en el resumen por APPLICATION el costo se DISTRIBUYE por porcentaje.
═══════════════════════════════════════════════════════════════════════════

Uso:
    pip install pandas openpyxl python-dotenv
    python inventario_valorizado.py
"""

import os
import json
import sys
import traceback
from datetime import datetime
from collections import Counter, defaultdict

import pandas as pd
from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG (.env)
# ─────────────────────────────────────────────────────────────────────────────

load_dotenv()

BASE_DIR = os.getenv(
    "BASE_DIR",
    os.path.join(os.getcwd(), "data"))

# Inventario crudo exportado del tenant
INVENTORY_RAW   = os.getenv("INVENTORY_RAW",
                            os.path.join(BASE_DIR, "Input", "Inventario.xlsx"))
INVENTORY_SHEET = os.getenv("INVENTORY_SHEET", "")     # vacío = primera hoja

# Dataset FOCUS del mes: rutas separadas por coma (archivos o carpetas).
# Soporta el dataset dividido en 2+ archivos por tamaño.
FOCUS_PATHS  = os.getenv("FOCUS_PATHS", os.path.join(BASE_DIR, "Input"))
CSV_ENCODING = os.getenv("CSV_ENCODING", "latin-1")

# Columna de costo del FOCUS a usar para valorizar
COST_COLUMN = os.getenv("COST_COLUMN", "EffectiveCost")

# Etiquetas canónicas a incluir en el archivo final (filtro del .env).
# Separadas por "|" porque algunos nombres contienen espacios.
CANONICAL_TAGS = [
    t.strip() for t in os.getenv(
        "CANONICAL_TAGS",
        "APPLICATION|ENVIRONMENT|TIER|CRITICALITY|SHAREDCOST|DEPARTMENT|"
        "OWNER_TECH|OWNER_TECH MAIL|OWNER_FUNC|OWNER_FUNC MAIL|BACKUPPOLICY"
    ).split("|") if t.strip()
]

# Tipos de recurso a EXCLUIR del costeo (substring, case-insensitive).
EXCLUDE_RESOURCE_TYPES = [
    s.strip().lower()
    for s in os.getenv("EXCLUDE_RESOURCE_TYPES", "reservation,snapshot").split(",")
    if s.strip()
]

OUTPUT_PATH = os.getenv("OUTPUT_PATH", "")   # vacío = autogenerado con el mes

# ─────────────────────────────────────────────────────────────────────────────
# VARIANTES DE COLUMNAS-TAG DEL INVENTARIO CRUDO
# Orden = prioridad (la MAYÚSCULA canónica prima). Extensible vía .env:
#   VARIANTS_APPLICATION=APPLICATION|APLICATION|Aplicativo|Aplicacion
# ─────────────────────────────────────────────────────────────────────────────

DEFAULT_VARIANTS: dict[str, list[str]] = {
    "APPLICATION":  ["APPLICATION", "APLICATION", "Aplicativo",
                     "Aplicacion", "Aplicativo ", "Aplicación"],
    "ENVIRONMENT":  ["ENVIRONMENT", "ENVIROMENT", "Ambiente", "Subambiente"],
    "TIER":         ["TIER", "Capa"],
    "CRITICALITY":  ["CRITICALITY"],
    "SHAREDCOST":   ["SHAREDCOST"],
    "DEPARTMENT":   ["DEPARTMENT", "DEPARTMENT.1", "Department"],
    "OWNER_TECH":   ["OWNER_TECH", "Owner", "Responsible"],
    "OWNER_TECH MAIL": ["OWNER_TECH MAIL", "OWNER_TECH_MAIL"],
    "OWNER_FUNC":   ["OWNER_FUNC"],
    "OWNER_FUNC MAIL": ["OWNER_FUNC MAIL", "OWNER_FUNC_MAIL"],
    "BACKUPPOLICY": ["BACKUPPOLICY"],
}

def _load_variants() -> dict[str, list[str]]:
    """Fusiona variantes del .env (VARIANTS_<CANON>=col1|col2) sobre defaults."""
    variants = {k: list(v) for k, v in DEFAULT_VARIANTS.items()}
    for key, val in os.environ.items():
        if key.startswith("VARIANTS_"):
            canon = key[len("VARIANTS_"):].replace("_", " ").strip()
            # Respetar nombres con guion bajo real si coinciden con canónicas
            if key[len("VARIANTS_"):] in DEFAULT_VARIANTS:
                canon = key[len("VARIANTS_"):]
            cols = [c.strip() for c in val.split("|") if c.strip()]
            if cols:
                variants[canon] = cols
    return variants

VARIANTS = _load_variants()

# ─────────────────────────────────────────────────────────────────────────────
# HOMOLOGACIÓN ENVIRONMENT → canónico (PROD | PREPROD | DEV | NO_DEFINIDO)
# ─────────────────────────────────────────────────────────────────────────────

ENV_HOM = {
    "PROD": "PROD",
    "PRODUCCION": "PROD",
    "PRODUCCIÓN": "PROD",
    "PREPROD": "PREPROD",
    "PREPRODUCCION": "PREPROD",
    "PREPRODUCCIÓN": "PREPROD",
    "PRE PRODUCCION": "PREPROD",
    "STAGE": "PREPROD",
    "TEST": "PREPROD",
    "DEV": "DEV",
    "DESARROLLO": "DEV",
    "NO DEFINIDO": "NO_DEFINIDO",
}

# Extensión opcional desde .env, sin hardcodear valores de un cliente:
# HOM_ENV_MAP=QA>PREPROD;STAGING>PREPROD;DISASTER_RECOVERY>PROD
for _pair in os.getenv("HOM_ENV_MAP", "").split(";"):
    if ">" in _pair:
        _o, _d = _pair.split(">", 1)
        if _o.strip() and _d.strip():
            ENV_HOM[_o.strip().upper()] = _d.strip().upper()

def homologar_env(value) -> str:
    if not value:
        return "NO_DEFINIDO"
    u = str(value).strip().upper()
    if u in ENV_HOM:
        return ENV_HOM[u]
    if "PRE" in u:
        return "PREPROD"
    if "PROD" in u:
        return "PROD"
    if "DESAR" in u or u == "DEV":     # cubre typos: Desarollo, Desarrolo
        return "DEV"
    return "NO_DEFINIDO"

# Homologación APPLICATION vía .env:  HOM_APP_<ORIGEN>=<DESTINO>
#   ej. HOM_APP_APPLEGACY=APP_LEGACY / HOM_APP_PORTAL_INTERNO no es válido como
#   clave de entorno, por eso también se acepta el par en HOM_APP_MAP:
#   HOM_APP_MAP=App Legacy>APP_LEGACY;Portal Interno>PORTAL_INTERNO
APP_HOM: dict[str, str] = {}
for _k, _v in os.environ.items():
    if _k.startswith("HOM_APP_") and _k != "HOM_APP_MAP" and _v.strip():
        APP_HOM[_k[len("HOM_APP_"):].upper()] = _v.strip().upper()
for _pair in os.getenv("HOM_APP_MAP", "").split(";"):
    if ">" in _pair:
        _o, _d = _pair.split(">", 1)
        if _o.strip() and _d.strip():
            APP_HOM[_o.strip().upper()] = _d.strip().upper()

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

BOX = "═" * 70

def clean_val(v):
    try:
        if v is None or pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    # Bytes invisibles ya conocidos de exports Azure
    s = s.replace("\ufeff", "").replace("\u200b", "").replace("\xa0", " ").strip()
    return s if s not in ("", "nan", "None", "NaN", "NULL", "null", "<NA>") else None

def normalize_resource_id(rid) -> str:
    return str(rid or "").strip().lower()

def parse_shared_cost(app_value):
    """'{APP_A:33.3, APP_B:33.3}' → [('APP_A',0.333), …] o None."""
    if not app_value or not str(app_value).strip().startswith("{"):
        return None
    try:
        parts = [p.strip() for p in str(app_value).strip("{}").split(",")]
        out = []
        for p in parts:
            name, pct = p.split(":")
            out.append((name.strip().upper(), float(pct.strip()) / 100.0))
        return out or None
    except Exception:
        return None

def homologar_app(value: str) -> str:
    """Homologa APPLICATION: mapa del .env + MAYÚSCULA. Soporta prorrateos."""
    if not value:
        return None
    shared = parse_shared_cost(value)
    if shared:
        comp = []
        for nm, pct in shared:
            nm_h = APP_HOM.get(nm, nm)
            comp.append(f"{nm_h}:{round(pct * 100, 1)}")
        return "{" + ", ".join(comp) + "}"
    u = str(value).strip().upper()
    return APP_HOM.get(u, u)

def is_excluded_type(resource_type) -> bool:
    t = str(resource_type or "").lower()
    return any(pat in t for pat in EXCLUDE_RESOURCE_TYPES)

def subscription_from_rid(rid: str) -> str:
    """/subscriptions/<guid>/... → <guid>"""
    parts = normalize_resource_id(rid).split("/")
    try:
        i = parts.index("subscriptions")
        return parts[i + 1]
    except (ValueError, IndexError):
        return ""

def resolve_focus_paths(raw: str) -> list[str]:
    paths = []
    for entry in raw.split(","):
        entry = entry.strip()
        if not entry:
            continue
        if os.path.isdir(entry):
            paths.extend(sorted(
                os.path.join(entry, f) for f in os.listdir(entry)
                if f.lower().endswith(".csv")))
        elif os.path.isfile(entry):
            paths.append(entry)
    return paths

# ═════════════════════════════════════════════════════════════════════════════
#  L1 — LIMPIEZA DEL INVENTARIO CRUDO
# ═════════════════════════════════════════════════════════════════════════════

BASE_COLS = ["NAME", "RESOURCE GROUP", "SUBSCRIPTION", "LOCATION",
             "RESOURCE ID", "TYPE"]

# Alias de columnas base: soporta el export crudo del tenant Y el maestro
# de etiquetas v3 (resource_name / resource_group / resource_type / region).
BASE_ALIASES = {
    "NAME":           ["NAME", "resource_name", "RESOURCE NAME", "ResourceName"],
    "RESOURCE GROUP": ["RESOURCE GROUP", "resource_group", "ResourceGroup"],
    "SUBSCRIPTION":   ["SUBSCRIPTION", "subscription", "SubAccountName"],
    "LOCATION":       ["LOCATION", "region", "Region", "location"],
    "RESOURCE ID":    ["RESOURCE ID", "resource_id", "ResourceId", "RESOURCE_ID"],
    "TYPE":           ["TYPE", "resource_type", "ResourceType"],
}

# Modo de cruce contra FOCUS, decidido en L1 según las columnas disponibles:
#   "id"      → por ResourceId normalizado (export del tenant)
#   "name_rg" → por resource_name + resource_group (maestro v3, sin ResourceId)
JOIN_MODE = "id"

def _key_name_rg(name, rg) -> str:
    return f"{str(name or '').strip().lower()}||{str(rg or '').strip().lower()}"

def _detectar_header(sheet) -> int:
    """Detecta la fila de encabezados (los maestros traen fila de título)."""
    preview = pd.read_excel(INVENTORY_RAW, sheet_name=sheet, header=None,
                            nrows=8, dtype=str)
    claves = {"resource id", "resource_id", "resourceid",
              "resource_name", "name"}
    for i in range(len(preview)):
        vals = {str(v).strip().lower() for v in preview.iloc[i].tolist()}
        if vals & claves:
            return i
    return 0

def l1_limpiar() -> pd.DataFrame:
    global JOIN_MODE
    print(f"\n{BOX}\n  L1 — LIMPIEZA DEL INVENTARIO CRUDO\n{BOX}")

    if not os.path.isfile(INVENTORY_RAW):
        raise FileNotFoundError(f"Inventario crudo no encontrado: {INVENTORY_RAW}")

    sheet = INVENTORY_SHEET or 0
    hdr = _detectar_header(sheet)
    if hdr:
        print(f"  ⓘ  Encabezados detectados en la fila {hdr + 1} "
              f"(fila de título ignorada)")
    df = pd.read_excel(INVENTORY_RAW, sheet_name=sheet, dtype=str, header=hdr)
    n0 = len(df)
    print(f"  Archivo : {INVENTORY_RAW}")
    print(f"  Filas leídas          : {n0:,} | Columnas: {len(df.columns)}")

    # Renombrar alias → nombres base canónicos (primer alias presente gana)
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    renames = {}
    for canon, aliases in BASE_ALIASES.items():
        if canon in df.columns:
            continue
        for a in aliases:
            col = lower_map.get(a.lower())
            if col is not None:
                renames[col] = canon
                break
    if renames:
        df = df.rename(columns=renames)
        print(f"  ⓘ  Columnas homologadas al esquema base: "
              f"{ {v: k for k, v in renames.items()} }")

    # Limpieza celda a celda de columnas base y variantes de tags
    cols_interes = set(BASE_COLS)
    for canon in CANONICAL_TAGS:
        cols_interes.update(VARIANTS.get(canon, [canon]))
    cols_interes &= set(df.columns)
    for c in cols_interes:
        df[c] = df[c].map(clean_val)

    # Modo de cruce: ResourceId si existe y está poblado; si no, nombre+RG
    tiene_id = ("RESOURCE ID" in df.columns
                and df["RESOURCE ID"].notna().mean() > 0.5)
    if tiene_id:
        JOIN_MODE = "id"
        df = df[df["RESOURCE ID"].notna()].copy()
        if n0 - len(df):
            print(f"  Filas sin RESOURCE ID descartadas: {n0 - len(df):,}")
        df["join_key"] = df["RESOURCE ID"].map(normalize_resource_id)
        df["resource_id_norm"] = df["join_key"]
        print(f"  Modo de cruce          : ResourceId")
    else:
        JOIN_MODE = "name_rg"
        if "NAME" not in df.columns or "RESOURCE GROUP" not in df.columns:
            raise KeyError(
                "El archivo no tiene RESOURCE ID ni el par "
                "NAME/resource_name + RESOURCE GROUP/resource_group "
                "necesarios para cruzar contra FOCUS.")
        df = df[df["NAME"].notna()].copy()
        if n0 - len(df):
            print(f"  Filas sin nombre de recurso descartadas: {n0 - len(df):,}")
        df["join_key"] = [
            _key_name_rg(n, g) for n, g in zip(df["NAME"], df["RESOURCE GROUP"])]
        df["resource_id_norm"] = None
        print(f"  Modo de cruce          : nombre + resource group "
              f"(el archivo no trae ResourceId)")

    dup = df["join_key"].duplicated().sum()
    if dup:
        df = df.drop_duplicates("join_key", keep="first")
        print(f"  Duplicados de llave eliminados: {dup:,}")

    # SUBSCRIPTION: derivar del RESOURCE ID si viene vacía (solo modo id)
    if "SUBSCRIPTION" not in df.columns:
        df["SUBSCRIPTION"] = None
    if JOIN_MODE == "id":
        vacias = df["SUBSCRIPTION"].isna().sum()
        mask = df["SUBSCRIPTION"].isna()
        if mask.any():
            df.loc[mask, "SUBSCRIPTION"] = df.loc[mask, "join_key"].map(
                subscription_from_rid)
            print(f"  SUBSCRIPTION derivada del RESOURCE ID: {vacias:,} filas")
    for c in BASE_COLS:
        if c not in df.columns:
            df[c] = None

    print(f"  ✔  Inventario limpio: {len(df):,} recursos")
    return df.reset_index(drop=True)

# ═════════════════════════════════════════════════════════════════════════════
#  L2 — HOMOLOGACIÓN A ETIQUETAS CANÓNICAS (filtradas en .env)
# ═════════════════════════════════════════════════════════════════════════════

def l2_homologar(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    print(f"\n{BOX}\n  L2 — HOMOLOGACIÓN DE ETIQUETAS CANÓNICAS\n{BOX}")
    print(f"  Etiquetas canónicas (.env): {CANONICAL_TAGS}")
    if APP_HOM:
        print(f"  Homologaciones APPLICATION (.env): {len(APP_HOM)} mapeos")

    audit = {
        "conflictos": Counter(),
        "origen":     defaultdict(Counter),
        "env_sin_homologar": Counter(),
    }

    for canon in CANONICAL_TAGS:
        present = [c for c in VARIANTS.get(canon, [canon]) if c in df.columns]
        if not present:
            df[f"_{canon}"] = None
            print(f"  ⚠  {canon}: sin columnas fuente en el inventario")
            continue

        def resolver(row, cols=present, tag=canon):
            vals = [(c, clean_val(row[c])) for c in cols]
            vals = [(c, v) for c, v in vals if v is not None]
            if not vals:
                return None
            if len({v.upper() for _, v in vals}) > 1:
                audit["conflictos"][tag] += 1
            col_src, val = vals[0]              # prioridad = orden de variantes
            audit["origen"][tag][col_src] += 1
            return val

        # apply puede castear a dtype str y volver los None → NaN float
        # (truthy); forzar dtype object garantiza None reales entre etapas.
        serie = df[present].apply(resolver, axis=1)
        serie = pd.Series([clean_val(x) for x in serie],
                          index=serie.index, dtype=object)
        con_dato = serie.notna().sum()

        # Homologación por tipo de etiqueta (listas → dtype object estable)
        if canon == "ENVIRONMENT":
            def _env(v):
                if v is None:
                    return "NO_DEFINIDO"
                h = homologar_env(v)
                if h == "NO_DEFINIDO" and v.upper() not in ENV_HOM:
                    audit["env_sin_homologar"][v] += 1
                return h
            serie = pd.Series([_env(v) for v in serie],
                              index=serie.index, dtype=object)
        elif canon == "APPLICATION":
            serie = pd.Series([homologar_app(v) if v is not None else None
                               for v in serie], index=serie.index, dtype=object)
        elif canon in ("TIER", "CRITICALITY", "SHAREDCOST"):
            serie = pd.Series([v.upper() if v is not None else None
                               for v in serie], index=serie.index, dtype=object)

        df[f"_{canon}"] = serie
        print(f"  {canon:<18} con dato: {con_dato:>6,} "
              f"| fuente: {dict(audit['origen'][canon])}")

    if audit["conflictos"]:
        print(f"  Conflictos entre variantes (ganó prioridad): "
              f"{dict(audit['conflictos'])}")
    if audit["env_sin_homologar"]:
        print(f"  ⚠  ENVIRONMENT sin homologar "
              f"({len(audit['env_sin_homologar'])} valores):")
        for v, n in audit["env_sin_homologar"].most_common(10):
            print(f"       {v!r} ({n})")

    return df, audit

# ═════════════════════════════════════════════════════════════════════════════
#  L3 — VALORIZACIÓN CONTRA FOCUS
# ═════════════════════════════════════════════════════════════════════════════

def _leer_focus(path: str) -> pd.DataFrame:
    """Lee un CSV FOCUS tolerando BOM, encoding y headers repetidos."""
    cols = ["ResourceId", "ResourceName", "ResourceType", "ServiceName",
            "x_ResourceGroupName", "SubAccountName", "ChargePeriodStart",
            COST_COLUMN]
    try:
        d = pd.read_csv(path, encoding=CSV_ENCODING, low_memory=False,
                        usecols=lambda c: c in cols)
    except UnicodeDecodeError:
        d = pd.read_csv(path, encoding="utf-8-sig", low_memory=False,
                        usecols=lambda c: c in cols)
    # Headers repetidos incrustados como datos (concatenaciones de exports)
    if "ResourceId" in d.columns:
        d = d[d["ResourceId"].astype(str) != "ResourceId"]
    # Costo a numérico (bytes invisibles → coerce)
    d[COST_COLUMN] = pd.to_numeric(
        d[COST_COLUMN].astype(str)
         .str.replace("\ufeff", "").str.replace("\u200b", "").str.strip(),
        errors="coerce").fillna(0.0)
    return d

def l3_valorizar(inv: pd.DataFrame):
    print(f"\n{BOX}\n  L3 — VALORIZACIÓN CONTRA FOCUS\n{BOX}")

    paths = resolve_focus_paths(FOCUS_PATHS)
    if not paths:
        raise FileNotFoundError(f"No se encontraron CSV FOCUS en: {FOCUS_PATHS}")
    print(f"  Archivos FOCUS ({len(paths)}):")
    for p in paths:
        print(f"    · {p}")

    frames = [_leer_focus(p) for p in paths]
    focus = pd.concat(frames, ignore_index=True)
    print(f"  Filas FOCUS totales    : {len(focus):,}")
    print(f"  Columna de costo       : {COST_COLUMN}")

    # Mes del dataset (para nombrar el archivo y la hoja)
    mes = ""
    if "ChargePeriodStart" in focus.columns:
        fechas = pd.to_datetime(
            focus["ChargePeriodStart"].astype(str).str.replace("\ufeff", ""),
            errors="coerce", utc=True)
        if fechas.notna().any():
            mes = fechas.dropna().dt.strftime("%Y-%m").mode().iloc[0]
    print(f"  Mes detectado          : {mes or '(no detectado)'}")

    if JOIN_MODE == "id":
        focus["join_key"] = focus["ResourceId"].map(normalize_resource_id)
    else:
        focus["join_key"] = [
            _key_name_rg(n, g) for n, g in
            zip(focus.get("ResourceName"), focus.get("x_ResourceGroupName"))]
    focus["resource_id_norm"] = focus["ResourceId"].map(normalize_resource_id)

    # Exclusiones (reservation / snapshot) — se apartan, no cruzan
    excl_mask = focus["ResourceType"].map(is_excluded_type)
    if "ServiceName" in focus.columns:
        excl_mask |= focus["ServiceName"].map(is_excluded_type)
    focus_excl = focus[excl_mask].copy()
    focus_cost = focus[~excl_mask].copy()
    print(f"  Excluidos ({', '.join(EXCLUDE_RESOURCE_TYPES)}): "
          f"{len(focus_excl):,} filas | "
          f"${focus_excl[COST_COLUMN].sum():,.2f}")

    # Costo agregado por la llave de cruce (ResourceId o nombre+RG)
    cost_by_key = (focus_cost[focus_cost["join_key"].astype(str).str.strip("|") != ""]
                   .groupby("join_key")[COST_COLUMN].sum())

    # LEFT JOIN: el inventario manda (mismas filas que el inventario limpio)
    inv = inv.copy()
    inv["COSTO_MES"] = inv["join_key"].map(cost_by_key).fillna(0.0)
    inv["ESTADO_COSTO"] = inv["COSTO_MES"].apply(
        lambda c: "CON_COSTO" if c > 0 else "SIN_COSTO_FOCUS")

    # Huérfanos: costo FOCUS que no está en el inventario
    inv_keys = set(inv["join_key"])
    orphan = focus_cost[(~focus_cost["join_key"].isin(inv_keys)) |
                        (focus_cost["join_key"].astype(str).str.strip("|") == "")].copy()
    orphan_agg = (orphan.groupby(
                     ["resource_id_norm", "ResourceName", "ResourceType",
                      "x_ResourceGroupName", "SubAccountName"], dropna=False)
                  [COST_COLUMN].sum().reset_index()
                  .sort_values(COST_COLUMN, ascending=False))

    excl_agg = (focus_excl.groupby(
                   ["resource_id_norm", "ResourceName", "ResourceType",
                    "x_ResourceGroupName"], dropna=False)
                [COST_COLUMN].sum().reset_index()
                .sort_values(COST_COLUMN, ascending=False))

    n_con = (inv["ESTADO_COSTO"] == "CON_COSTO").sum()
    print(f"  Recursos con costo     : {n_con:,}")
    print(f"  Sin costo en FOCUS     : {len(inv) - n_con:,}")
    print(f"  Costo total inventario : ${inv['COSTO_MES'].sum():,.2f}")
    print(f"  Huérfanos FOCUS        : {len(orphan_agg):,} recursos | "
          f"${orphan_agg[COST_COLUMN].sum():,.2f}")

    return inv, orphan_agg, excl_agg, mes

# ═════════════════════════════════════════════════════════════════════════════
#  RESÚMENES
# ═════════════════════════════════════════════════════════════════════════════

def resumen_application(inv: pd.DataFrame) -> pd.DataFrame:
    """Costo por APPLICATION distribuyendo prorrateos por porcentaje."""
    filas = []
    for app, costo in zip(inv.get("_APPLICATION"), inv["COSTO_MES"]):
        shared = parse_shared_cost(app)
        if shared:
            # Normalizar pesos: tags como {A:33.3, B:33.3, C:33.3} suman 99.9%;
            # se reescala para conservar el costo total del inventario.
            tot_pct = sum(p for _, p in shared) or 1.0
            for nm, pct in shared:
                filas.append((nm, costo * pct / tot_pct, 1 / len(shared)))
        else:
            filas.append((app or "NO_ETIQUETADO", costo, 1))
    r = pd.DataFrame(filas, columns=["APPLICATION", "COSTO_MES", "RECURSOS"])
    out = (r.groupby("APPLICATION", dropna=False)
             .agg(COSTO_MES=("COSTO_MES", "sum"), RECURSOS=("RECURSOS", "sum"))
             .reset_index().sort_values("COSTO_MES", ascending=False))
    total = out["COSTO_MES"].sum()
    out["% COSTO"] = out["COSTO_MES"] / total if total else 0.0
    out["RECURSOS"] = out["RECURSOS"].round(1)
    return out

def resumen_simple(inv: pd.DataFrame, col: str, nombre: str) -> pd.DataFrame:
    out = (inv.assign(**{nombre: inv[col].fillna("NO_ETIQUETADO")})
              .groupby(nombre, dropna=False)
              .agg(COSTO_MES=("COSTO_MES", "sum"),
                   RECURSOS=("COSTO_MES", "size"))
              .reset_index().sort_values("COSTO_MES", ascending=False))
    total = out["COSTO_MES"].sum()
    out["% COSTO"] = out["COSTO_MES"] / total if total else 0.0
    return out

# ═════════════════════════════════════════════════════════════════════════════
#  EXCEL
# ═════════════════════════════════════════════════════════════════════════════

H_FILL = PatternFill("solid", fgColor="1F3864")
H_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
ALT    = PatternFill("solid", fgColor="F2F6FC")
WHT    = PatternFill("solid", fgColor="FFFFFF")
BRD    = Border(*[Side(style="thin", color="D0D7E5")] * 4)
C_ALN  = Alignment(horizontal="center", vertical="center")
L_ALN  = Alignment(horizontal="left",   vertical="center")

def _title(ws, text, nc, row=1):
    ws.merge_cells(f"A{row}:{get_column_letter(nc)}{row}")
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(bold=True, size=12, color="1F3864", name="Arial")
    c.alignment = C_ALN
    c.fill = PatternFill("solid", fgColor="D9E1F2")

def _header(ws, row, nc):
    for c in range(1, nc + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = H_FILL, H_FONT
        cell.alignment, cell.border = C_ALN, BRD

def _body(ws, r0, r1, nc):
    for r in range(r0, r1 + 1):
        fill = ALT if r % 2 == 0 else WHT
        for c in range(1, nc + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill, cell.border, cell.alignment = fill, BRD, L_ALN

def _autowidth(ws, mn=10, mx=55):
    for col in ws.columns:
        w = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            min(max(w + 3, mn), mx)

def _formato_moneda(ws, df, startrow):
    for j, col in enumerate(df.columns, start=1):
        if col in ("COSTO_MES", COST_COLUMN):
            for r in range(startrow + 1, startrow + 1 + len(df)):
                ws.cell(row=r, column=j).number_format = "#,##0.00"
        elif col == "% COSTO":
            for r in range(startrow + 1, startrow + 1 + len(df)):
                ws.cell(row=r, column=j).number_format = "0.0%"

def _escribir_hoja(writer, df, sheet, titulo):
    df.to_excel(writer, sheet_name=sheet, index=False, startrow=1)
    ws = writer.sheets[sheet]
    nc = max(len(df.columns), 1)
    _title(ws, titulo, nc)
    _header(ws, 2, nc)
    if len(df):
        _body(ws, 3, 2 + len(df), nc)
    _formato_moneda(ws, df, startrow=2)
    ws.auto_filter.ref = f"A2:{get_column_letter(nc)}2"
    ws.freeze_panes = "A3"
    _autowidth(ws)

def generar_excel(inv, orphan, excl, mes, audit):
    # OUTPUT_PATH puede ser: vacío (autogenerar en BASE_DIR), una CARPETA
    # (autogenerar el nombre dentro de ella) o la ruta completa del .xlsx.
    nombre = f"inventario_valorizado_{mes or 'mes'}.xlsx"
    salida = OUTPUT_PATH.strip()
    if not salida:
        salida = os.path.join(BASE_DIR, nombre)
    elif os.path.isdir(salida) or not salida.lower().endswith((".xlsx", ".xlsm")):
        os.makedirs(salida, exist_ok=True)
        salida = os.path.join(salida, nombre)
    os.makedirs(os.path.dirname(salida) or ".", exist_ok=True)
    print(f"\n{BOX}\n  GENERANDO EXCEL\n{BOX}")

    # Hoja 1: detalle — columnas base + etiquetas canónicas filtradas + costo
    detalle = pd.DataFrame({
        "NAME":           inv["NAME"],
        "RESOURCE GROUP": inv["RESOURCE GROUP"],
        "SUBSCRIPTION":   inv["SUBSCRIPTION"],
        "LOCATION":       inv.get("LOCATION"),
        "RESOURCE ID":    inv["RESOURCE ID"],
        "TYPE":           inv["TYPE"],
    })
    for canon in CANONICAL_TAGS:
        detalle[canon] = inv.get(f"_{canon}")
    detalle["COSTO_MES"]    = inv["COSTO_MES"]
    detalle["ESTADO_COSTO"] = inv["ESTADO_COSTO"]
    detalle = detalle.sort_values("COSTO_MES", ascending=False)

    r_app = resumen_application(inv)
    r_env = (resumen_simple(inv, "_ENVIRONMENT", "ENVIRONMENT")
             if "_ENVIRONMENT" in inv.columns else pd.DataFrame())
    r_typ = resumen_simple(inv, "TYPE", "TYPE")

    orphan = orphan.rename(columns={COST_COLUMN: "COSTO_MES"})
    excl   = excl.rename(columns={COST_COLUMN: "COSTO_MES"})

    try:
        writer_cm = pd.ExcelWriter(salida, engine="openpyxl")
    except PermissionError:
        print(f"\n  ✖  Sin permiso para escribir: {salida}")
        print("     Si el archivo está abierto en Excel, ciérralo y")
        print("     vuelve a ejecutar la opción 1.")
        raise
    with writer_cm as writer:
        _escribir_hoja(writer, detalle, "Inventario_Valorizado",
                       f"INVENTARIO VALORIZADO — {mes or ''} — "
                       f"Generado {datetime.now():%Y-%m-%d %H:%M}")
        _escribir_hoja(writer, r_app, "Resumen_APPLICATION",
                       "COSTO POR APPLICATION (prorrateos distribuidos)")
        if len(r_env):
            _escribir_hoja(writer, r_env, "Resumen_ENVIRONMENT",
                           "COSTO POR ENVIRONMENT")
        _escribir_hoja(writer, r_typ, "Resumen_TYPE", "COSTO POR TIPO DE RECURSO")
        _escribir_hoja(writer, orphan, "Huerfanos_FOCUS",
                       "COSTO FOCUS SIN RECURSO EN EL INVENTARIO")
        _escribir_hoja(writer, excl, "Excluidos",
                       f"EXCLUIDOS DEL COSTEO ({', '.join(EXCLUDE_RESOURCE_TYPES)})")

        # Hoja Auditoría
        tot = inv["COSTO_MES"].sum()
        n_con = (inv["ESTADO_COSTO"] == "CON_COSTO").sum()
        rep = [("Métrica", "Valor"),
               ("Mes valorizado", mes or "(no detectado)"),
               ("Recursos inventario limpio", f"{len(inv):,}"),
               ("Recursos con costo FOCUS", f"{n_con:,}"),
               ("Recursos sin costo FOCUS", f"{len(inv) - n_con:,}"),
               ("Costo total inventario", f"${tot:,.2f}"),
               ("Huérfanos FOCUS (recursos)", f"{len(orphan):,}"),
               ("Huérfanos FOCUS (costo)", f"${orphan['COSTO_MES'].sum():,.2f}"),
               ("Excluidos (costo)", f"${excl['COSTO_MES'].sum():,.2f}"),
               ("", ""),
               ("Etiquetas canónicas incluidas", " | ".join(CANONICAL_TAGS)),
               ("", ""), ("Cobertura de etiquetado (sobre costo)", "")]
        for canon in CANONICAL_TAGS:
            col = f"_{canon}"
            if col in inv.columns:
                con = inv[inv[col].notna() &
                          (~inv[col].isin(["NO_DEFINIDO", "NO DEFINIDO"]))]
                pct = con["COSTO_MES"].sum() / tot if tot else 0
                rep.append((f"  {canon}", f"{pct * 100:.1f}%"))
        rep += [("", ""), ("Conflictos entre variantes", "")]
        for tag, n in audit["conflictos"].items():
            rep.append((f"  {tag}", n))
        if audit["env_sin_homologar"]:
            rep += [("", ""), ("ENVIRONMENT sin homologar", "")]
            for v, n in audit["env_sin_homologar"].most_common(20):
                rep.append((f"  {v}", n))

        ws = writer.book.create_sheet("Auditoria")
        ws.sheet_view.showGridLines = False
        SKEYS = {"Métrica", "Cobertura de etiquetado (sobre costo)",
                 "Conflictos entre variantes", "ENVIRONMENT sin homologar"}
        for i, (k, v) in enumerate(rep, start=1):
            ck = ws.cell(row=i, column=1, value=k)
            cv = ws.cell(row=i, column=2, value=v)
            sec = k in SKEYS
            for cell in (ck, cv):
                cell.font = Font(bold=sec, name="Arial", size=10,
                                 color="1F3864" if sec else "000000")
                if sec:
                    cell.fill = PatternFill("solid", fgColor="D9E1F2")
                cell.alignment = L_ALN
        ws.column_dimensions["A"].width = 45
        ws.column_dimensions["B"].width = 30

    print(f"  ✔  Archivo generado: {salida}")
    return salida


# ═════════════════════════════════════════════════════════════════════════════
#  REPORTE KPI — COBERTURA DE COSTO ETIQUETADO
# ═════════════════════════════════════════════════════════════════════════════

META_COBERTURA = float(os.getenv("META_COBERTURA", "80"))   # % objetivo

SIN_ETIQUETA = {"NO_DEFINIDO", "NO DEFINIDO", "NO_ETIQUETADO"}

def _mask_etiquetado(serie: pd.Series) -> pd.Series:
    return serie.notna() & (~serie.astype(str).str.upper().isin(SIN_ETIQUETA))

def reporte_kpi(inv: pd.DataFrame, mes: str) -> str:
    """
    KPI de cobertura de costo etiquetado sobre un maestro valorizado:
    cuánto cuesta el inventario y qué % del costo tiene cada etiqueta.
    """
    print(f"\n{BOX}\n  REPORTE KPI — COBERTURA DE COSTO ETIQUETADO\n{BOX}")
    tot   = inv["COSTO_MES"].sum()
    n_tot = len(inv)
    n_con = (inv["ESTADO_COSTO"] == "CON_COSTO").sum()

    print(f"  Mes                    : {mes or '(no detectado)'}")
    print(f"  Recursos del maestro   : {n_tot:,}  (con costo FOCUS: {n_con:,})")
    print(f"  COSTO TOTAL INVENTARIO : ${tot:,.2f}")
    print(f"  Meta de cobertura      : {META_COBERTURA:.0f}%\n")

    filas = []
    for canon in CANONICAL_TAGS:
        col = f"_{canon}"
        if col not in inv.columns:
            continue
        m = _mask_etiquetado(inv[col])
        c_tag = inv.loc[m, "COSTO_MES"].sum()
        pct_c = (c_tag / tot * 100) if tot else 0.0
        pct_r = m.sum() / n_tot * 100 if n_tot else 0.0
        estado = "✔ CUMPLE" if pct_c >= META_COBERTURA else "✖ BAJO META"
        filas.append({
            "ETIQUETA": canon,
            "RECURSOS ETIQUETADOS": int(m.sum()),
            "% RECURSOS": round(pct_r / 100, 4),
            "COSTO ETIQUETADO": round(c_tag, 2),
            "% COSTO ETIQUETADO": round(pct_c / 100, 4),
            "BRECHA A META (USD)": round(max(0.0, tot * META_COBERTURA / 100 - c_tag), 2),
            "ESTADO": estado,
        })
        print(f"  {canon:<18} costo etiquetado: {pct_c:6.1f}%  "
              f"(recursos: {pct_r:5.1f}%)  {estado}")

    kpi = pd.DataFrame(filas)

    # Gap accionable: recursos con mayor costo SIN etiqueta APPLICATION
    gap_cols = ["NAME", "RESOURCE GROUP", "SUBSCRIPTION", "TYPE",
                "RESOURCE ID", "COSTO_MES"]
    if "_APPLICATION" in inv.columns:
        sin_app = inv[~_mask_etiquetado(inv["_APPLICATION"])]
        gap = (sin_app[gap_cols].sort_values("COSTO_MES", ascending=False)
               .head(200).copy())
        costo_gap = sin_app["COSTO_MES"].sum()
        print(f"\n  Costo SIN etiqueta APPLICATION: ${costo_gap:,.2f} "
              f"({(costo_gap / tot * 100) if tot else 0:.1f}% del total) "
              f"— top 200 recursos en la hoja Gap_APPLICATION")
    else:
        gap = pd.DataFrame(columns=gap_cols)

    # Cobertura por RESOURCE GROUP (dónde atacar primero)
    if "_APPLICATION" in inv.columns and tot:
        rg = inv.assign(_etq=_mask_etiquetado(inv["_APPLICATION"]))
        rg = (rg.groupby(rg["RESOURCE GROUP"].fillna("(sin RG)"))
                .apply(lambda g: pd.Series({
                    "COSTO_MES": g["COSTO_MES"].sum(),
                    "COSTO ETIQUETADO": g.loc[g["_etq"], "COSTO_MES"].sum(),
                    "RECURSOS": len(g)}), include_groups=False)
                .reset_index()
                .rename(columns={"RESOURCE GROUP": "RESOURCE GROUP"}))
        rg["% COSTO ETIQUETADO"] = (
            rg["COSTO ETIQUETADO"] / rg["COSTO_MES"]).where(rg["COSTO_MES"] > 0, 0).round(4)
        rg = rg.sort_values("COSTO_MES", ascending=False).head(100)
    else:
        rg = pd.DataFrame()

    nombre = f"kpi_cobertura_{mes or 'mes'}.xlsx"
    salida = OUTPUT_PATH.strip()
    if not salida:
        salida = os.path.join(BASE_DIR, nombre)
    elif os.path.isdir(salida) or not salida.lower().endswith((".xlsx", ".xlsm")):
        os.makedirs(salida, exist_ok=True)
        salida = os.path.join(salida, nombre)
    else:
        salida = os.path.join(os.path.dirname(salida) or ".", nombre)
    os.makedirs(os.path.dirname(salida) or ".", exist_ok=True)

    try:
        writer_cm = pd.ExcelWriter(salida, engine="openpyxl")
    except PermissionError:
        print(f"\n  ✖  Sin permiso para escribir: {salida}")
        print("     Si el archivo está abierto en Excel, ciérralo y reintenta.")
        raise
    with writer_cm as writer:
        resumen = pd.DataFrame([
            ("Mes", mes or "(no detectado)"),
            ("Recursos del maestro", n_tot),
            ("Recursos con costo FOCUS", n_con),
            ("COSTO TOTAL INVENTARIO (USD)", round(tot, 2)),
            ("Meta de cobertura (%)", META_COBERTURA),
        ], columns=["Métrica", "Valor"])
        _escribir_hoja(writer, resumen, "Resumen",
                       f"KPI COBERTURA DE COSTO ETIQUETADO — {mes or ''} — "
                       f"Generado {datetime.now():%Y-%m-%d %H:%M}")
        _escribir_hoja(writer, kpi, "KPI_Cobertura",
                       f"COBERTURA POR ETIQUETA CANÓNICA (meta {META_COBERTURA:.0f}%)")
        _escribir_hoja(writer, gap, "Gap_APPLICATION",
                       "TOP RECURSOS POR COSTO SIN ETIQUETA APPLICATION")
        if len(rg):
            _escribir_hoja(writer, rg, "Cobertura_por_RG",
                           "COBERTURA APPLICATION POR RESOURCE GROUP (top 100 por costo)")

    print(f"\n  ✔  Reporte KPI generado: {salida}")
    return salida

def ejecutar_kpi(inv_path: str):
    global INVENTORY_RAW
    INVENTORY_RAW = inv_path
    inv = l1_limpiar()
    inv, _audit = l2_homologar(inv)
    inv, _orphan, _excl, mes = l3_valorizar(inv)
    salida = reporte_kpi(inv, mes)
    print(f"\n{BOX}\n  ✔  PROCESO COMPLETADO\n  →  {salida}\n{BOX}")

# ═════════════════════════════════════════════════════════════════════════════
#  MENÚ INTERACTIVO
# ═════════════════════════════════════════════════════════════════════════════

def _fmt_size(path: str) -> str:
    try:
        kb = os.path.getsize(path) / 1024
        return f"{kb / 1024:,.1f} MB" if kb >= 1024 else f"{kb:,.0f} KB"
    except OSError:
        return "?"

def _autodetectar_inventario() -> str:
    """Si INVENTORY_RAW no existe, busca un .xlsx en las carpetas de insumos."""
    if os.path.isfile(INVENTORY_RAW):
        return INVENTORY_RAW
    candidatos = []
    for carpeta in {os.path.dirname(INVENTORY_RAW),
                    os.path.join(BASE_DIR, "Input"), FOCUS_PATHS}:
        if carpeta and os.path.isdir(carpeta):
            candidatos.extend(
                os.path.join(carpeta, f) for f in sorted(os.listdir(carpeta))
                if f.lower().endswith((".xlsx", ".xlsm"))
                and not f.startswith("~$"))
    return candidatos[0] if candidatos else INVENTORY_RAW

def verificar_insumos(silencioso=False):
    """Muestra qué encuentra el script. Devuelve (inv_path, csv_paths, ok)."""
    inv_path  = _autodetectar_inventario()
    csv_paths = resolve_focus_paths(FOCUS_PATHS)
    inv_ok    = os.path.isfile(inv_path)
    csv_ok    = len(csv_paths) > 0

    if not silencioso:
        print(f"\n  [Verificación de insumos]")
        print(f"  {'✔' if inv_ok else '✖'}  Inventario crudo (xlsx)")
        if inv_ok:
            print(f"       · {inv_path}  ({_fmt_size(inv_path)})")
            if inv_path != INVENTORY_RAW:
                print(f"       ⓘ  autodetectado (INVENTORY_RAW del .env no existe)")
        else:
            print(f"       ✖  no encontrado: {INVENTORY_RAW}")
        print(f"  {'✔' if csv_ok else '✖'}  Dataset FOCUS (csv): "
              f"{len(csv_paths)} archivo(s)")
        for p in csv_paths:
            print(f"       · {p}  ({_fmt_size(p)})")
        if not csv_ok:
            print(f"       ✖  sin CSV en: {FOCUS_PATHS}")
        print(f"  ⚙  Columna de costo : {COST_COLUMN}")
        print(f"  ⚙  Exclusiones      : {', '.join(EXCLUDE_RESOURCE_TYPES)}")
        print(f"  ⚙  Etiquetas (.env) : {' | '.join(CANONICAL_TAGS)}")

    return inv_path, csv_paths, (inv_ok and csv_ok)

def ejecutar_pipeline(inv_path: str):
    global INVENTORY_RAW
    INVENTORY_RAW = inv_path
    inv = l1_limpiar()
    inv, audit = l2_homologar(inv)
    inv, orphan, excl, mes = l3_valorizar(inv)
    salida = generar_excel(inv, orphan, excl, mes, audit)
    print(f"\n{BOX}\n  ✔  PROCESO COMPLETADO\n  →  {salida}\n{BOX}")

def menu():
    while True:
        print(f"\n{BOX}")
        print("  INVENTARIO MENSUAL VALORIZADO — Azure FOCUS")
        print("  Azure / FOCUS")
        print(BOX)
        inv_path, csv_paths, ok = verificar_insumos()
        print(f"\n  ┌{'─' * 58}┐")
        print(f"  │  1.  Ejecutar pipeline completo (L1→L2→L3 + Excel)       │")
        print(f"  │  2.  KPI de cobertura de costo etiquetado (maestro)      │")
        print(f"  │  3.  Re-verificar insumos                                │")
        print(f"  │  0.  Salir                                               │")
        print(f"  └{'─' * 58}┘")
        op = input("  Opción: ").strip()
        if op in ("1", "2"):
            if not ok:
                print("\n  ✖  Faltan insumos. Revisa las rutas del .env "
                      "antes de ejecutar.")
                continue
            try:
                if op == "1":
                    ejecutar_pipeline(inv_path)
                else:
                    ejecutar_kpi(inv_path)
            except Exception:
                print("\n  ✖  ERROR:")
                traceback.print_exc()
            input("\n  Presiona ENTER para volver al menú...")
        elif op == "3":
            continue
        elif op == "0":
            print("  Hasta luego.\n")
            break
        else:
            print("  Opción no válida.")

def main():
    try:
        menu()
    except KeyboardInterrupt:
        print("\n  Interrumpido por el usuario.\n")

if __name__ == "__main__":
    main()
