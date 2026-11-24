import base64
import hashlib
import io
import json
import math
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

import numpy as np
import openpyxl
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import streamlit.components.v1 as components
from fpdf import FPDF

# -----------------------------------------------------------------------------
# CONFIGURATION DE LA PAGE & STYLE UI
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="GRAFF — Race Engineering & Driver Debrief",
    page_icon="🏁",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp {
        background-color: #0e1117;
        color: #f0f2f6;
    }
    .graff-banner {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background-color: #0f172a;
        border: 1px solid #1e293b;
        border-radius: 12px;
        padding: 14px 22px;
        margin-bottom: 18px;
        gap: 16px;
    }
    .graff-banner-title {
        font-size: 1.35rem;
        font-weight: 800;
        color: #f8fafc;
        letter-spacing: 0.04em;
        margin: 0;
    }
    .graff-banner-sub {
        font-size: 0.85rem;
        color: #94a3b8;
        margin-top: 3px;
    }
    .graff-logo-fallback {
        font-size: 1.25rem;
        font-weight: 900;
        letter-spacing: 0.18em;
        color: #f8fafc;
        border: 1.5px solid #475569;
        padding: 6px 14px;
        border-radius: 6px;
    }
    .driver-card {
        background: linear-gradient(135deg, #161b26 0%, #1d2434 100%);
        border: 1px solid #2d3748;
        border-radius: 12px;
        padding: 16px 20px;
        margin-top: 8px;
        margin-bottom: 16px;
    }
    .driver-card-ref {
        background: linear-gradient(135deg, #16222a 0%, #1b2d38 100%);
        border: 1px solid #2b6cb0;
        border-radius: 12px;
        padding: 16px 20px;
        margin-top: 8px;
        margin-bottom: 16px;
    }
    .metric-label {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #94a3b8;
        margin-bottom: 2px;
    }
    .metric-value {
        font-size: 1.35rem;
        font-weight: 700;
        color: #f8fafc;
        font-family: 'Consolas', monospace;
    }
    .total-gap-row {
        display: flex;
        align-items: center;
        gap: 20px;
        margin: 14px 0 22px 0;
    }
    .total-gap-label {
        font-size: 1.25rem;
        font-weight: 700;
        color: #e2e8f0;
        min-width: 140px;
    }
    .gap-box-pos {
        background-color: rgba(239, 68, 68, 0.20);
        border: 2px solid #ef4444;
        color: #fca5a5;
        border-radius: 8px;
        padding: 10px 36px;
        font-size: 1.55rem;
        font-weight: 800;
        font-family: 'Consolas', monospace;
        text-align: center;
    }
    .gap-box-neg {
        background-color: rgba(34, 197, 94, 0.20);
        border: 2px solid #22c55e;
        color: #86efac;
        border-radius: 8px;
        padding: 10px 36px;
        font-size: 1.55rem;
        font-weight: 800;
        font-family: 'Consolas', monospace;
        text-align: center;
    }
    .theo-hero-card {
        background: linear-gradient(145deg, #111827 0%, #172033 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 16px 18px;
        height: 410px;
        box-sizing: border-box;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        margin-bottom: 14px;
    }
    .theo-main-box {
        background: rgba(56, 189, 248, 0.10);
        border: 1.5px solid #0284c7;
        border-radius: 10px;
        padding: 12px 14px;
        text-align: center;
        margin: 8px 0;
    }
    .theo-time-big {
        font-size: 2.0rem;
        font-weight: 800;
        color: #38bdf8;
        font-family: 'Consolas', monospace;
        letter-spacing: 0.03em;
    }
    .netlify-link-box {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.14) 0%, rgba(14, 165, 233, 0.14) 100%);
        border: 1.5px solid #10b981;
        border-radius: 10px;
        padding: 12px 16px;
        margin: 10px 0 14px 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

PRESETS_FILE = "circuits_presets.json"
NETLIFY_CONFIG_FILE = "netlify_config.json"
PUBLISHED_DIR = "published_pages"

DEFAULT_NETLIFY_SITE_ID = "0b9fc4d5-5e60-426a-ba67-d5a11101ab6e"
DEFAULT_NETLIFY_TOKEN = "nfp_y86nSY1FzGoCNUoQjp8tZ2HWy6sZoh4hc915"


def load_netlify_config():
    cfg = {"site_id": DEFAULT_NETLIFY_SITE_ID, "token": DEFAULT_NETLIFY_TOKEN}
    if os.path.exists(NETLIFY_CONFIG_FILE):
        try:
            with open(NETLIFY_CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                if saved.get("site_id"):
                    cfg["site_id"] = saved["site_id"]
                if saved.get("token"):
                    cfg["token"] = saved["token"]
        except Exception:
            pass
    return cfg


def save_netlify_config(site_id, token):
    try:
        with open(NETLIFY_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"site_id": site_id.strip(), "token": token.strip()}, f, indent=2)
    except Exception:
        pass


def _rebuild_published_index_html():
    """Génère une page d'accueil propre listant toutes les fiches pilotes publiées."""
    os.makedirs(PUBLISHED_DIR, exist_ok=True)
    html_files = sorted(
        [f for f in os.listdir(PUBLISHED_DIR) if f.endswith(".html") and f != "index.html"],
        key=lambda fn: os.path.getmtime(os.path.join(PUBLISHED_DIR, fn)),
        reverse=True,
    )
    items_html = ""
    for fn in html_files:
        clean_label = fn.replace("DriverPage_", "").replace(".html", "").replace("_", " ")
        dt_str = datetime.fromtimestamp(os.path.getmtime(os.path.join(PUBLISHED_DIR, fn))).strftime("%d/%m/%Y %H:%M")
        items_html += (
            f'<a href="./{urllib.parse.quote(fn)}" class="item">'
            f'<div><div class="item-title">{clean_label}</div><div class="item-sub">Publié le {dt_str}</div></div>'
            f'<div class="item-arrow">Ouvrir →</div></a>'
        )
    if not items_html:
        items_html = '<div style="color:#94a3b8; text-align:center; padding:24px;">Aucune page pilote publiée pour le moment.</div>'

    index_html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>GRAFF — Driver Pages</title>
<style>
  body {{ margin:0; padding:24px; background:#0b0f19; color:#f1f5f9; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; }}
  .wrap {{ max-width:720px; margin:0 auto; }}
  .hdr {{ display:flex; justify-content:space-between; align-items:center; background:#0f172a; border:1px solid #1e293b; border-radius:12px; padding:16px 20px; margin-bottom:20px; }}
  .logo {{ font-weight:900; letter-spacing:0.18em; border:1.5px solid #475569; padding:6px 12px; border-radius:6px; }}
  .item {{ display:flex; justify-content:space-between; align-items:center; background:#111827; border:1px solid #1f2937; border-radius:10px; padding:14px 18px; margin-bottom:10px; text-decoration:none; color:#f8fafc; }}
  .item:hover {{ border-color:#38bdf8; background:#162033; }}
  .item-title {{ font-weight:700; font-size:1rem; }}
  .item-sub {{ font-size:0.8rem; color:#94a3b8; margin-top:3px; }}
  .item-arrow {{ color:#38bdf8; font-weight:700; font-size:0.9rem; }}
</style>
</head>
<body>
<div class="wrap">
  <div class="hdr">
    <div><div style="font-size:1.25rem; font-weight:800;">GRAFF — DRIVER DEBRIEF</div><div style="font-size:0.82rem; color:#94a3b8;">Sélectionne ta fiche télémétrie</div></div>
    <div class="logo">GRAFF</div>
  </div>
  {items_html}
</div>
</body>
</html>"""
    with open(os.path.join(PUBLISHED_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)


def publish_to_netlify(filename, html_content, site_id, token):
    """
    Publie la page HTML sur Netlify via l'API File Digest sans écraser les fichiers précédents.
    Retourne (ok: bool, url_or_error: str).
    """
    site_id = (site_id or "").strip()
    token = (token or "").strip()
    if not site_id or not token:
        return False, "Site ID ou Token Netlify manquant dans la barre latérale."

    os.makedirs(PUBLISHED_DIR, exist_ok=True)
    target_path = os.path.join(PUBLISHED_DIR, filename)
    with open(target_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    _rebuild_published_index_html()

    files_sha1 = {}
    files_bytes = {}
    for fname in os.listdir(PUBLISHED_DIR):
        fpath = os.path.join(PUBLISHED_DIR, fname)
        if os.path.isfile(fpath) and fname.endswith(".html"):
            with open(fpath, "rb") as fb:
                raw = fb.read()
            web_path = f"/{fname}"
            files_sha1[web_path] = hashlib.sha1(raw).hexdigest()
            files_bytes[web_path] = raw

    sha_to_paths = {}
    for wpath, sha in files_sha1.items():
        sha_to_paths.setdefault(sha, []).append(wpath)

    deploy_url = f"https://api.netlify.com/api/v1/sites/{urllib.parse.quote(site_id)}/deploys"
    req_body = json.dumps({"files": files_sha1}).encode("utf-8")
    req = urllib.request.Request(
        deploy_url,
        data=req_body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "User-Agent": "GRAFF-RaceEngineering/1.0",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=25) as resp:
            deploy_info = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        err_txt = e.read().decode("utf-8", errors="replace")
        return False, f"Erreur HTTP Netlify ({e.code}) : {err_txt[:220]}"
    except Exception as e:
        return False, f"Erreur réseau Netlify : {e}"

    deploy_id = deploy_info.get("id")
    required_shas = deploy_info.get("required", []) or []

    for sha in required_shas:
        for wpath in sha_to_paths.get(sha, []):
            encoded_wpath = urllib.parse.quote(wpath, safe="/")
            up_url = f"https://api.netlify.com/api/v1/deploys/{deploy_id}/files{encoded_wpath}"
            up_req = urllib.request.Request(
                up_url,
                data=files_bytes[wpath],
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/octet-stream",
                    "User-Agent": "GRAFF-RaceEngineering/1.0",
                },
                method="PUT",
            )
            try:
                with urllib.request.urlopen(up_req, timeout=30) as up_resp:
                    up_resp.read()
            except urllib.error.HTTPError as e:
                err_txt = e.read().decode("utf-8", errors="replace")
                return False, f"Échec upload fichier ({e.code}) : {err_txt[:220]}"
            except Exception as e:
                return False, f"Échec upload fichier : {e}"

    base_site_url = deploy_info.get("ssl_url") or deploy_info.get("url") or f"https://{site_id}.netlify.app"
    final_url = f"{base_site_url.rstrip('/')}/{urllib.parse.quote(filename)}"
    return True, final_url


def load_circuit_presets():
    if os.path.exists(PRESETS_FILE):
        try:
            with open(PRESETS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_circuit_preset(circuit_name, corners_df):
    presets = load_circuit_presets()
    key = circuit_name.strip().lower()
    records = []
    for _, r in corners_df.iterrows():
        records.append(
            {
                "Turn": str(r["Turn"]).strip(),
                "ApexDist_m": float(r["ApexDist_m"]),
            }
        )
    records.sort(key=lambda x: x["ApexDist_m"])
    presets[key] = records
    with open(PRESETS_FILE, "w", encoding="utf-8") as f:
        json.dump(presets, f, ensure_ascii=False, indent=2)


def delete_circuit_preset(circuit_name):
    presets = load_circuit_presets()
    key = circuit_name.strip().lower()
    if key in presets:
        del presets[key]
        with open(PRESETS_FILE, "w", encoding="utf-8") as f:
            json.dump(presets, f, ensure_ascii=False, indent=2)


def find_graff_logo():
    """Cherche le logo GRAFF dans le dossier courant."""
    for fname in ("logo_graff.png", "graff.png", "logo_graff.jpg", "graff.jpg"):
        if os.path.exists(fname):
            return fname
    return None


def get_graff_logo_base64(logo_path=None):
    """Retourne une data URI base64 pour intégrer le logo GRAFF en HTML/Streamlit."""
    if not logo_path or not os.path.exists(logo_path):
        return None
    ext = os.path.splitext(logo_path)[1].lower()
    mime = "image/jpeg" if ext in (".jpg", ".jpeg") else "image/png"
    try:
        with open(logo_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        return f"data:{mime};base64,{b64}"
    except Exception:
        return None


def render_graff_banner(title, subtitle, logo_path=None):
    """Affiche l'entête GRAFF (bandeau sombre + logo en haut à droite) dans Streamlit."""
    logo_uri = get_graff_logo_base64(logo_path)
    if logo_uri:
        right_html = f'<img src="{logo_uri}" alt="GRAFF" style="height:44px; width:auto; object-fit:contain;">'
    else:
        right_html = '<div class="graff-logo-fallback">GRAFF</div>'

    st.markdown(
        f"""
        <div class="graff-banner">
            <div>
                <div class="graff-banner-title">{title}</div>
                <div class="graff-banner-sub">{subtitle}</div>
            </div>
            <div>{right_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_clean_table(df_to_show, cols, highlight_cars=None, max_height=420, fixed_height=None):
    """Affiche un tableau HTML stylisé avec surlignage rose pour notre voiture."""
    highlight_set = {c.strip() for c in (highlight_cars or []) if c.strip()}
    rows_html = []
    headers_html = "".join(f"<th>{c}</th>" for c in cols)
    for idx, r in df_to_show.iterrows():
        car_val = str(r["N°"]).strip() if "N°" in df_to_show.columns else ""
        if car_val in highlight_set:
            bg_style = 'style="background-color: rgba(236, 72, 153, 0.28); font-weight: 700;"'
        else:
            bg_style = ""
        tds = "".join(f"<td>{r[c]}</td>" for c in cols)
        rows_html.append(f"<tr {bg_style}>{tds}</tr>")

    height_css = f"height: {fixed_height}px;" if fixed_height is not None else f"max-height: {max_height}px;"
    html_tbl = f"<table class='dataframe'><thead><tr>{headers_html}</tr></thead><tbody>{''.join(rows_html)}</tbody></table>"
    st.markdown(
        f"""
        <div style="{height_css} overflow-y: auto; border: 1px solid #2d3748; border-radius: 8px; margin-bottom: 14px; box-sizing: border-box;">
        <style>
            table.dataframe {{ width: 100%; border-collapse: collapse; font-family: Consolas, sans-serif; font-size: 0.9rem; }}
            table.dataframe th {{ background-color: #1e293b; color: #f8fafc; padding: 10px; text-align: left; position: sticky; top: 0; z-index: 1; }}
            table.dataframe td {{ padding: 8px 10px; border-bottom: 1px solid #1f2937; color: #e2e8f0; }}
            table.dataframe tr:nth-child(even) {{ background-color: #111827; }}
        </style>
        {html_tbl}
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# FONCTIONS UTILITAIRES : TEMPS, PARSING ORGANISATEUR, MOTEC (.CSV) & VBOX (.VBO)
# -----------------------------------------------------------------------------
def parse_time_str(t_str):
    if t_str is None:
        return None
    s = str(t_str).strip()
    if s in ("", "-", "Pit In", "None", "nan", "NaN"):
        return None
    parts = s.split(":")
    try:
        if len(parts) == 3:
            return float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
        elif len(parts) == 2:
            return float(parts[0]) * 60 + float(parts[1])
        return float(parts[0])
    except ValueError:
        return None


def format_laptime(sec):
    if sec is None or pd.isna(sec):
        return "-"
    sign = "-" if sec < 0 else ""
    sec = abs(float(sec))
    mins = int(sec // 60)
    rem = sec - mins * 60
    if mins > 0:
        return f"{sign}{mins}:{rem:06.3f}"
    return f"{sign}{rem:.3f}"


def _pick_channel(df, candidates, keyword_fallbacks=()):
    """Trouve la colonne télémétrie correspondante (insensible à la casse/espaces/underscores)."""
    cols_map = {c.strip().lower(): c for c in df.columns}
    cols_norm_map = {re.sub(r"[^a-z0-9]", "", c.lower()): c for c in df.columns}

    for cand in candidates:
        key = cand.strip().lower()
        if key in cols_map and df[cols_map[key]].notna().sum() > 10:
            return cols_map[key]
        key_n = re.sub(r"[^a-z0-9]", "", key)
        if key_n in cols_norm_map and df[cols_norm_map[key_n]].notna().sum() > 10:
            return cols_norm_map[key_n]

    for kw in keyword_fallbacks:
        kw_l = kw.lower()
        for col_l, orig_col in cols_map.items():
            if kw_l in col_l and df[orig_col].notna().sum() > 10:
                return orig_col
    return None


def parse_lap_by_lap(uploaded_file):
    """Parse le fichier Excel (.xlsx) ou CSV Lap-by-Lap de l'organisateur."""
    name = uploaded_file.name.lower()
    rows = []
    if name.endswith(".xlsx") or name.endswith(".xls"):
        wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
    else:
        content = uploaded_file.getvalue().decode("utf-8", errors="replace").splitlines()
        sep = ";" if ";" in content[0] else ","
        rows = [line.split(sep) for line in content]

    records = []
    curr_car, curr_driver, curr_cat = None, None, None
    for r in rows:
        if not r:
            continue
        c0 = str(r[0]).strip() if r[0] is not None else ""
        if c0.startswith("No."):
            m = re.match(r"No\.\s*([^,]+),\s*([^(]+)\(([^)]+)\)", c0)
            if m:
                curr_car = m.group(1).strip()
                curr_driver = m.group(2).strip()
                curr_cat = m.group(3).strip()
            else:
                curr_car = c0
                curr_driver = c0
                curr_cat = "?"
        elif curr_car is not None and c0.isdigit():
            lap_num = int(c0)
            lt_raw = str(r[7]).strip() if len(r) > 7 and r[7] is not None else ""
            lt_sec = parse_time_str(lt_raw)
            records.append(
                {
                    "Car": curr_car,
                    "Driver": curr_driver,
                    "Category": curr_cat,
                    "Lap": lap_num,
                    "TimeOfDay": str(r[1]).strip() if len(r) > 1 and r[1] is not None else "",
                    "S1": parse_time_str(r[2]) if len(r) > 2 else None,
                    "S2": parse_time_str(r[3]) if len(r) > 3 else None,
                    "S3": parse_time_str(r[4]) if len(r) > 4 else None,
                    "IsPitIn": ("Pit" in lt_raw),
                    "LapTimeRaw": lt_raw,
                    "LapTimeSec": lt_sec,
                }
            )

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records).sort_values(["Car", "Lap"]).reset_index(drop=True)
    df["PrevPit"] = df.groupby("Car")["IsPitIn"].shift(1, fill_value=True)
    df["PrevLap"] = df.groupby("Car")["Lap"].shift(1, fill_value=0)
    df["IsOutLap"] = (df["Lap"] == 1) | (df["PrevPit"]) | ((df["Lap"] - df["PrevLap"]) > 1)

    valid_timed = df.loc[(~df["IsPitIn"]) & (~df["IsOutLap"]) & (df["LapTimeSec"].notna()), "LapTimeSec"]
    session_best = valid_timed.min() if len(valid_timed) > 0 else 999.0
    df["IsClean"] = (
        (~df["IsPitIn"])
        & (~df["IsOutLap"])
        & (df["LapTimeSec"].notna())
        & (df["LapTimeSec"] <= session_best * 1.07)
    )
    return df


def compute_all_rankings(df_lbl):
    """Calcule tous les classements demandés (Best & Moy 50% pour Pilotes, Voitures, Silver et Bronze)."""
    driver_rows = []
    for (car, driver, cat), grp in df_lbl.groupby(["Car", "Driver", "Category"]):
        clean = grp[grp["IsClean"]].sort_values("LapTimeSec")
        if len(clean) == 0:
            clean = grp[grp["LapTimeSec"].notna()].sort_values("LapTimeSec")
        if len(clean) == 0:
            continue

        best_sec = float(clean["LapTimeSec"].iloc[0])
        n_top50 = max(1, math.ceil(len(clean) * 0.5))
        top50_sec = float(clean["LapTimeSec"].iloc[:n_top50].mean())
        short_name = driver.split()[-1] if len(driver.split()) > 1 else driver

        driver_rows.append(
            {
                "N°": car,
                "Pilote": driver,
                "ShortName": short_name,
                "Cat.": cat,
                "Tours Valides": int(len(clean)),
                "Tours Top 50%": int(n_top50),
                "BestSec": best_sec,
                "Meilleur Tour": format_laptime(best_sec),
                "Top50Sec": top50_sec,
                "Moy. 50% Meilleurs": format_laptime(top50_sec),
            }
        )

    df_base = pd.DataFrame(driver_rows)
    if df_base.empty:
        empty = pd.DataFrame()
        return empty, empty, empty, empty, empty, empty, empty, empty

    def make_best_table(sub_df):
        if sub_df.empty:
            return sub_df
        t = sub_df.sort_values("BestSec").reset_index(drop=True).copy()
        t.insert(0, "Pos", t.index + 1)
        leader = t["BestSec"].iloc[0]
        t["Écart"] = t["BestSec"].apply(lambda x: f"+{x - leader:.3f} s" if x > leader else "Leader")
        return t

    def make_avg_table(sub_df):
        if sub_df.empty:
            return sub_df
        t = sub_df.sort_values("Top50Sec").reset_index(drop=True).copy()
        t.insert(0, "Pos", t.index + 1)
        leader = t["Top50Sec"].iloc[0]
        t["Écart"] = t["Top50Sec"].apply(lambda x: f"+{x - leader:.3f} s" if x > leader else "Leader")
        return t

    df_pilotes_best = make_best_table(df_base)
    df_pilotes_avg = make_avg_table(df_base)

    df_silver_base = df_base[df_base["Cat."] == "S"]
    df_silver_best = make_best_table(df_silver_base)
    df_silver_avg = make_avg_table(df_silver_base)

    df_bronze_base = df_base[df_base["Cat."] == "B"]
    df_bronze_best = make_best_table(df_bronze_base)
    df_bronze_avg = make_avg_table(df_bronze_base)

    car_rows = []
    for car, grp in df_base.groupby("N°"):
        car_best = float(grp["BestSec"].min())
        car_avg_of_avgs = float(grp["Top50Sec"].mean())
        equip_str = " / ".join(grp["Pilote"].tolist())
        cats_str = "-".join(grp["Cat."].tolist())
        detail_parts = []
        for _, r in grp.iterrows():
            star = "*" if r["Tours Valides"] <= 3 else ""
            detail_parts.append(f"{r['ShortName']} {r['Moy. 50% Meilleurs']} ({r['Tours Top 50%']}{star})")
        detail_str = " / ".join(detail_parts)

        car_rows.append(
            {
                "N°": car,
                "Équipage": equip_str,
                "Cat.": cats_str,
                "BestSec": car_best,
                "Meilleur tour": format_laptime(car_best),
                "AvgSec": car_avg_of_avgs,
                "Moy. des moy. 50%": format_laptime(car_avg_of_avgs),
                "Détail : moy. 50% par pilote (tours)": detail_str,
            }
        )

    df_cars_base = pd.DataFrame(car_rows)

    df_cars_best = df_cars_base.sort_values("BestSec").reset_index(drop=True).copy()
    df_cars_best.insert(0, "Pos", df_cars_best.index + 1)
    leader_car_best = df_cars_best["BestSec"].iloc[0]
    df_cars_best["Écart"] = df_cars_best["BestSec"].apply(
        lambda x: f"+{x - leader_car_best:.3f} s" if x > leader_car_best else "Leader"
    )
    car_to_best_rank = dict(zip(df_cars_best["N°"], df_cars_best["Pos"]))

    df_cars_avg = df_cars_base.sort_values("AvgSec").reset_index(drop=True).copy()
    df_cars_avg.insert(0, "Pos", df_cars_avg.index + 1)
    df_cars_avg["Rang tour"] = df_cars_avg["N°"].map(car_to_best_rank)
    leader_car_avg = df_cars_avg["AvgSec"].iloc[0]
    df_cars_avg["Écart"] = df_cars_avg["AvgSec"].apply(
        lambda x: f"+{x - leader_car_avg:.3f} s" if x > leader_car_avg else "Leader"
    )

    return (
        df_pilotes_best,
        df_pilotes_avg,
        df_cars_best,
        df_cars_avg,
        df_silver_best,
        df_silver_avg,
        df_bronze_best,
        df_bronze_avg,
    )


def parse_motec_csv(uploaded_file):
    """Parse un export CSV MoTeC et découpe les tours via Beacon Markers."""
    raw_bytes = uploaded_file.getvalue()
    text_lines = raw_bytes.decode("utf-8", errors="replace").splitlines()

    meta = {"Format": "MoTeC CSV"}
    header_idx = 14
    for idx, line in enumerate(text_lines[:25]):
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) >= 2 and parts[0]:
            meta[parts[0]] = parts[1]
        if len(parts) >= 5 and parts[4]:
            meta[parts[4]] = parts[5] if len(parts) > 5 else ""
        if parts[0] == "Time":
            header_idx = idx
            break

    df = pd.read_csv(io.BytesIO(raw_bytes), skiprows=header_idx, low_memory=False)
    df = df.iloc[1:].dropna(subset=["Time"]).copy()
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["Time"]).reset_index(drop=True)

    speed_col = _pick_channel(df, ["Corr Speed", "Vehicle Speed", "Ground Speed", "Speed"], ("speed",))
    if speed_col and speed_col != "Corr Speed":
        df["Corr Speed"] = df[speed_col]

    beacons_str = meta.get("Beacon Markers", "")
    beacons = [float(x) for x in beacons_str.split() if x.replace(".", "", 1).isdigit()]

    if len(beacons) >= 2:
        df["BeaconLapIdx"] = np.searchsorted(beacons, df["Time"].values, side="right")
    else:
        unique_laps = df["Lap Number"].dropna().unique()
        lap_map = {val: i for i, val in enumerate(unique_laps)}
        df["BeaconLapIdx"] = df["Lap Number"].map(lap_map).fillna(0).astype(int)

    laps_info = []
    for b_idx, grp in df.groupby("BeaconLapIdx"):
        if b_idx == 0 or b_idx >= len(beacons):
            continue
        t_start = beacons[b_idx - 1]
        t_end = beacons[b_idx]
        lap_dur = t_end - t_start
        dt = np.gradient(grp["Time"].values)
        spd_ms = grp["Corr Speed"].fillna(0).values / 3.6
        dist_m = float(np.sum(spd_ms * dt))
        laps_info.append(
            {
                "BeaconLapIdx": int(b_idx),
                "MotecLapNum": int(grp["Lap Number"].iloc[0])
                if "Lap Number" in grp.columns and pd.notna(grp["Lap Number"].iloc[0])
                else int(b_idx),
                "BeaconStart": float(t_start),
                "BeaconEnd": float(t_end),
                "LapTimeSec": float(lap_dur),
                "LapTimeStr": format_laptime(lap_dur),
                "DistKm": dist_m / 1000.0,
            }
        )

    return meta, df, pd.DataFrame(laps_info)


def _vbo_hhmmss_to_sec(val):
    """Convertit un timestamp VBOX HHMMSS.sss en secondes depuis minuit."""
    v = float(val)
    hh = int(v // 10000)
    rem = v - hh * 10000
    mm = int(rem // 100)
    ss = rem - mm * 100
    return hh * 3600.0 + mm * 60.0 + ss


def _seg_intersect_2d(ax, ay, bx, by, cx, cy, dx, dy):
    """Retourne le paramètre t in [0, 1] si le segment AB coupe le segment CD, sinon None."""
    rx, ry = bx - ax, by - ay
    sx, sy = dx - cx, dy - cy
    denom = rx * sy - ry * sx
    if abs(denom) < 1e-15:
        return None
    qpx, qpy = cx - ax, cy - ay
    t = (qpx * sy - qpy * sx) / denom
    u = (qpx * ry - qpy * rx) / denom
    if 0.0 <= t <= 1.0 and 0.0 <= u <= 1.0:
        return t
    return None


def parse_vbo_file(uploaded_file):
    """
    Parse un fichier Racelogic VBOX (.vbo) (ex: Porsche Cup VBOX HD2 / Video VBOX)
    et retourne (meta, df, df_laps) avec exactement le même schéma que parse_motec_csv.
    """
    raw_bytes = uploaded_file.getvalue()
    text_lines = raw_bytes.decode("utf-8", errors="replace").splitlines()

    meta = {"Format": "Racelogic VBOX (.vbo)"}
    sections = {}
    curr_sec = None
    for idx, line in enumerate(text_lines):
        s = line.strip()
        if s.startswith("[") and s.endswith("]"):
            curr_sec = s.lower()
            sections[curr_sec] = []
        elif curr_sec is not None:
            sections[curr_sec].append(line)
        elif idx < 5 and s:
            meta["HeaderInfo"] = s

    col_names = []
    if "[column names]" in sections:
        for l in sections["[column names]"]:
            if l.strip():
                col_names = l.strip().split()
                break

    if not col_names and "[header]" in sections:
        header_map = {
            "satellites": "sats",
            "time": "time",
            "latitude": "lat",
            "longitude": "long",
            "velocity kmh": "velocity",
            "heading": "heading",
            "height": "height",
            "vertical velocity m/s": "vert-vel",
            "sampleperiod": "Tsample",
            "solution type": "solution_type",
            "avifileindex": "avifileindex",
            "avisynctime": "avitime",
        }
        for l in sections["[header]"]:
            k = l.strip()
            if k:
                col_names.append(header_map.get(k.lower(), k.replace(" ", "_")))

    data_lines = [l for l in sections.get("[data]", []) if l.strip()]
    if not data_lines or not col_names:
        return meta, pd.DataFrame(), pd.DataFrame()

    df = pd.read_csv(
        io.StringIO("\n".join(data_lines)),
        sep=r"\s+",
        names=col_names[: len(data_lines[0].strip().split())],
        engine="python",
    )
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    if "time" not in df.columns:
        return meta, pd.DataFrame(), pd.DataFrame()

    df = df.dropna(subset=["time"]).reset_index(drop=True)
    t_arr = np.array([_vbo_hhmmss_to_sec(v) for v in df["time"].values], dtype=float)
    for i in range(1, len(t_arr)):
        if t_arr[i] < t_arr[i - 1] - 40000.0:
            t_arr[i:] += 86400.0
    t_origin = float(t_arr[0])
    df["Time"] = t_arr - t_origin

    raw_lat = df["lat"].values.copy() if "lat" in df.columns else np.zeros(len(df))
    raw_lon = df["long"].values.copy() if "long" in df.columns else np.zeros(len(df))
    if np.nanmedian(np.abs(raw_lat)) > 90.0:
        df["GPS Latitude"] = raw_lat / 60.0
        df["GPS Longitude"] = -raw_lon / 60.0
    else:
        df["GPS Latitude"] = raw_lat
        df["GPS Longitude"] = raw_lon

    spd_col = _pick_channel(df, ["velocity", "velocity_kmh", "speed"], ("vel", "speed"))
    df["Corr Speed"] = df[spd_col] if spd_col else 0.0

    beacons = []
    time_vals = df["Time"].values
    spd_vals = df["Corr Speed"].fillna(0).values

    if "[laptiming]" in sections and len(raw_lat) > 10:
        for l in sections["[laptiming]"]:
            if l.strip().lower().startswith("start"):
                nums = re.findall(r"[+-]?\d+(?:\.\d+)?", l)
                if len(nums) >= 4:
                    c1, c2, c3, c4 = float(nums[0]), float(nums[1]), float(nums[2]), float(nums[3])
                    med_lat = float(np.nanmedian(raw_lat))
                    med_lon = float(np.nanmedian(raw_lon))
                    if abs(c1 - med_lon) + abs(c2 - med_lat) <= abs(c1 - med_lat) + abs(c2 - med_lon):
                        gx1, gy1, gx2, gy2 = c1, c2, c3, c4
                    else:
                        gy1, gx1, gy2, gx2 = c1, c2, c3, c4

                    mx, my = (gx1 + gx2) / 2.0, (gy1 + gy2) / 2.0
                    wx1, wy1 = mx + (gx1 - mx) * 2.5, my + (gy1 - my) * 2.5
                    wx2, wy2 = mx + (gx2 - mx) * 2.5, my + (gy2 - my) * 2.5

                    raw_crossings = []
                    for i in range(len(df) - 1):
                        if np.isnan(raw_lon[i]) or np.isnan(raw_lon[i + 1]):
                            continue
                        frac = _seg_intersect_2d(
                            raw_lon[i], raw_lat[i], raw_lon[i + 1], raw_lat[i + 1], wx1, wy1, wx2, wy2
                        )
                        if frac is not None:
                            tc = float(time_vals[i] + frac * (time_vals[i + 1] - time_vals[i]))
                            sp = float(spd_vals[i])
                            if not raw_crossings or (tc - raw_crossings[-1][0] > 20.0):
                                raw_crossings.append((tc, sp))

                    track_crossings = [tc for tc, sp in raw_crossings if sp > 65.0]
                    if len(track_crossings) >= 2:
                        beacons = track_crossings
                    elif len(raw_crossings) >= 2:
                        beacons = [tc for tc, _ in raw_crossings]
                break

    if len(beacons) < 2 and "Lap_Time" in df.columns:
        lt_vals = df["Lap_Time"].fillna(0).values
        for i in range(1, len(lt_vals)):
            if lt_vals[i - 1] > 20.0 and lt_vals[i] < 2.0:
                tc = float(time_vals[i])
                if not beacons or (tc - beacons[-1] > 20.0):
                    beacons.append(tc)

    if len(beacons) < 2 and len(df) > 100:
        fast_idx = np.where(spd_vals > 110.0)[0]
        if len(fast_idx) > 0:
            ref_i = int(fast_idx[len(fast_idx) // 2])
            lat_m = df["GPS Latitude"].values * 110540.0
            lon_m = df["GPS Longitude"].values * np.cos(np.deg2rad(np.nanmean(df["GPS Latitude"].values))) * 111320.0
            d_ref = np.hypot(lon_m - lon_m[ref_i], lat_m - lat_m[ref_i])
            in_zone = False
            best_local_t = None
            best_local_d = 999.0
            for i in range(len(d_ref)):
                if d_ref[i] < 22.0:
                    in_zone = True
                    if d_ref[i] < best_local_d:
                        best_local_d = d_ref[i]
                        best_local_t = float(time_vals[i])
                elif in_zone:
                    if best_local_t is not None and (not beacons or best_local_t - beacons[-1] > 30.0):
                        beacons.append(best_local_t)
                    in_zone = False
                    best_local_t = None
                    best_local_d = 999.0

    if len(beacons) < 2 and len(df) > 40:
        beacons = [float(time_vals[0]), float(time_vals[-1])]

    df["BeaconLapIdx"] = np.searchsorted(beacons, df["Time"].values, side="right")

    laps_info = []
    for b_idx, grp in df.groupby("BeaconLapIdx"):
        if b_idx == 0 or b_idx >= len(beacons):
            continue
        t_start = beacons[b_idx - 1]
        t_end = beacons[b_idx]
        lap_dur = t_end - t_start
        if lap_dur <= 15.0:
            continue
        dt = np.gradient(grp["Time"].values)
        spd_ms = grp["Corr Speed"].fillna(0).values / 3.6
        dist_m = float(np.sum(spd_ms * dt))
        laps_info.append(
            {
                "BeaconLapIdx": int(b_idx),
                "MotecLapNum": int(b_idx),
                "BeaconStart": float(t_start),
                "BeaconEnd": float(t_end),
                "LapTimeSec": float(lap_dur),
                "LapTimeStr": format_laptime(lap_dur),
                "DistKm": dist_m / 1000.0,
            }
        )

    return meta, df, pd.DataFrame(laps_info)


def parse_telemetry_file(uploaded_file):
    """Détecte automatiquement le format (.vbo Racelogic ou .csv MoTeC) et parse le fichier."""
    fname = uploaded_file.name.lower()
    if fname.endswith(".vbo"):
        return parse_vbo_file(uploaded_file)
    head_preview = uploaded_file.getvalue()[:300].decode("utf-8", errors="replace")
    if "[header]" in head_preview.lower() or "racelogic" in head_preview.lower():
        return parse_vbo_file(uploaded_file)
    return parse_motec_csv(uploaded_file)


def align_motec_with_lbl_and_stints(
    df_motec_laps, df_lbl, car_num, fuel_start, fuel_cons_per_lap, refuel_events, tyre_start_km, tyre_events
):
    df_m = df_motec_laps.copy()
    if df_m.empty:
        return df_m

    car_lbl = pd.DataFrame()
    if df_lbl is not None and not df_lbl.empty and str(car_num).strip() != "":
        car_lbl = df_lbl[df_lbl["Car"] == str(car_num).strip()].sort_values("Lap").reset_index(drop=True)

    official_laps = []
    drivers = []
    used_lbl_indices = set()

    for _, row in df_m.iterrows():
        matched_lap = None
        matched_driver = "Pilote Voiture"
        if not car_lbl.empty:
            diffs = (car_lbl["LapTimeSec"] - row["LapTimeSec"]).abs()
            valid_candidates = diffs[(diffs < 1.8) & (~diffs.index.isin(used_lbl_indices))]
            if not valid_candidates.empty:
                best_idx = valid_candidates.idxmin()
                used_lbl_indices.add(best_idx)
                matched_lap = int(car_lbl.loc[best_idx, "Lap"])
                matched_driver = str(car_lbl.loc[best_idx, "Driver"])

        official_laps.append(matched_lap)
        drivers.append(matched_driver)

    for i in range(len(official_laps)):
        if official_laps[i] is None:
            if i > 0 and official_laps[i - 1] is not None:
                official_laps[i] = official_laps[i - 1] + 1
            else:
                official_laps[i] = i + 1

    if not car_lbl.empty:
        lap_to_driver = dict(zip(car_lbl["Lap"], car_lbl["Driver"]))
        for i, olap in enumerate(official_laps):
            if olap in lap_to_driver:
                drivers[i] = lap_to_driver[olap]

    df_m["OfficialLap"] = official_laps
    df_m["Driver"] = drivers

    refuel_dict = {
        int(r["Tour"]): float(r["Litres ajoutés"])
        for r in refuel_events
        if pd.notna(r.get("Tour")) and pd.notna(r.get("Litres ajoutés"))
    }
    tyre_dict = {
        int(t["Tour"]): float(t["Nouveau km départ"])
        for t in tyre_events
        if pd.notna(t.get("Tour")) and pd.notna(t.get("Nouveau km départ"))
    }

    curr_fuel = float(fuel_start)
    curr_tyre_km = float(tyre_start_km)
    fuels, tyres = [], []

    for _, row in df_m.iterrows():
        olap = int(row["OfficialLap"])
        if olap in refuel_dict:
            curr_fuel += refuel_dict[olap]
        if olap in tyre_dict:
            curr_tyre_km = tyre_dict[olap]

        fuels.append(round(max(0.0, curr_fuel), 1))
        tyres.append(round(curr_tyre_km, 1))

        curr_fuel -= float(fuel_cons_per_lap)
        lap_dist = row["DistKm"] if row["DistKm"] > 1.0 else 4.65
        curr_tyre_km += lap_dist

    df_m["Fuel_L"] = fuels
    df_m["Tyre_km"] = tyres
    return df_m


# -----------------------------------------------------------------------------
# DÉTECTION INTELLIGENTE DES VIRAGES (APEX SIGNÉS + PRESET CIRCUIT)
# ET EXTRACTION DES CANAUX DE TÉLÉMÉTRIE (VITESSE, GAZ, FREIN, RAPPORT, VOLANT, RPM, GPS)
# -----------------------------------------------------------------------------
def _find_braking_or_lift_point(grid_d, throttle_b, brake_b, speed_b, corner_entry_idx, prev_apex_idx=0):
    entry_d = grid_d[corner_entry_idx]
    min_d = max(grid_d[prev_apex_idx] + 15.0, entry_d - 320.0)
    search_start = int(np.searchsorted(grid_d, min_d))
    if search_start >= corner_entry_idx:
        return max(prev_apex_idx + 1, corner_entry_idx - 3)

    if brake_b is not None:
        b_seg = brake_b[search_start : corner_entry_idx + 1]
        if len(b_seg) > 3 and np.nanmax(b_seg) > 5.0:
            active = np.where(b_seg > 3.0)[0]
            if len(active) > 0:
                return max(search_start, search_start + int(active[0]) - 2)

    if throttle_b is not None:
        t_seg = throttle_b[search_start : corner_entry_idx + 1]
        if len(t_seg) > 3 and np.nanmax(t_seg) > 85.0:
            idx = corner_entry_idx
            while idx > search_start and throttle_b[idx] < 95.0:
                idx -= 1
            if idx < corner_entry_idx:
                return max(search_start, idx)

    s_seg = speed_b[search_start : corner_entry_idx + 1]
    if len(s_seg) > 2:
        peak_rel = int(np.argmax(s_seg))
        return max(search_start, search_start + peak_rel)

    return max(search_start, corner_entry_idx - 8)


def _detect_signed_apexes(steer_signed, steer_thresh, min_sep_pts=14):
    n = len(steer_signed)
    smooth = pd.Series(steer_signed).rolling(11, center=True, min_periods=1).mean().values
    abs_s = np.abs(smooth)

    raw_peaks = []
    for i in range(4, n - 4):
        if abs_s[i] >= steer_thresh:
            if abs_s[i] >= np.max(abs_s[i - 4 : i + 5]):
                raw_peaks.append(i)

    filtered_peaks = []
    for p in raw_peaks:
        if not filtered_peaks:
            filtered_peaks.append(p)
            continue
        prev = filtered_peaks[-1]
        same_sign = (smooth[p] * smooth[prev]) > 0
        valley = float(np.min(abs_s[prev : p + 1]))
        peak_min = min(abs_s[prev], abs_s[p])

        if not same_sign:
            if p - prev >= 6:
                filtered_peaks.append(p)
            elif abs_s[p] > abs_s[prev]:
                filtered_peaks[-1] = p
        else:
            if (p - prev >= min_sep_pts) and (valley < peak_min * 0.72 or (p - prev >= min_sep_pts * 1.6)):
                filtered_peaks.append(p)
            elif abs_s[p] > abs_s[prev]:
                filtered_peaks[-1] = p

    segments = []
    for idx, pk in enumerate(filtered_peaks):
        left_bound = 0 if idx == 0 else (filtered_peaks[idx - 1] + pk) // 2
        right_bound = (n - 1) if idx == len(filtered_peaks) - 1 else (pk + filtered_peaks[idx + 1]) // 2

        s = pk
        while s > left_bound and abs_s[s] > steer_thresh * 0.65:
            s -= 1
        e = pk
        while e < right_bound and abs_s[e] > steer_thresh * 0.65:
            e += 1
        segments.append((s, pk, e))

    return segments


def _interp_channel_nearest(grid_norm, norm_d, series):
    """Interpolation plus proche voisin (idéale pour les rapports engagés Gear)."""
    idx = np.searchsorted(norm_d, grid_norm, side="left")
    idx = np.clip(idx, 0, len(series) - 1)
    return series[idx]


def _compute_heading_deg(lat_arr, lon_arr):
    d_lon = np.gradient(lon_arr) * np.cos(np.deg2rad(lat_arr))
    d_lat = np.gradient(lat_arr)
    hdg = (np.rad2deg(np.arctan2(d_lon, d_lat)) + 360.0) % 360.0
    return hdg


def compute_corner_analysis(
    df_motec_a,
    lap_row_a,
    df_motec_b,
    lap_row_b,
    steer_thresh=12.0,
    custom_corners=None,
):
    beacon_idx_a = int(lap_row_a["BeaconLapIdx"])
    beacon_idx_b = int(lap_row_b["BeaconLapIdx"])

    lap_a = df_motec_a[df_motec_a["BeaconLapIdx"] == beacon_idx_a].copy().reset_index(drop=True)
    lap_b = df_motec_b[df_motec_b["BeaconLapIdx"] == beacon_idx_b].copy().reset_index(drop=True)

    if len(lap_a) < 20 or len(lap_b) < 20:
        return None

    for lp, lrow in ((lap_a, lap_row_a), (lap_b, lap_row_b)):
        t_s = float(lrow["BeaconStart"])
        t_e = float(lrow["BeaconEnd"])
        dur = t_e - t_s

        dt = np.gradient(lp["Time"].values)
        ds = (lp["Corr Speed"].fillna(0).values / 3.6) * dt
        raw_d = np.cumsum(ds) - ds[0]

        v_first = float(lp["Corr Speed"].iloc[0]) / 3.6
        v_last = float(lp["Corr Speed"].iloc[-1]) / 3.6
        d_pre = max(0.0, (float(lp["Time"].iloc[0]) - t_s) * v_first)
        d_post = max(0.0, (t_e - float(lp["Time"].iloc[-1])) * v_last)

        full_d = raw_d + d_pre
        tot_d = float(full_d[-1] + d_post) if (full_d[-1] + d_post) > 0 else 1.0

        lp["lap_t"] = lp["Time"].values - t_s
        lp["norm_d"] = full_d / tot_d
        lp["tot_d"] = tot_d
        lp["exact_dur"] = dur

    n_pts = 1500
    grid_norm = np.linspace(0.0, 1.0, n_pts)
    ref_len_m = float(lap_b["tot_d"].iloc[0])
    grid_d = grid_norm * ref_len_m

    t_a = np.interp(
        grid_norm,
        np.concatenate([[0.0], lap_a["norm_d"].values, [1.0]]),
        np.concatenate([[0.0], lap_a["lap_t"].values, [float(lap_a["exact_dur"].iloc[0])]]),
    )
    t_b = np.interp(
        grid_norm,
        np.concatenate([[0.0], lap_b["norm_d"].values, [1.0]]),
        np.concatenate([[0.0], lap_b["lap_t"].values, [float(lap_b["exact_dur"].iloc[0])]]),
    )
    delta_t = t_a - t_b

    lat_a = np.interp(grid_norm, lap_a["norm_d"], lap_a["GPS Latitude"])
    lon_a = np.interp(grid_norm, lap_a["norm_d"], lap_a["GPS Longitude"])
    lat = np.interp(grid_norm, lap_b["norm_d"], lap_b["GPS Latitude"])
    lon = np.interp(grid_norm, lap_b["norm_d"], lap_b["GPS Longitude"])
    speed_a = np.interp(grid_norm, lap_a["norm_d"], lap_a["Corr Speed"])
    speed_b = np.interp(grid_norm, lap_b["norm_d"], lap_b["Corr Speed"])

    # --- Extraction des canaux de pilotage (A = Pilote, B = Référence) ---
    brake_cands = [
        "Press_Brake_Ft",
        "Brake_Pressure_Front",
        "Brake Pressure Front",
        "Brake Pressure",
        "Brake Pos",
        "Brake",
        "P_Brake_F",
        "pbrake_f",
    ]
    brake_col_a = _pick_channel(lap_a, brake_cands, ("brake",))
    brake_col_b = _pick_channel(lap_b, brake_cands, ("brake",))
    brake_a = (
        np.interp(grid_norm, lap_a["norm_d"], lap_a[brake_col_a].fillna(0)) if brake_col_a else np.zeros(n_pts)
    )
    brake_b = (
        np.interp(grid_norm, lap_b["norm_d"], lap_b[brake_col_b].fillna(0)) if brake_col_b else np.zeros(n_pts)
    )
    brake_unit = "bar" if (brake_col_b and ("press" in brake_col_b.lower() or "p_" in brake_col_b.lower())) else "%"

    throt_cands = [
        "%_Throttle_Pedal",
        "Throttle_Pedal",
        "Throttle Pedal",
        "Throttle Pos",
        "Throttle",
        "TPS",
        "Accel Pedal",
        "aps",
    ]
    throt_col_a = _pick_channel(lap_a, throt_cands, ("throttle", "tps", "pedal"))
    throt_col_b = _pick_channel(lap_b, throt_cands, ("throttle", "tps", "pedal"))
    throttle_a = (
        np.interp(grid_norm, lap_a["norm_d"], lap_a[throt_col_a].fillna(0)) if throt_col_a else np.zeros(n_pts)
    )
    throttle_b = (
        np.interp(grid_norm, lap_b["norm_d"], lap_b[throt_col_b].fillna(0)) if throt_col_b else np.zeros(n_pts)
    )

    gear_col_a = _pick_channel(lap_a, ["Gear", "Gear Pos", "Gear_Pos", "Current Gear"], ("gear",))
    gear_col_b = _pick_channel(lap_b, ["Gear", "Gear Pos", "Gear_Pos", "Current Gear"], ("gear",))
    gear_a = (
        np.round(_interp_channel_nearest(grid_norm, lap_a["norm_d"].values, lap_a[gear_col_a].ffill().fillna(0).values))
        if gear_col_a
        else None
    )
    gear_b = (
        np.round(_interp_channel_nearest(grid_norm, lap_b["norm_d"].values, lap_b[gear_col_b].ffill().fillna(0).values))
        if gear_col_b
        else None
    )

    rpm_cands = ["Engine Speed", "Engine_Speed", "RPM", "Engine RPM", "rpm", "N_Engine"]
    rpm_col_a = _pick_channel(lap_a, rpm_cands, ("rpm", "engine"))
    rpm_col_b = _pick_channel(lap_b, rpm_cands, ("rpm", "engine"))
    rpm_a = np.interp(grid_norm, lap_a["norm_d"], lap_a[rpm_col_a].fillna(0)) if rpm_col_a else np.zeros(n_pts)
    rpm_b = np.interp(grid_norm, lap_b["norm_d"], lap_b[rpm_col_b].fillna(0)) if rpm_col_b else np.zeros(n_pts)

    steer_cands = ["Angl_Steering", "Steering_Angle", "Steering Angle", "Steer Angle", "Steering", "swa"]
    steer_col_a = _pick_channel(lap_a, steer_cands, ("steer",))
    steer_col_b = _pick_channel(lap_b, steer_cands, ("steer",))

    if steer_col_b:
        steer_b = np.interp(grid_norm, lap_b["norm_d"], lap_b[steer_col_b].fillna(0))
        steer_signed = steer_b
    else:
        dx = np.gradient(lon)
        dy = np.gradient(lat)
        heading_tmp = np.unwrap(np.arctan2(dy, dx))
        steer_signed = np.gradient(heading_tmp) * 500.0
        steer_b = steer_signed

    if steer_col_a:
        steer_a = np.interp(grid_norm, lap_a["norm_d"], lap_a[steer_col_a].fillna(0))
    else:
        steer_a = steer_b

    heading_a = _compute_heading_deg(lat_a, lon_a)
    heading_b = _compute_heading_deg(lat, lon)

    if custom_corners is not None and len(custom_corners) > 0:
        sorted_custom = sorted(custom_corners, key=lambda x: float(x["ApexDist_m"]))
        turn_names = [str(c["Turn"]).strip() for c in sorted_custom]
        apex_indices = [
            int(np.clip(np.searchsorted(grid_d, float(c["ApexDist_m"])), 2, n_pts - 3)) for c in sorted_custom
        ]
        raw_segments = []
        for idx, pk in enumerate(apex_indices):
            left_b = 0 if idx == 0 else (apex_indices[idx - 1] + pk) // 2
            right_b = (n_pts - 1) if idx == len(apex_indices) - 1 else (pk + apex_indices[idx + 1]) // 2
            s = max(left_b, pk - 14)
            e = min(right_b, pk + 14)
            raw_segments.append((s, pk, e))
    else:
        raw_segments = _detect_signed_apexes(steer_signed, steer_thresh=steer_thresh, min_sep_pts=18)
        turn_names = [f"T{i + 1}" for i in range(len(raw_segments))]

    brake_points = []
    for idx, (s_turn, pk_turn, _) in enumerate(raw_segments):
        prev_pk = raw_segments[idx - 1][1] if idx > 0 else 0
        bp = _find_braking_or_lift_point(
            grid_d,
            throttle_b if throt_col_b else None,
            brake_b if brake_col_b else None,
            speed_b,
            s_turn,
            prev_apex_idx=prev_pk,
        )
        brake_points.append(bp)

    for i in range(1, len(brake_points)):
        if brake_points[i] <= brake_points[i - 1]:
            brake_points[i] = max(brake_points[i - 1] + 4, raw_segments[i][0])

    corners = []
    n_c = len(raw_segments)
    for idx in range(n_c):
        s_turn, pk_turn, e_turn = raw_segments[idx]
        sec_start = 0 if idx == 0 else brake_points[idx]
        sec_end = (n_pts - 1) if idx == n_c - 1 else brake_points[idx + 1]
        gap_val = float(delta_t[sec_end] - delta_t[sec_start])

        corners.append(
            {
                "Turn": turn_names[idx],
                "TurnNum": idx + 1,
                "ApexDist_m": round(float(grid_d[pk_turn]), 0),
                "s_idx": max(0, s_turn - 4),
                "e_idx": min(n_pts - 1, e_turn + 4),
                "sec_start": sec_start,
                "sec_end": sec_end,
                "mid_idx": pk_turn,
                "DistStart": round(float(grid_d[sec_start]), 0),
                "DistEnd": round(float(grid_d[sec_end]), 0),
                "GapSec": round(gap_val, 3),
                "MinSpeedA": round(float(np.min(speed_a[s_turn : e_turn + 1])), 1),
                "MinSpeedB": round(float(np.min(speed_b[s_turn : e_turn + 1])), 1),
            }
        )

    exact_total_gap = float(lap_row_a["LapTimeSec"] - lap_row_b["LapTimeSec"])

    return {
        "grid_d": grid_d,
        "t_a": t_a,
        "t_b": t_b,
        "delta_t": delta_t,
        "total_gap": exact_total_gap,
        "lat_a": lat_a,
        "lon_a": lon_a,
        "lat": lat,
        "lon": lon,
        "heading_a": heading_a,
        "heading_b": heading_b,
        "speed_a": speed_a,
        "speed_b": speed_b,
        "throttle_a": throttle_a,
        "throttle_b": throttle_b,
        "brake_a": brake_a,
        "brake_b": brake_b,
        "brake_unit": brake_unit,
        "gear_a": gear_a,
        "gear_b": gear_b,
        "rpm_a": rpm_a,
        "rpm_b": rpm_b,
        "steer_a": steer_a,
        "steer_b": steer_b,
        "corners": pd.DataFrame(corners),
    }


def compute_driver_personal_theoretical(df_motec_raw, df_motec_laps, driver_name, ref_corners_df):
    """
    Calcule le meilleur tour théorique idéal PERSONNEL du pilote sélectionné
    (somme de ses meilleurs temps sur chaque secteur T1..Tn parmi tous ses tours propres).
    """
    if df_motec_laps.empty or ref_corners_df is None or ref_corners_df.empty:
        return None

    drv_laps = df_motec_laps[df_motec_laps["Driver"] == driver_name].copy()
    if drv_laps.empty:
        return None

    best_real_sec = float(drv_laps["LapTimeSec"].min())
    best_real_row = drv_laps.sort_values("LapTimeSec").iloc[0]
    clean_laps = drv_laps[drv_laps["LapTimeSec"] <= best_real_sec * 1.07].copy()
    if clean_laps.empty:
        clean_laps = drv_laps.copy()

    n_pts = 1500
    grid_norm = np.linspace(0.0, 1.0, n_pts)
    sec_starts = ref_corners_df["sec_start"].astype(int).values
    sec_ends = ref_corners_df["sec_end"].astype(int).values

    all_sec_times = []
    valid_lap_nums = []
    for _, lrow in clean_laps.iterrows():
        b_idx = int(lrow["BeaconLapIdx"])
        lp = df_motec_raw[df_motec_raw["BeaconLapIdx"] == b_idx].copy().reset_index(drop=True)
        if len(lp) < 20:
            continue

        t_s = float(lrow["BeaconStart"])
        t_e = float(lrow["BeaconEnd"])
        dur = t_e - t_s

        dt = np.gradient(lp["Time"].values)
        ds = (lp["Corr Speed"].fillna(0).values / 3.6) * dt
        raw_d = np.cumsum(ds) - ds[0]

        v_first = float(lp["Corr Speed"].iloc[0]) / 3.6
        v_last = float(lp["Corr Speed"].iloc[-1]) / 3.6
        d_pre = max(0.0, (float(lp["Time"].iloc[0]) - t_s) * v_first)
        d_post = max(0.0, (t_e - float(lp["Time"].iloc[-1])) * v_last)

        full_d = raw_d + d_pre
        tot_d = float(full_d[-1] + d_post) if (full_d[-1] + d_post) > 0 else 1.0

        lap_t = lp["Time"].values - t_s
        norm_d = full_d / tot_d

        t_interp = np.interp(
            grid_norm,
            np.concatenate([[0.0], norm_d, [1.0]]),
            np.concatenate([[0.0], lap_t, [dur]]),
        )
        sec_t = t_interp[sec_ends] - t_interp[sec_starts]
        all_sec_times.append(sec_t)
        valid_lap_nums.append(int(lrow["OfficialLap"]))

    if not all_sec_times:
        return None

    mat = np.vstack(all_sec_times)
    best_per_sec = np.min(mat, axis=0)
    best_lap_idx_per_sec = np.argmin(mat, axis=0)
    theo_sec = float(np.sum(best_per_sec))
    gain_vs_best = theo_sec - best_real_sec

    best_real_bidx = int(best_real_row["BeaconLapIdx"])
    best_real_pos = None
    for i, (_, lrow) in enumerate(clean_laps.iterrows()):
        if int(lrow["BeaconLapIdx"]) == best_real_bidx and i < len(all_sec_times):
            best_real_pos = i
            break

    potential_gains = []
    if best_real_pos is not None:
        real_sec_vec = mat[best_real_pos]
        diffs_vec = real_sec_vec - best_per_sec
        order = np.argsort(-diffs_vec)
        for idx_s in order[:5]:
            if diffs_vec[idx_s] > 0.002:
                potential_gains.append(
                    {
                        "Turn": str(ref_corners_df["Turn"].iloc[idx_s]),
                        "GainSec": round(float(diffs_vec[idx_s]), 3),
                        "FromLap": valid_lap_nums[int(best_lap_idx_per_sec[idx_s])],
                    }
                )

    return {
        "driver": driver_name,
        "theo_sec": round(theo_sec, 3),
        "theo_str": format_laptime(theo_sec),
        "best_real_sec": round(best_real_sec, 3),
        "best_real_str": format_laptime(best_real_sec),
        "best_real_lap": int(best_real_row["OfficialLap"]),
        "gain_vs_best": round(gain_vs_best, 3),
        "n_clean_laps": len(all_sec_times),
        "potential_gains": potential_gains,
    }


def orient_track_main_straight_bottom(lat, lon):
    mean_lat_rad = np.deg2rad(np.mean(lat))
    x0 = (lon - np.mean(lon)) * np.cos(mean_lat_rad) * 111320.0
    y0 = (lat - np.mean(lat)) * 110540.0

    n_sf = max(15, len(x0) // 25)
    vx = float(x0[n_sf] - x0[-n_sf])
    vy = float(y0[n_sf] - y0[-n_sf])
    angle = math.atan2(vy, vx)

    cos_a = math.cos(-angle)
    sin_a = math.sin(-angle)
    xr = x0 * cos_a - y0 * sin_a
    yr = x0 * sin_a + y0 * cos_a

    if yr[0] > np.mean(yr):
        xr = -xr
        yr = -yr

    return xr, yr


def build_track_map_figure(analysis, circuit_name="Portimão"):
    x, y = orient_track_main_straight_bottom(analysis["lat"], analysis["lon"])
    df_c = analysis["corners"]

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=x, y=y, mode="lines", line=dict(color="#1e293b", width=14), hoverinfo="skip", showlegend=False
        )
    )
    fig.add_trace(
        go.Scatter(
            x=x, y=y, mode="lines", line=dict(color="#475569", width=8), hoverinfo="skip", showlegend=False
        )
    )
    fig.add_trace(
        go.Scatter(
            x=x, y=y, mode="lines", line=dict(color="#94a3b8", width=1.5), hoverinfo="skip", showlegend=False
        )
    )

    sf_arrow = "S/F ➔" if x[5] >= x[0] else "⬅ S/F"
    fig.add_trace(
        go.Scatter(
            x=[x[0]],
            y=[y[0]],
            mode="markers+text",
            marker=dict(size=12, color="#ffffff", symbol="line-ns", line=dict(color="#ffffff", width=3)),
            text=[sf_arrow],
            textposition="bottom center",
            textfont=dict(color="#e2e8f0", size=11, family="Consolas"),
            hoverinfo="skip",
            showlegend=False,
        )
    )

    cx_center = float(np.mean(x))
    cy_center = float(np.mean(y))
    span_x = float(np.max(x) - np.min(x)) if np.max(x) > np.min(x) else 100.0
    span_y = float(np.max(y) - np.min(y)) if np.max(y) > np.min(y) else 100.0
    offset_dist = max(span_x, span_y) * 0.095

    if not df_c.empty:
        for _, row in df_c.iterrows():
            s_i = int(row["s_idx"])
            e_i = int(row["e_idx"])
            m_i = int(row["mid_idx"])
            gap = float(row["GapSec"])
            is_loss = gap > 0
            seg_color = "#ef4444" if is_loss else "#22c55e"
            badge_bg = "#7f1d1d" if is_loss else "#14532d"

            fig.add_trace(
                go.Scatter(
                    x=x[s_i : e_i + 1],
                    y=y[s_i : e_i + 1],
                    mode="lines",
                    line=dict(color=seg_color, width=10),
                    hovertemplate=(
                        f"<b>{row['Turn']}</b> (Apex: {int(row['ApexDist_m'])}m | Secteur: {int(row['DistStart'])}m → {int(row['DistEnd'])}m)<br>"
                        f"Gap secteur (frein à frein): <b>{gap:+.3f} s</b><br>"
                        f"Vmin Pilote: {row['MinSpeedA']} km/h<br>"
                        f"Vmin Réf: {row['MinSpeedB']} km/h<extra></extra>"
                    ),
                    showlegend=False,
                )
            )

            vx = x[m_i] - cx_center
            vy = y[m_i] - cy_center
            norm = math.hypot(vx, vy) or 1.0
            lx = x[m_i] + (vx / norm) * offset_dist
            ly = y[m_i] + (vy / norm) * offset_dist

            fig.add_annotation(
                x=x[m_i],
                y=y[m_i],
                ax=lx,
                ay=ly,
                xref="x",
                yref="y",
                axref="x",
                ayref="y",
                text=f"<b>{row['Turn']}</b>  {gap:+.3f}s",
                showarrow=True,
                arrowhead=2,
                arrowsize=0.8,
                arrowwidth=1.5,
                arrowcolor=seg_color,
                font=dict(color="#ffffff", size=11, family="Consolas"),
                bgcolor=badge_bg,
                bordercolor=seg_color,
                borderwidth=1.5,
                borderpad=4,
            )

    pad_x = span_x * 0.18
    pad_y = span_y * 0.22
    fig.update_layout(
        paper_bgcolor="#111827",
        plot_bgcolor="#111827",
        margin=dict(l=15, r=15, t=20, b=20),
        height=520,
        xaxis=dict(
            visible=False,
            range=[float(np.min(x)) - pad_x, float(np.max(x)) + pad_x],
            scaleanchor="y",
            scaleratio=1,
        ),
        yaxis=dict(
            visible=False,
            range=[float(np.min(y)) - pad_y, float(np.max(y)) + pad_y],
        ),
    )
    return fig


def build_telemetry_traces_figure(analysis, drv_a_name, drv_b_name, x_range=None):
    """
    Construit la vue télémétrie multi-canaux (façon MoTeC i2) avec :
    1. Delta Temps (s)
    2. Rapport engagé (Gear)
    3. Vitesse (km/h)
    4. Accélérateur (%)
    5. Frein (bar / %)
    6. Angle volant (°)
    """
    d = analysis["grid_d"]
    color_a = "#ef4444"
    color_b = "#22d3ee"

    fig = make_subplots(
        rows=6,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.018,
        row_heights=[0.11, 0.11, 0.28, 0.17, 0.17, 0.16],
    )

    fig.add_hline(y=0, line=dict(color="#475569", width=1, dash="dot"), row=1, col=1)
    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["delta_t"],
            name="Delta (s)",
            line=dict(color="#facc15", width=1.8),
            hovertemplate="%{y:+.3f} s<extra>Delta</extra>",
            showlegend=False,
        ),
        row=1,
        col=1,
    )

    if analysis["gear_b"] is not None:
        fig.add_trace(
            go.Scatter(
                x=d,
                y=analysis["gear_b"],
                name=f"Gear {drv_b_name}",
                line=dict(color=color_b, width=1.6, shape="hv"),
                hovertemplate="R%{y:.0f}<extra>" + drv_b_name + "</extra>",
                showlegend=False,
            ),
            row=2,
            col=1,
        )
    if analysis["gear_a"] is not None:
        fig.add_trace(
            go.Scatter(
                x=d,
                y=analysis["gear_a"],
                name=f"Gear {drv_a_name}",
                line=dict(color=color_a, width=1.8, shape="hv"),
                hovertemplate="R%{y:.0f}<extra>" + drv_a_name + "</extra>",
                showlegend=False,
            ),
            row=2,
            col=1,
        )

    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["speed_b"],
            name=f"{drv_b_name} (Réf)",
            line=dict(color=color_b, width=1.7),
            hovertemplate="%{y:.1f} km/h<extra>" + drv_b_name + "</extra>",
        ),
        row=3,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["speed_a"],
            name=f"{drv_a_name} (Pilote)",
            line=dict(color=color_a, width=2.0),
            hovertemplate="%{y:.1f} km/h<extra>" + drv_a_name + "</extra>",
        ),
        row=3,
        col=1,
    )

    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["throttle_b"],
            name=f"Gaz {drv_b_name}",
            line=dict(color=color_b, width=1.6),
            hovertemplate="%{y:.1f} %<extra>" + drv_b_name + "</extra>",
            showlegend=False,
        ),
        row=4,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["throttle_a"],
            name=f"Gaz {drv_a_name}",
            line=dict(color="#22c55e", width=1.9),
            hovertemplate="%{y:.1f} %<extra>" + drv_a_name + "</extra>",
            showlegend=False,
        ),
        row=4,
        col=1,
    )

    b_unit = analysis.get("brake_unit", "bar")
    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["brake_b"],
            name=f"Frein {drv_b_name}",
            line=dict(color=color_b, width=1.6),
            hovertemplate=f"%{{y:.1f}} {b_unit}<extra>{drv_b_name}</extra>",
            showlegend=False,
        ),
        row=5,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["brake_a"],
            name=f"Frein {drv_a_name}",
            line=dict(color=color_a, width=1.9),
            hovertemplate=f"%{{y:.1f}} {b_unit}<extra>{drv_a_name}</extra>",
            showlegend=False,
        ),
        row=5,
        col=1,
    )

    fig.add_hline(y=0, line=dict(color="#334155", width=1, dash="dot"), row=6, col=1)
    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["steer_b"],
            name=f"Volant {drv_b_name}",
            line=dict(color=color_b, width=1.6),
            hovertemplate="%{y:.1f}°<extra>" + drv_b_name + "</extra>",
            showlegend=False,
        ),
        row=6,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=d,
            y=analysis["steer_a"],
            name=f"Volant {drv_a_name}",
            line=dict(color=color_a, width=1.8),
            hovertemplate="%{y:.1f}°<extra>" + drv_a_name + "</extra>",
            showlegend=False,
        ),
        row=6,
        col=1,
    )

    df_c = analysis["corners"]
    if not df_c.empty:
        for _, cr in df_c.iterrows():
            apex_d = float(cr["ApexDist_m"])
            fig.add_vline(
                x=apex_d,
                line=dict(color="rgba(148, 163, 184, 0.22)", width=1, dash="dash"),
            )
            fig.add_annotation(
                x=apex_d,
                y=1.0,
                xref="x",
                yref="paper",
                text=str(cr["Turn"]),
                showarrow=False,
                yshift=11,
                font=dict(color="#94a3b8", size=10, family="Consolas"),
            )

    y_titles = ["Delta (s)", "Gear", "Speed (km/h)", "Throttle (%)", f"Brake ({b_unit})", "Steer (°)"]
    for idx_r, y_title in enumerate(y_titles, start=1):
        fig.update_yaxes(
            title_text=y_title,
            title_font=dict(size=10, color="#94a3b8", family="Consolas"),
            tickfont=dict(size=9, color="#94a3b8", family="Consolas"),
            gridcolor="#1a2230",
            zerolinecolor="#273244",
            showline=True,
            linecolor="#2d3748",
            mirror=True,
            row=idx_r,
            col=1,
        )

    fig.update_yaxes(dtick=1, range=[0.5, 6.8], row=2, col=1)
    fig.update_yaxes(range=[-4, 106], row=4, col=1)

    x_min_val = float(d[0]) if x_range is None else float(x_range[0])
    x_max_val = float(d[-1]) if x_range is None else float(x_range[1])

    for idx_r in range(1, 7):
        fig.update_xaxes(
            range=[x_min_val, x_max_val],
            gridcolor="#1a2230",
            showline=True,
            linecolor="#2d3748",
            mirror=True,
            tickfont=dict(size=10, color="#cbd5e1", family="Consolas"),
            row=idx_r,
            col=1,
        )
    fig.update_xaxes(title_text="Distance (m)", title_font=dict(size=11, color="#cbd5e1", family="Consolas"), row=6, col=1)

    fig.update_layout(
        height=780,
        paper_bgcolor="#080b10",
        plot_bgcolor="#080b10",
        margin=dict(l=60, r=20, t=36, b=42),
        hovermode="x unified",
        dragmode="zoom",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.025,
            xanchor="right",
            x=1.0,
            font=dict(color="#f8fafc", size=11, family="Consolas"),
            bgcolor="rgba(15,23,42,0.7)",
        ),
    )
    return fig


# -----------------------------------------------------------------------------
# CLASSE PDF GRAFF
# -----------------------------------------------------------------------------
def _limit_top_n_keep_our_cars(df_tab, top_n=50, highlight_cars=None):
    if df_tab.empty or len(df_tab) <= top_n:
        return df_tab
    highlight_set = {c.strip() for c in (highlight_cars or []) if c.strip()}
    top_df = df_tab.iloc[:top_n].copy()
    if "N°" in df_tab.columns and highlight_set:
        extra_ours = df_tab.iloc[top_n:]
        extra_ours = extra_ours[extra_ours["N°"].astype(str).str.strip().isin(highlight_set)]
        if not extra_ours.empty:
            top_df = pd.concat([top_df, extra_ours], ignore_index=True)
    return top_df


class GraffPDF(FPDF):
    def __init__(self, doc_title, doc_subtitle, logo_path=None, orientation="P"):
        super().__init__(orientation=orientation, unit="mm", format="A4")
        self.doc_title = doc_title
        self.doc_subtitle = doc_subtitle
        self.logo_path = logo_path

    def header(self):
        page_w = 297 if self.cur_orientation == "L" else 210
        self.set_fill_color(15, 23, 42)
        self.rect(0, 0, page_w, 24, "F")

        if self.logo_path and os.path.exists(self.logo_path):
            logo_w = 46
            self.image(self.logo_path, x=page_w - logo_w - 10, y=6.5, w=logo_w)

        self.set_text_color(255, 255, 255)
        self.set_font("Helvetica", "B", 14)
        self.set_xy(10, 5.5)
        self.cell(
            page_w - 65,
            7,
            self.doc_title.encode("latin-1", "replace").decode("latin-1"),
            new_x="LMARGIN",
            new_y="NEXT",
        )
        self.set_font("Helvetica", "", 9)
        self.set_x(10)
        self.cell(
            page_w - 65,
            5.5,
            self.doc_subtitle.encode("latin-1", "replace").decode("latin-1"),
            new_x="LMARGIN",
            new_y="NEXT",
        )
        self.set_y(29)


def _pdf_add_section_table(
    pdf,
    title,
    df_tab,
    cols,
    widths,
    highlight_cars=None,
    left_cols=("Pilote", "Équipage", "Détail : moy. 50% par pilote (tours)"),
):
    if df_tab.empty:
        return
    highlight_set = {c.strip() for c in (highlight_cars or []) if c.strip()}

    pdf.set_text_color(15, 23, 42)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, title.encode("latin-1", "replace").decode("latin-1"), new_x="LMARGIN", new_y="NEXT")

    def draw_table_header():
        pdf.set_fill_color(24, 32, 46)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8)
        for col, w in zip(cols, widths):
            pdf.cell(w, 7, col.encode("latin-1", "replace").decode("latin-1"), border=1, align="C", fill=True)
        pdf.ln()

    draw_table_header()

    for idx, r in df_tab.iterrows():
        if pdf.get_y() + 7 > pdf.page_break_trigger:
            pdf.add_page()
            draw_table_header()

        car_val = str(r["N°"]).strip() if "N°" in df_tab.columns else ""
        is_our_car = car_val in highlight_set

        if is_our_car:
            pdf.set_fill_color(252, 207, 232)
            pdf.set_font("Helvetica", "B", 8)
        else:
            if idx % 2 == 1:
                pdf.set_fill_color(241, 245, 249)
            else:
                pdf.set_fill_color(255, 255, 255)
            pdf.set_font("Helvetica", "", 8)

        pdf.set_text_color(15, 23, 42)
        for col, w in zip(cols, widths):
            val = str(r[col]).encode("latin-1", "replace").decode("latin-1")
            align = "L" if col in left_cols else "C"
            pdf.cell(w, 6.2, val, border=1, align=align, fill=True)
        pdf.ln()
    pdf.ln(5)


def generate_drivers_pdf(
    session_name,
    circuit_name,
    df_pilotes_best,
    df_pilotes_avg,
    df_silver_best,
    df_silver_avg,
    df_bronze_best,
    df_bronze_avg,
    highlight_cars=None,
    logo_path=None,
):
    pdf = GraffPDF(
        doc_title=f"CLASSEMENTS PILOTES - {circuit_name.upper()} ({session_name})",
        doc_subtitle=f"Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')} | Top 50 pilotes (hors Pit/Out/FCY)",
        logo_path=logo_path,
        orientation="P",
    )
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    _pdf_add_section_table(
        pdf,
        "1. Classement Pilotes - Meilleur Tour (Top 50)",
        _limit_top_n_keep_our_cars(df_pilotes_best, 50, highlight_cars),
        ["Pos", "N°", "Pilote", "Cat.", "Tours Valides", "Meilleur Tour", "Écart"],
        [12, 14, 72, 14, 24, 28, 26],
        highlight_cars=highlight_cars,
    )
    _pdf_add_section_table(
        pdf,
        "2. Classement Pilotes - Moyenne des 50% Meilleurs Tours (Top 50)",
        _limit_top_n_keep_our_cars(df_pilotes_avg, 50, highlight_cars),
        ["Pos", "N°", "Pilote", "Cat.", "Tours Top 50%", "Moy. 50% Meilleurs", "Écart"],
        [12, 14, 70, 14, 24, 30, 26],
        highlight_cars=highlight_cars,
    )
    _pdf_add_section_table(
        pdf,
        "3. Classement SILVER (S) - Meilleur Tour",
        _limit_top_n_keep_our_cars(df_silver_best, 50, highlight_cars),
        ["Pos", "N°", "Pilote", "Tours Valides", "Meilleur Tour", "Écart"],
        [14, 16, 80, 26, 28, 26],
        highlight_cars=highlight_cars,
    )
    _pdf_add_section_table(
        pdf,
        "4. Classement SILVER (S) - Moyenne des 50% Meilleurs Tours",
        _limit_top_n_keep_our_cars(df_silver_avg, 50, highlight_cars),
        ["Pos", "N°", "Pilote", "Tours Top 50%", "Moy. 50% Meilleurs", "Écart"],
        [14, 16, 78, 26, 30, 26],
        highlight_cars=highlight_cars,
    )
    _pdf_add_section_table(
        pdf,
        "5. Classement BRONZE (B) - Meilleur Tour",
        _limit_top_n_keep_our_cars(df_bronze_best, 50, highlight_cars),
        ["Pos", "N°", "Pilote", "Tours Valides", "Meilleur Tour", "Écart"],
        [14, 16, 80, 26, 28, 26],
        highlight_cars=highlight_cars,
    )
    _pdf_add_section_table(
        pdf,
        "6. Classement BRONZE (B) - Moyenne des 50% Meilleurs Tours",
        _limit_top_n_keep_our_cars(df_bronze_avg, 50, highlight_cars),
        ["Pos", "N°", "Pilote", "Tours Top 50%", "Moy. 50% Meilleurs", "Écart"],
        [14, 16, 78, 26, 30, 26],
        highlight_cars=highlight_cars,
    )

    return bytes(pdf.output())


def generate_cars_pdf(session_name, circuit_name, df_cars_best, df_cars_avg, highlight_cars=None, logo_path=None):
    pdf = GraffPDF(
        doc_title=f"CLASSEMENTS PAR VOITURE - {circuit_name.upper()} ({session_name})",
        doc_subtitle=f"Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')} | Tours chronometres uniquement (hors in-laps et out-laps)",
        logo_path=logo_path,
        orientation="L",
    )
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    _pdf_add_section_table(
        pdf,
        "1. Classement par Voiture - Meilleur Tour",
        df_cars_best,
        ["Pos", "N°", "Équipage", "Cat.", "Meilleur tour", "Écart"],
        [14, 18, 155, 24, 34, 32],
        highlight_cars=highlight_cars,
    )
    _pdf_add_section_table(
        pdf,
        "2. Car ranking - average of drivers' best 50% laps",
        df_cars_avg,
        ["Pos", "N°", "Meilleur tour", "Moy. des moy. 50%", "Rang tour", "Détail : moy. 50% par pilote (tours)"],
        [14, 16, 30, 34, 22, 161],
        highlight_cars=highlight_cars,
    )

    return bytes(pdf.output())


# -----------------------------------------------------------------------------
# PACKING DONNÉES 2D + 3D NOVA NP02 & GÉNÉRATEURS HTML (STREAMLIT & NETLIFY/MOBILE)
# -----------------------------------------------------------------------------
def _pack_analysis_for_html(opt_id, label, drv_name, cand_row, ana):
    x_track, y_track = orient_track_main_straight_bottom(ana["lat"], ana["lon"])
    corners_list = []
    for _, cr in ana["corners"].iterrows():
        corners_list.append(
            {
                "Turn": cr["Turn"],
                "ApexDist": int(cr["ApexDist_m"]),
                "s_idx": int(cr["s_idx"]),
                "e_idx": int(cr["e_idx"]),
                "mid_idx": int(cr["mid_idx"]),
                "DistStart": int(cr["DistStart"]),
                "DistEnd": int(cr["DistEnd"]),
                "GapSec": float(cr["GapSec"]),
                "MinSpeedA": float(cr["MinSpeedA"]),
                "MinSpeedB": float(cr["MinSpeedB"]),
            }
        )

    step = 2
    step3d = 3  # 500 points par tour : ultra fluide en 3D et compact (< 35 Ko/tour)
    b_unit = ana.get("brake_unit", "bar")
    max_brk = max(65.0, float(np.nanmax(ana["brake_a"])), float(np.nanmax(ana["brake_b"])))

    def _norm_brk_3d(arr):
        if b_unit == "bar" or max_brk > 105.0:
            return [round(float(min(100.0, max(0.0, v / max_brk * 100.0))), 1) for v in arr]
        return [round(float(min(100.0, max(0.0, v))), 1) for v in arr]

    return {
        "id": opt_id,
        "label": label,
        "driver": drv_name,
        "lap": int(cand_row["OfficialLap"]),
        "lapTime": cand_row["LapTimeStr"],
        "lapTimeSec": float(cand_row["LapTimeSec"]),
        "fuel": float(cand_row["Fuel_L"]),
        "tyre": float(cand_row["Tyre_km"]),
        "totalGap": round(float(ana["total_gap"]), 3),
        "x": [round(float(v), 2) for v in x_track[::step]],
        "y": [round(float(v), 2) for v in y_track[::step]],
        "corners": [
            {
                **c,
                "s_idx": c["s_idx"] // step,
                "e_idx": c["e_idx"] // step,
                "mid_idx": c["mid_idx"] // step,
            }
            for c in corners_list
        ],
        "tel": {
            "d": [round(float(v), 1) for v in ana["grid_d"][::step]],
            "dt": [round(float(v), 3) for v in ana["delta_t"][::step]],
            "sa": [round(float(v), 1) for v in ana["speed_a"][::step]],
            "sb": [round(float(v), 1) for v in ana["speed_b"][::step]],
            "ta": [round(float(v), 1) for v in ana["throttle_a"][::step]],
            "tb": [round(float(v), 1) for v in ana["throttle_b"][::step]],
            "ba": [round(float(v), 1) for v in ana["brake_a"][::step]],
            "bb": [round(float(v), 1) for v in ana["brake_b"][::step]],
            "bUnit": b_unit,
            "ga": [int(v) for v in ana["gear_a"][::step]] if ana["gear_a"] is not None else [],
            "gb": [int(v) for v in ana["gear_b"][::step]] if ana["gear_b"] is not None else [],
            "sta": [round(float(v), 1) for v in ana["steer_a"][::step]],
            "stb": [round(float(v), 1) for v in ana["steer_b"][::step]],
        },
        "tel3d": {
            "bUnit": b_unit,
            "dist": [round(float(v), 1) for v in ana["grid_d"][::step3d]],
            "ta": [round(float(v), 3) for v in ana["t_a"][::step3d]],
            "lata": [round(float(v), 6) for v in ana["lat_a"][::step3d]],
            "lona": [round(float(v), 6) for v in ana["lon_a"][::step3d]],
            "spda": [round(float(v), 1) for v in ana["speed_a"][::step3d]],
            "hdga": [round(float(v), 1) for v in ana["heading_a"][::step3d]],
            "stra": [round(float(v), 1) for v in ana["steer_a"][::step3d]],
            "thra": [round(float(v), 1) for v in ana["throttle_a"][::step3d]],
            "brka": _norm_brk_3d(ana["brake_a"][::step3d]),
            "brkra": [round(float(v), 0) for v in ana["brake_a"][::step3d]],
            "ga": [int(v) for v in ana["gear_a"][::step3d]] if ana["gear_a"] is not None else [],
            "rpma": [int(v) for v in ana["rpm_a"][::step3d]],
            "tb": [round(float(v), 3) for v in ana["t_b"][::step3d]],
            "latb": [round(float(v), 6) for v in ana["lat"][::step3d]],
            "lonb": [round(float(v), 6) for v in ana["lon"][::step3d]],
            "spdb": [round(float(v), 1) for v in ana["speed_b"][::step3d]],
            "hdgb": [round(float(v), 1) for v in ana["heading_b"][::step3d]],
            "strb": [round(float(v), 1) for v in ana["steer_b"][::step3d]],
            "thrb": [round(float(v), 1) for v in ana["throttle_b"][::step3d]],
            "brkb": _norm_brk_3d(ana["brake_b"][::step3d]),
            "brkrb": [round(float(v), 0) for v in ana["brake_b"][::step3d]],
            "gb": [int(v) for v in ana["gear_b"][::step3d]] if ana["gear_b"] is not None else [],
            "rpmb": [int(v) for v in ana["rpm_b"][::step3d]],
        },
    }


def _get_3d_viewer_markup_and_script():
    """
    Retourne (css_3d, html_3d, js_3d) partagés entre Streamlit et la page HTML publiée.
    Intègre MapLibre GL JS + Relief 3D + Three.js + modèles Nova NP02 Rose & Bleue + import .glb/.gltf.
    """
    css_3d = r"""
  .viewer3d-wrap {
    position: relative;
    width: 100%;
    height: 620px;
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid #1e293b;
    background: #0b0f17;
  }
  @media (max-width: 768px) {
    .viewer3d-wrap { height: 490px; }
  }
  #map3d { width: 100%; height: 100%; }
  .maplibregl-ctrl-attrib { display: none !important; }

  .v3d-topbar {
    position: absolute; top: 10px; left: 10px; right: 10px;
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
    background: rgba(13, 17, 23, 0.90); border: 1px solid rgba(255, 255, 255, 0.13);
    border-radius: 9px; padding: 7px 11px; backdrop-filter: blur(8px);
    z-index: 20; flex-wrap: wrap;
  }
  .v3d-brand { display: flex; align-items: center; gap: 8px; font-weight: 700; font-size: 12px; color: #f1f5f9; }
  .v3d-badge {
    background: linear-gradient(135deg, #ff1e8a, #0095ff);
    padding: 2px 7px; border-radius: 4px; font-size: 10px; font-weight: 800; color: #fff;
  }
  .v3d-actions { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
  .v3d-btn {
    background: #232b3a; color: #f1f5f9; border: 1px solid rgba(255, 255, 255, 0.14);
    padding: 5px 8px; border-radius: 6px; font-size: 11px; font-weight: 600;
    cursor: pointer; transition: 0.15s;
  }
  .v3d-btn:hover { background: #313c52; }
  .v3d-btn.active { background: #2563eb; border-color: #60a5fa; color: #fff; }
  .v3d-btn.pink-btn { background: #ff1e8a; border-color: #ff1e8a; color: #fff; }

  .v3d-controls-panel {
    position: absolute; top: 54px; left: 10px; width: 255px;
    background: rgba(13, 17, 23, 0.90); border: 1px solid rgba(255, 255, 255, 0.13);
    border-radius: 9px; padding: 9px; backdrop-filter: blur(8px);
    display: flex; flex-direction: column; gap: 7px; z-index: 15;
  }
  .v3d-panel-title {
    font-size: 10px; font-weight: 700; text-transform: uppercase; color: #94a3b8;
    display: flex; justify-content: space-between; align-items: center;
  }
  .v3d-row-btns { display: flex; gap: 4px; flex-wrap: wrap; }
  .v3d-row-btns .v3d-btn { flex: 1; font-size: 10.5px; padding: 4px 5px; }
  .v3d-zoom-row {
    display: flex; align-items: center; gap: 6px; font-size: 10.5px; color: #94a3b8;
  }
  .v3d-zoom-row input[type="range"] { flex: 1; accent-color: #3b82f6; cursor: pointer; }

  .v3d-hud-right {
    position: absolute; top: 54px; right: 10px; width: 268px;
    background: rgba(13, 17, 23, 0.90); border: 1px solid rgba(255, 255, 255, 0.13);
    border-radius: 9px; padding: 9px; backdrop-filter: blur(8px);
    display: flex; flex-direction: column; gap: 7px; z-index: 15;
  }
  .v3d-delta-box {
    text-align: center; padding: 6px; border-radius: 7px;
    background: rgba(0, 0, 0, 0.42); border: 1px solid rgba(255, 255, 255, 0.12);
  }
  .v3d-delta-lbl { font-size: 9.5px; text-transform: uppercase; color: #94a3b8; }
  .v3d-delta-val { font-size: 18px; font-weight: 800; font-variant-numeric: tabular-nums; color: #f8fafc; }
  .v3d-telem-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; font-size: 11px; }
  .v3d-telem-col { background: rgba(24, 30, 41, 0.88); padding: 7px; border-radius: 6px; }
  .v3d-telem-col h4 { font-size: 10px; margin: 0 0 5px 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .v3d-stat { display: flex; justify-content: space-between; margin-bottom: 3px; font-variant-numeric: tabular-nums; font-size: 10.5px; }
  .v3d-bar-bg { width: 100%; height: 5px; background: #0b0e14; border-radius: 3px; overflow: hidden; margin-bottom: 4px; }
  .v3d-bar-fill { height: 100%; width: 0%; }

  .v3d-player {
    position: absolute; bottom: 10px; left: 10px; right: 10px;
    background: rgba(13, 17, 23, 0.92); border: 1px solid rgba(255, 255, 255, 0.13);
    border-radius: 9px; padding: 8px 12px; backdrop-filter: blur(8px);
    display: flex; align-items: center; gap: 10px; z-index: 20;
  }
  #v3dTimelineSlider { flex: 1; accent-color: #ff1e8a; cursor: pointer; }
  .v3d-time-display { font-variant-numeric: tabular-nums; font-size: 12px; font-weight: 600; min-width: 115px; text-align: right; color: #f1f5f9; }

  @media (max-width: 768px) {
    .v3d-controls-panel { width: calc(100% - 20px); top: auto; bottom: 58px; left: 10px; padding: 7px; }
    .v3d-hud-right { width: 225px; top: 52px; padding: 7px; }
    .v3d-hud-right.collapsed { display: none; }
    .v3d-controls-panel.collapsed { display: none; }
    .v3d-time-display { min-width: 92px; font-size: 11px; }
  }
"""

    html_3d = r"""
<div class="viewer3d-wrap" id="viewer3dContainer">
  <div id="map3d"></div>

  <div class="v3d-topbar">
    <div class="v3d-brand">
      <span class="v3d-badge">NOVA NP02 3D</span>
      <span id="v3dMatchLabel">Comparaison Tour</span>
    </div>
    <div class="v3d-actions">
      <button type="button" class="v3d-btn" id="btnToggleControls3d">Caméras</button>
      <button type="button" class="v3d-btn" id="btnToggleHud3d">HUD</button>
      <button type="button" class="v3d-btn" id="btnToggleRelief3d">Relief : ON</button>
      <button type="button" class="v3d-btn" id="btnModel3d1" title="Charger un modèle .glb/.gltf pour la Rose">.GLB Rose</button>
      <button type="button" class="v3d-btn" id="btnModel3d2" title="Charger un modèle .glb/.gltf pour la Bleue">.GLB Bleue</button>
      <input type="file" id="modelFile3d1" accept=".glb,.gltf" hidden />
      <input type="file" id="modelFile3d2" accept=".glb,.gltf" hidden />
    </div>
  </div>

  <div class="v3d-controls-panel" id="v3dControlsPanel">
    <div class="v3d-panel-title"><span>Caméra & Angle</span></div>
    <div class="v3d-row-btns">
      <button type="button" class="v3d-btn active" id="camChase1">Poursuite Rose</button>
      <button type="button" class="v3d-btn" id="camChase2">Poursuite Bleue</button>
    </div>
    <div class="v3d-row-btns">
      <button type="button" class="v3d-btn" id="camOverview">Vue circuit</button>
      <button type="button" class="v3d-btn" id="camFree">Libre</button>
    </div>
    <div class="v3d-zoom-row">
      <span>Zoom</span>
      <button type="button" class="v3d-btn" id="btnZoomOut3d" style="padding:2px 7px;">-</button>
      <input type="range" id="zoomSlider3d" min="16.0" max="22.5" step="0.1" value="19.5" />
      <button type="button" class="v3d-btn" id="btnZoomIn3d" style="padding:2px 7px;">+</button>
    </div>
  </div>

  <div class="v3d-hud-right" id="v3dHudRight">
    <div class="v3d-delta-box">
      <div class="v3d-delta-lbl">Écart instantané (Rose vs Bleue)</div>
      <div class="v3d-delta-val" id="v3dDeltaVal">+0.00 s</div>
    </div>
    <div class="v3d-telem-grid">
      <div class="v3d-telem-col" style="border-top: 2px solid #ff1e8a;">
        <h4 style="color: #ff1e8a;" id="v3dName1">ROSE</h4>
        <div class="v3d-stat"><span>Vitesse</span><b id="v1Spd3d">0 km/h</b></div>
        <div class="v3d-stat"><span>Rapport</span><b id="v1Gear3d">-</b></div>
        <div class="v3d-stat"><span>Régime</span><b id="v1Rpm3d">0</b></div>
        <div class="v3d-stat"><span>Volant</span><b id="v1Steer3d">0°</b></div>
        <div class="v3d-stat" style="font-size:9.5px; color:#94a3b8; margin-bottom:1px;"><span>Gaz</span><span id="v1ThrTxt3d">0%</span></div>
        <div class="v3d-bar-bg"><div class="v3d-bar-fill" id="v1Thr3d" style="background:#10b981;"></div></div>
        <div class="v3d-stat" style="font-size:9.5px; color:#94a3b8; margin-bottom:1px;"><span>Frein</span><span id="v1BrkTxt3d">0</span></div>
        <div class="v3d-bar-bg"><div class="v3d-bar-fill" id="v1Brk3d" style="background:#ef4444;"></div></div>
      </div>
      <div class="v3d-telem-col" style="border-top: 2px solid #0095ff;">
        <h4 style="color: #0095ff;" id="v3dName2">BLEUE</h4>
        <div class="v3d-stat"><span>Vitesse</span><b id="v2Spd3d">0 km/h</b></div>
        <div class="v3d-stat"><span>Rapport</span><b id="v2Gear3d">-</b></div>
        <div class="v3d-stat"><span>Régime</span><b id="v2Rpm3d">0</b></div>
        <div class="v3d-stat"><span>Volant</span><b id="v2Steer3d">0°</b></div>
        <div class="v3d-stat" style="font-size:9.5px; color:#94a3b8; margin-bottom:1px;"><span>Gaz</span><span id="v2ThrTxt3d">0%</span></div>
        <div class="v3d-bar-bg"><div class="v3d-bar-fill" id="v2Thr3d" style="background:#10b981;"></div></div>
        <div class="v3d-stat" style="font-size:9.5px; color:#94a3b8; margin-bottom:1px;"><span>Frein</span><span id="v2BrkTxt3d">0</span></div>
        <div class="v3d-bar-bg"><div class="v3d-bar-fill" id="v2Brk3d" style="background:#ef4444;"></div></div>
      </div>
    </div>
  </div>

  <div class="v3d-player">
    <button type="button" id="btnPlay3d" class="v3d-btn pink-btn" style="min-width: 68px;">Lecture</button>
    <select id="speedSelect3d" style="width: 66px; margin: 0; background:#10141c; color:#f1f5f9; border:1px solid rgba(255,255,255,0.15); border-radius:6px; padding:4px; font-size:11px;">
      <option value="0.25">0.25x</option>
      <option value="0.5">0.5x</option>
      <option value="1" selected>1.0x</option>
      <option value="2">2.0x</option>
    </select>
    <input type="range" id="v3dTimelineSlider" min="0" max="100" step="0.02" value="0" />
    <div class="v3d-time-display" id="v3dTimeDisplay">00:00.0 / 00:00.0</div>
  </div>
</div>
"""

    js_3d = r"""
/* ============================================================================
   MOTEUR 3D NOVA PROTO NP02 (MAPLIBRE GL JS + THREE.JS)
   ============================================================================ */
function buildNovaNP02Mesh(bodyHex) {
  const car = new THREE.Group();

  const bodyMat = new THREE.MeshStandardMaterial({ color: bodyHex, metalness: 0.4, roughness: 0.22 });
  const carbonMat = new THREE.MeshStandardMaterial({ color: 0x151820, metalness: 0.3, roughness: 0.5 });
  const glassMat = new THREE.MeshStandardMaterial({ color: 0x080c12, metalness: 0.92, roughness: 0.06 });
  const lightMat = new THREE.MeshStandardMaterial({ color: 0xe2e8f0, emissive: 0x64748b, roughness: 0.1 });
  const brakeMat = new THREE.MeshStandardMaterial({ color: 0x550000, emissive: 0x220000, roughness: 0.3 });

  function box(mat, x, y, z, w, l, h, rotX = 0) {
    const mesh = new THREE.Mesh(new THREE.BoxGeometry(w, l, h), mat);
    mesh.position.set(x, y, z);
    if (rotX) mesh.rotation.x = rotX;
    car.add(mesh);
    return mesh;
  }

  box(carbonMat, 0, 0.0, 0.08, 1.88, 4.45, 0.06);
  box(carbonMat, 0, 2.24, 0.10, 1.92, 0.36, 0.04);
  box(carbonMat, -0.95, 2.15, 0.22, 0.04, 0.38, 0.24);
  box(carbonMat,  0.95, 2.15, 0.22, 0.04, 0.38, 0.24);

  box(bodyMat, -0.68, 1.48, 0.34, 0.52, 1.48, 0.40, -0.08);
  box(bodyMat,  0.68, 1.48, 0.34, 0.52, 1.48, 0.40, -0.08);
  box(lightMat, -0.68, 2.18, 0.27, 0.24, 0.08, 0.16, -0.25);
  box(lightMat,  0.68, 2.18, 0.27, 0.24, 0.08, 0.16, -0.25);
  box(bodyMat, 0, 1.56, 0.24, 0.88, 1.38, 0.22, -0.07);
  box(carbonMat, 0, 1.48, 0.33, 0.38, 1.18, 0.10, -0.06);

  box(bodyMat, -0.73, -0.25, 0.34, 0.44, 2.15, 0.42);
  box(bodyMat,  0.73, -0.25, 0.34, 0.44, 2.15, 0.42);
  box(bodyMat, -0.68, -1.55, 0.42, 0.52, 1.15, 0.48);
  box(bodyMat,  0.68, -1.55, 0.42, 0.52, 1.15, 0.48);
  box(carbonMat, 0, -1.15, 0.46, 0.86, 1.85, 0.42);

  const canopy = new THREE.Mesh(new THREE.CylinderGeometry(0.42, 0.54, 1.45, 16), glassMat);
  canopy.position.set(0, 0.35, 0.66);
  canopy.scale.set(1.0, 1.0, 0.82);
  car.add(canopy);

  box(bodyMat,  0, 0.02, 0.92, 0.72, 0.85, 0.11);
  box(bodyMat,  0, -0.05, 1.02, 0.26, 0.56, 0.14);
  box(carbonMat, 0, 0.24, 1.02, 0.20, 0.05, 0.09);

  box(bodyMat, 0, -1.15, 0.78, 0.03, 1.55, 0.44);
  box(carbonMat, 0, -2.08, 0.95, 1.88, 0.34, 0.05);
  box(carbonMat, -0.45, -1.98, 0.72, 0.04, 0.22, 0.44);
  box(carbonMat,  0.45, -1.98, 0.72, 0.04, 0.22, 0.44);
  box(bodyMat, -0.94, -2.05, 0.86, 0.04, 0.48, 0.36);
  box(bodyMat,  0.94, -2.05, 0.86, 0.04, 0.48, 0.36);

  const brakeLight = box(brakeMat, 0, -2.23, 0.36, 0.24, 0.06, 0.11);

  function createWheel(x, y, z, radius, width) {
    const pivot = new THREE.Group();
    pivot.position.set(x, y, z);
    const geom = new THREE.CylinderGeometry(radius, radius, width, 24);
    geom.rotateZ(Math.PI / 2);
    const mesh = new THREE.Mesh(geom, carbonMat);
    pivot.add(mesh);
    car.add(pivot);
    return pivot;
  }

  const flWheel = createWheel(-0.78,  1.42, 0.31, 0.31, 0.32);
  const frWheel = createWheel( 0.78,  1.42, 0.31, 0.31, 0.32);
  createWheel(-0.78, -1.48, 0.33, 0.33, 0.36);
  createWheel( 0.78, -1.48, 0.33, 0.33, 0.36);

  return { group: car, flWheel, frWheel, brakeMat, brakeLight, isCustom: false };
}

const state3d = {
  car1: { activeData: [], model: null },
  car2: { activeData: [], model: null },
  originMercator: null,
  reliefOn: true,
  camMode: "chase1",
  chaseZoom: 19.5,
  playing: false,
  playbackSpeed: 1.0,
  currentTime: 0,
  maxDuration: 0,
  lastFrameMs: performance.now(),
  mapReady: false,
  pendingItem: null
};

const map3d = new maplibregl.Map({
  container: "map3d",
  center: [-8.6267, 37.2272], // Portimão par défaut
  zoom: 15.2,
  maxZoom: 23.5,
  pitch: 65,
  bearing: 0,
  antialias: true,
  style: {
    version: 8,
    sources: {
      satellite: {
        type: "raster",
        tiles: ["https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"],
        tileSize: 256,
        maxzoom: 19
      },
      terrainDem: {
        type: "raster-dem",
        tiles: ["https://s3.amazonaws.com/elevation-tiles-prod/terrarium/{z}/{x}/{y}.png"],
        encoding: "terrarium",
        tileSize: 256,
        maxzoom: 15
      }
    },
    layers: [
      { id: "satellite-layer", type: "raster", source: "satellite" }
    ],
    terrain: { source: "terrainDem", exaggeration: 1.25 }
  }
});

function loadCustomGlbModel3d(file, carKey) {
  if (!THREE.GLTFLoader) return;
  const url = URL.createObjectURL(file);
  const loader = new THREE.GLTFLoader();
  loader.load(url, (gltf) => {
    const root = gltf.scene;
    const wrapper = new THREE.Group();
    root.rotation.x = Math.PI / 2;
    wrapper.add(root);

    const box3 = new THREE.Box3().setFromObject(wrapper);
    const size = new THREE.Vector3();
    box3.getSize(size);
    const maxDim = Math.max(size.x, size.y, size.z) || 1;
    wrapper.scale.setScalar(4.60 / maxDim);

    const scaledBox = new THREE.Box3().setFromObject(wrapper);
    const center = new THREE.Vector3();
    scaledBox.getCenter(center);
    wrapper.position.set(-center.x, -center.y, -scaledBox.min.z);

    const container = new THREE.Group();
    container.add(wrapper);

    const brakeMat = new THREE.MeshStandardMaterial({ color: 0x550000, emissive: 0x220000 });
    const scene = threeLayer3d.scene;
    const oldGroup = state3d[carKey].model.group;
    const wasVisible = oldGroup.visible;
    scene.remove(oldGroup);

    container.visible = wasVisible;
    scene.add(container);

    state3d[carKey].model = { group: container, flWheel: null, frWheel: null, brakeMat, isCustom: true };
    map3d.triggerRepaint();
  });
}

function sample3dAtTime(data, t) {
  if (!data || data.length === 0) return null;
  if (t <= data[0].t) return data[0];
  if (t >= data[data.length - 1].t) return data[data.length - 1];

  let lo = 0, hi = data.length - 1;
  while (hi - lo > 1) {
    const mid = (lo + hi) >> 1;
    if (data[mid].t <= t) lo = mid; else hi = mid;
  }
  const a = data[lo], b = data[hi];
  const span = b.t - a.t;
  const f = span > 0 ? (t - a.t) / span : 0;

  let dHead = b.heading - a.heading;
  if (dHead > 180) dHead -= 360;
  if (dHead < -180) dHead += 360;

  return {
    lat: a.lat + (b.lat - a.lat) * f,
    lon: a.lon + (b.lon - a.lon) * f,
    speed: a.speed + (b.speed - a.speed) * f,
    heading: (a.heading + dHead * f + 360) % 360,
    steer: a.steer + (b.steer - a.steer) * f,
    throttle: a.throttle + (b.throttle - a.throttle) * f,
    brake: a.brake + (b.brake - a.brake) * f,
    brakeLabel: a.brakeLabel,
    gear: a.gear,
    rpm: Math.round(a.rpm + (b.rpm - a.rpm) * f),
    dist: a.dist + (b.dist - a.dist) * f
  };
}

const threeLayer3d = {
  id: "nova-np02-3d-layer",
  type: "custom",
  renderingMode: "3d",

  onAdd(mapInstance, gl) {
    this.camera = new THREE.Camera();
    this.scene = new THREE.Scene();

    const hemi = new THREE.HemisphereLight(0xffffff, 0x334155, 1.45);
    this.scene.add(hemi);
    const dir = new THREE.DirectionalLight(0xffffff, 1.45);
    dir.position.set(120, -150, 260);
    this.scene.add(dir);

    state3d.car1.model = buildNovaNP02Mesh(0xff1e8a);
    state3d.car2.model = buildNovaNP02Mesh(0x0095ff);
    state3d.car1.model.group.visible = false;
    state3d.car2.model.group.visible = false;

    this.scene.add(state3d.car1.model.group);
    this.scene.add(state3d.car2.model.group);

    this.renderer = new THREE.WebGLRenderer({
      canvas: mapInstance.getCanvas(),
      context: gl,
      antialias: true
    });
    this.renderer.autoClear = false;
  },

  render(gl, matrix) {
    if (!state3d.originMercator) return;
    const origin = state3d.originMercator;
    const scale = origin.meterInMercatorCoordinateUnits();

    const updateCar3D = (carObj, sample) => {
      if (!carObj.model) return;
      if (!sample) {
        carObj.model.group.visible = false;
        return;
      }
      carObj.model.group.visible = true;

      const merc = maplibregl.MercatorCoordinate.fromLngLat([sample.lon, sample.lat], 0);
      const eastMeters = (merc.x - origin.x) / scale;
      const northMeters = -(merc.y - origin.y) / scale;
      const groundZ = map3d.queryTerrainElevation([sample.lon, sample.lat]) || 0;

      carObj.model.group.position.set(eastMeters, northMeters, groundZ);
      carObj.model.group.rotation.z = -THREE.MathUtils.degToRad(sample.heading);

      if (carObj.model.flWheel && carObj.model.frWheel) {
        const wheelAngle = -THREE.MathUtils.degToRad(Math.max(-35, Math.min(35, sample.steer * 0.12)));
        carObj.model.flWheel.rotation.z = wheelAngle;
        carObj.model.frWheel.rotation.z = wheelAngle;
      }

      const braking = sample.brake > 8;
      carObj.model.brakeMat.color.setHex(braking ? 0xff1111 : 0x440000);
      carObj.model.brakeMat.emissive.setHex(braking ? 0xff0000 : 0x110000);
    };

    const s1 = sample3dAtTime(state3d.car1.activeData, state3d.currentTime);
    const s2 = sample3dAtTime(state3d.car2.activeData, state3d.currentTime);
    updateCar3D(state3d.car1, s1);
    updateCar3D(state3d.car2, s2);

    const worldToMercator = new THREE.Matrix4()
      .makeTranslation(origin.x, origin.y, 0)
      .scale(new THREE.Vector3(scale, -scale, scale));

    this.camera.projectionMatrix = new THREE.Matrix4().fromArray(matrix).multiply(worldToMercator);
    this.renderer.resetState();
    this.renderer.render(this.scene, this.camera);
  }
};

map3d.on("style.load", () => {
  map3d.addSource("traj3d1", { type: "geojson", data: { type: "FeatureCollection", features: [] } });
  map3d.addSource("traj3d2", { type: "geojson", data: { type: "FeatureCollection", features: [] } });

  map3d.addLayer({
    id: "traj3d2-line", type: "line", source: "traj3d2",
    paint: { "line-color": "#0095ff", "line-width": 4, "line-opacity": 0.85 }
  });
  map3d.addLayer({
    id: "traj3d1-line", type: "line", source: "traj3d1",
    paint: { "line-color": "#ff1e8a", "line-width": 4, "line-opacity": 0.9 }
  });
  map3d.addLayer(threeLayer3d);
  state3d.mapReady = true;
  if (state3d.pendingItem) {
    load3DComparison(state3d.pendingItem);
    state3d.pendingItem = null;
  }
});

function fmtTime3d(sec) {
  const m = Math.floor(sec / 60);
  const s = (sec % 60).toFixed(1).padStart(4, "0");
  return `${String(m).padStart(2, "0")}:${s}`;
}

function flyToOverview3d() {
  const pts = state3d.car1.activeData.length ? state3d.car1.activeData : state3d.car2.activeData;
  if (!pts.length) return;
  const bounds = new maplibregl.LngLatBounds();
  for (let i = 0; i < pts.length; i += 5) bounds.extend([pts[i].lon, pts[i].lat]);
  map3d.fitBounds(bounds, { padding: 80, pitch: 48, bearing: 0, duration: 700 });
}

function updateFrame3d(t) {
  const s1 = sample3dAtTime(state3d.car1.activeData, t);
  const s2 = sample3dAtTime(state3d.car2.activeData, t);

  if (s1) {
    document.getElementById("v1Spd3d").textContent = `${Math.round(s1.speed)} km/h`;
    document.getElementById("v1Gear3d").textContent = s1.gear || "-";
    document.getElementById("v1Rpm3d").textContent = s1.rpm ? `${s1.rpm}` : "-";
    document.getElementById("v1Steer3d").textContent = `${Math.round(s1.steer)}°`;
    document.getElementById("v1Thr3d").style.width = `${Math.min(100, Math.max(0, s1.throttle))}%`;
    document.getElementById("v1ThrTxt3d").textContent = `${Math.round(s1.throttle)}%`;
    document.getElementById("v1Brk3d").style.width = `${Math.min(100, Math.max(0, s1.brake))}%`;
    document.getElementById("v1BrkTxt3d").textContent = s1.brakeLabel;
  }
  if (s2) {
    document.getElementById("v2Spd3d").textContent = `${Math.round(s2.speed)} km/h`;
    document.getElementById("v2Gear3d").textContent = s2.gear || "-";
    document.getElementById("v2Rpm3d").textContent = s2.rpm ? `${s2.rpm}` : "-";
    document.getElementById("v2Steer3d").textContent = `${Math.round(s2.steer)}°`;
    document.getElementById("v2Thr3d").style.width = `${Math.min(100, Math.max(0, s2.throttle))}%`;
    document.getElementById("v2ThrTxt3d").textContent = `${Math.round(s2.throttle)}%`;
    document.getElementById("v2Brk3d").style.width = `${Math.min(100, Math.max(0, s2.brake))}%`;
    document.getElementById("v2BrkTxt3d").textContent = s2.brakeLabel;
  }

  const deltaEl = document.getElementById("v3dDeltaVal");
  if (s1 && s2) {
    const avgSpdMs = Math.max(10, ((s1.speed + s2.speed) / 2) / 3.6);
    const deltaSec = (s2.dist - s1.dist) / avgSpdMs;
    const sign = deltaSec >= 0 ? "+" : "";
    deltaEl.textContent = `${sign}${deltaSec.toFixed(2)} s`;
    deltaEl.style.color = deltaSec <= 0 ? "#10b981" : "#ef4444";
  }

  const targetSample = state3d.camMode === "chase1" ? s1 : (state3d.camMode === "chase2" ? s2 : null);
  if (targetSample && state3d.mapReady) {
    map3d.jumpTo({
      center: [targetSample.lon, targetSample.lat],
      zoom: state3d.chaseZoom,
      pitch: 67,
      bearing: targetSample.heading
    });
  }

  if (state3d.mapReady) map3d.triggerRepaint();
  document.getElementById("v3dTimeDisplay").textContent = `${fmtTime3d(t)} / ${fmtTime3d(state3d.maxDuration)}`;
}

function load3DComparison(d) {
  if (!d || !d.tel3d || !d.tel3d.lata || d.tel3d.lata.length < 5) return;
  if (!state3d.mapReady) {
    state3d.pendingItem = d;
    return;
  }

  const t3 = d.tel3d;
  const unitStr = t3.bUnit || "bar";
  const buildPts = (tArr, latArr, lonArr, spdArr, hdgArr, strArr, thrArr, brkArr, brkRawArr, gArr, rpmArr, distArr) => {
    const out = [];
    for (let i = 0; i < tArr.length; i++) {
      out.push({
        t: tArr[i],
        lat: latArr[i],
        lon: lonArr[i],
        speed: spdArr[i],
        heading: hdgArr[i],
        steer: strArr[i],
        throttle: thrArr[i],
        brake: brkArr[i],
        brakeLabel: unitStr === "bar" ? `${brkRawArr[i]} bar` : `${Math.round(brkArr[i])}%`,
        gear: gArr && gArr.length > i ? gArr[i] : 0,
        rpm: rpmArr && rpmArr.length > i ? rpmArr[i] : 0,
        dist: distArr[i]
      });
    }
    return out;
  };

  state3d.car1.activeData = buildPts(
    t3.ta, t3.lata, t3.lona, t3.spda, t3.hdga, t3.stra, t3.thra, t3.brka, t3.brkra, t3.ga, t3.rpma, t3.dist
  );
  state3d.car2.activeData = buildPts(
    t3.tb, t3.latb, t3.lonb, t3.spdb, t3.hdgb, t3.strb, t3.thrb, t3.brkb, t3.brkrb, t3.gb, t3.rpmb, t3.dist
  );

  if (Math.abs(t3.lata[0]) > 0.01) {
    state3d.originMercator = maplibregl.MercatorCoordinate.fromLngLat([t3.lona[0], t3.lata[0]], 0);
  }

  const shortA = DRIVER_A.split(" ").slice(-1)[0];
  const shortB = d.driver.split(" ").slice(-1)[0];
  document.getElementById("v3dName1").textContent = `${shortA} (ROSE)`;
  document.getElementById("v3dName2").textContent = `${shortB} (BLEUE)`;
  document.getElementById("v3dMatchLabel").textContent = `${shortA} vs ${shortB} (T${d.lap})`;

  const max1 = state3d.car1.activeData.length ? state3d.car1.activeData[state3d.car1.activeData.length - 1].t : 0;
  const max2 = state3d.car2.activeData.length ? state3d.car2.activeData[state3d.car2.activeData.length - 1].t : 0;
  state3d.maxDuration = Math.max(max1, max2, 0.1);
  state3d.currentTime = 0;

  const slider = document.getElementById("v3dTimelineSlider");
  slider.max = state3d.maxDuration;
  slider.value = 0;

  const s1 = map3d.getSource("traj3d1");
  const s2 = map3d.getSource("traj3d2");
  if (s1) {
    s1.setData({
      type: "FeatureCollection",
      features: [{ type: "Feature", geometry: { type: "LineString", coordinates: state3d.car1.activeData.map(p => [p.lon, p.lat]) } }]
    });
  }
  if (s2) {
    s2.setData({
      type: "FeatureCollection",
      features: [{ type: "Feature", geometry: { type: "LineString", coordinates: state3d.car2.activeData.map(p => [p.lon, p.lat]) } }]
    });
  }

  map3d.resize();
  updateFrame3d(0);
}

function animate3d(now) {
  const dt = (now - state3d.lastFrameMs) / 1000;
  state3d.lastFrameMs = now;
  if (state3d.playing && state3d.maxDuration > 0) {
    state3d.currentTime += dt * state3d.playbackSpeed;
    if (state3d.currentTime >= state3d.maxDuration) state3d.currentTime = 0;
    document.getElementById("v3dTimelineSlider").value = state3d.currentTime;
    updateFrame3d(state3d.currentTime);
  }
  requestAnimationFrame(animate3d);
}
requestAnimationFrame(animate3d);

// Contrôles UI 3D
if (window.innerWidth <= 768) {
  document.getElementById("v3dControlsPanel").classList.add("collapsed");
}
document.getElementById("btnToggleControls3d").onclick = () => {
  document.getElementById("v3dControlsPanel").classList.toggle("collapsed");
};
document.getElementById("btnToggleHud3d").onclick = () => {
  document.getElementById("v3dHudRight").classList.toggle("collapsed");
};
document.getElementById("btnToggleRelief3d").onclick = (e) => {
  state3d.reliefOn = !state3d.reliefOn;
  map3d.setTerrain(state3d.reliefOn ? { source: "terrainDem", exaggeration: 1.25 } : null);
  e.target.textContent = `Relief : ${state3d.reliefOn ? "ON" : "OFF"}`;
};
document.getElementById("btnModel3d1").onclick = () => document.getElementById("modelFile3d1").click();
document.getElementById("modelFile3d1").onchange = (e) => {
  if (e.target.files[0]) loadCustomGlbModel3d(e.target.files[0], "car1");
};
document.getElementById("btnModel3d2").onclick = () => document.getElementById("modelFile3d2").click();
document.getElementById("modelFile3d2").onchange = (e) => {
  if (e.target.files[0]) loadCustomGlbModel3d(e.target.files[0], "car2");
};

document.getElementById("btnPlay3d").onclick = (e) => {
  state3d.playing = !state3d.playing;
  e.target.textContent = state3d.playing ? "Pause" : "Lecture";
};
document.getElementById("speedSelect3d").onchange = (e) => {
  state3d.playbackSpeed = parseFloat(e.target.value);
};
document.getElementById("v3dTimelineSlider").oninput = (e) => {
  state3d.currentTime = parseFloat(e.target.value);
  updateFrame3d(state3d.currentTime);
};
document.getElementById("zoomSlider3d").oninput = (e) => {
  state3d.chaseZoom = parseFloat(e.target.value);
  if (state3d.camMode === "free") map3d.setZoom(state3d.chaseZoom);
  else updateFrame3d(state3d.currentTime);
};
document.getElementById("btnZoomIn3d").onclick = () => {
  state3d.chaseZoom = Math.min(22.5, state3d.chaseZoom + 0.4);
  document.getElementById("zoomSlider3d").value = state3d.chaseZoom.toFixed(1);
  if (state3d.camMode === "free") map3d.zoomIn();
  else updateFrame3d(state3d.currentTime);
};
document.getElementById("btnZoomOut3d").onclick = () => {
  state3d.chaseZoom = Math.max(16.0, state3d.chaseZoom - 0.4);
  document.getElementById("zoomSlider3d").value = state3d.chaseZoom.toFixed(1);
  if (state3d.camMode === "free") map3d.zoomOut();
  else updateFrame3d(state3d.currentTime);
};

function setActiveCamBtn3d(activeId) {
  ["camChase1", "camChase2", "camOverview", "camFree"].forEach(id => {
    document.getElementById(id).classList.toggle("active", id === activeId);
  });
}
document.getElementById("camChase1").onclick = () => { state3d.camMode = "chase1"; setActiveCamBtn3d("camChase1"); updateFrame3d(state3d.currentTime); };
document.getElementById("camChase2").onclick = () => { state3d.camMode = "chase2"; setActiveCamBtn3d("camChase2"); updateFrame3d(state3d.currentTime); };
document.getElementById("camOverview").onclick = () => { state3d.camMode = "free"; setActiveCamBtn3d("camOverview"); flyToOverview3d(); };
document.getElementById("camFree").onclick = () => { state3d.camMode = "free"; setActiveCamBtn3d("camFree"); };

map3d.getCanvas().addEventListener("wheel", (e) => {
  if (state3d.camMode === "chase1" || state3d.camMode === "chase2") {
    e.preventDefault();
    const delta = e.deltaY < 0 ? 0.35 : -0.35;
    state3d.chaseZoom = Math.min(22.5, Math.max(16.0, state3d.chaseZoom + delta));
    document.getElementById("zoomSlider3d").value = state3d.chaseZoom.toFixed(1);
    updateFrame3d(state3d.currentTime);
  }
}, { passive: false });
"""
    return css_3d, html_3d, js_3d


def build_streamlit_3d_html(drv_a, cand_row_b, ana):
    """Construit le composant 3D autonome pour affichage direct dans Streamlit."""
    packed = _pack_analysis_for_html("st_3d", f"Tour {cand_row_b['OfficialLap']}", cand_row_b["Driver"], cand_row_b, ana)
    css_3d, html_3d, js_3d = _get_3d_viewer_markup_and_script()

    page = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link href="https://unpkg.com/maplibre-gl@4.7.1/dist/maplibre-gl.css" rel="stylesheet" />
<script src="https://unpkg.com/maplibre-gl@4.7.1/dist/maplibre-gl.js"></script>
<script src="https://unpkg.com/three@0.146.0/build/three.min.js"></script>
<script src="https://unpkg.com/three@0.146.0/examples/js/loaders/GLTFLoader.js"></script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body, html { width: 100%; height: 100%; background: #0e1117; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; overflow: hidden; }
  __CSS_3D__
</style>
</head>
<body>
__HTML_3D__
<script>
const DRIVER_A = __DRIVER_A_JSON__;
const INITIAL_ITEM = __ITEM_JSON__;
__JS_3D__
load3DComparison(INITIAL_ITEM);
</script>
</body>
</html>"""

    return (
        page.replace("__CSS_3D__", css_3d)
        .replace("__HTML_3D__", html_3d)
        .replace("__JS_3D__", js_3d)
        .replace("__DRIVER_A_JSON__", json.dumps(drv_a, ensure_ascii=False))
        .replace("__ITEM_JSON__", json.dumps(packed, ensure_ascii=False))
    )


def generate_driver_html(
    circuit_name,
    session_name,
    drv_a,
    lap_a_info,
    df_m1_raw,
    all_sources,
    default_ref_key,
    personal_theo_info,
    steer_thresh=12.0,
    custom_corners=None,
    logo_path=None,
):
    ref_options = []
    seen_keys = set()

    for src_name, (df_raw_src, df_laps_src) in all_sources.items():
        if df_laps_src.empty:
            continue
        for drv_name, grp in df_laps_src.groupby("Driver"):
            sorted_grp = grp.sort_values("LapTimeSec")
            candidates = [sorted_grp.iloc[0]]
            if len(sorted_grp) > 1:
                candidates.append(sorted_grp.iloc[1])

            for cand_row in candidates:
                opt_id = f"{src_name}_{drv_name}_{int(cand_row['BeaconLapIdx'])}"
                if opt_id in seen_keys:
                    continue
                if (
                    src_name == "Voiture 1"
                    and drv_name == drv_a
                    and int(cand_row["BeaconLapIdx"]) == int(lap_a_info["BeaconLapIdx"])
                    and len(sorted_grp) > 1
                ):
                    continue

                ana = compute_corner_analysis(
                    df_m1_raw,
                    lap_a_info,
                    df_raw_src,
                    cand_row,
                    steer_thresh=steer_thresh,
                    custom_corners=custom_corners,
                )
                if ana is None:
                    continue

                seen_keys.add(opt_id)
                label = f"{drv_name} ({src_name} — Tour {int(cand_row['OfficialLap'])} : {cand_row['LapTimeStr']})"
                ref_options.append(_pack_analysis_for_html(opt_id, label, drv_name, cand_row, ana))
                break

    if default_ref_key not in seen_keys:
        parts = default_ref_key.split("_")
        src_name_def = parts[0]
        b_idx_def = int(parts[-1])
        if src_name_def in all_sources:
            df_raw_def, df_laps_def = all_sources[src_name_def]
            match_rows = df_laps_def[df_laps_def["BeaconLapIdx"] == b_idx_def]
            if not match_rows.empty:
                cand_row = match_rows.iloc[0]
                ana = compute_corner_analysis(
                    df_m1_raw,
                    lap_a_info,
                    df_raw_def,
                    cand_row,
                    steer_thresh=steer_thresh,
                    custom_corners=custom_corners,
                )
                if ana is not None:
                    drv_def = str(cand_row["Driver"])
                    label = f"{drv_def} ({src_name_def} — Tour {int(cand_row['OfficialLap'])} : {cand_row['LapTimeStr']})"
                    ref_options.insert(0, _pack_analysis_for_html(default_ref_key, label, drv_def, cand_row, ana))

    ref_options.sort(key=lambda item: 0 if item["id"] == default_ref_key else 1)
    payload_json = json.dumps(ref_options, ensure_ascii=False)

    options_html = "\n".join(
        f'<option value="{idx}">{opt["label"]}</option>' for idx, opt in enumerate(ref_options)
    )

    theo_str = personal_theo_info["theo_str"] if personal_theo_info else "-"
    gain_pot = personal_theo_info["gain_vs_best"] if personal_theo_info else 0.0
    n_clean = personal_theo_info["n_clean_laps"] if personal_theo_info else 0

    gains_list_html = ""
    if personal_theo_info and personal_theo_info.get("potential_gains"):
        items = "".join(
            f"<div style='display:flex; justify-content:space-between; padding:7px 10px; background:#0f172a; border-radius:6px; margin-bottom:6px; font-family:Consolas, monospace; font-size:0.86rem;'>"
            f"<span><strong>{g['Turn']}</strong> <span style='color:#94a3b8;'>(vu au T{g['FromLap']})</span></span>"
            f"<span style='color:#4ade80; font-weight:700;'>-{g['GainSec']:.3f} s</span></div>"
            for g in personal_theo_info["potential_gains"]
        )
        gains_list_html = (
            f"<div style='margin-top:10px; flex:1; display:flex; flex-direction:column; justify-content:space-between;'>"
            f"<div style='font-size:0.74rem; text-transform:uppercase; color:#94a3b8; margin-bottom:6px;'>Où gagner par rapport à ton meilleur tour :</div>"
            f"<div>{items}</div></div>"
        )

    logo_uri = get_graff_logo_base64(logo_path)
    if logo_uri:
        graff_header_right = f'<img src="{logo_uri}" alt="GRAFF" style="height:40px; width:auto; object-fit:contain;">'
    else:
        graff_header_right = '<div class="graff-logo-badge">GRAFF</div>'

    css_3d, html_3d, js_3d = _get_3d_viewer_markup_and_script()

    template = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0" />
<title>Driver Page — __DRV_A__ | __CIRCUIT_NAME__ (__SESSION_NAME__)</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<link href="https://unpkg.com/maplibre-gl@4.7.1/dist/maplibre-gl.css" rel="stylesheet" />
<script src="https://unpkg.com/maplibre-gl@4.7.1/dist/maplibre-gl.js"></script>
<script src="https://unpkg.com/three@0.146.0/build/three.min.js"></script>
<script src="https://unpkg.com/three@0.146.0/examples/js/loaders/GLTFLoader.js"></script>
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 24px;
    background-color: #0b0f19; color: #f1f5f9;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  }
  .container { max-width: 1180px; margin: 0 auto; }
  .topbar {
    display: flex; justify-content: space-between; align-items: center;
    background: #0f172a; border: 1px solid #1e293b; border-radius: 12px;
    padding: 14px 20px; margin-bottom: 18px; flex-wrap: wrap; gap: 10px;
  }
  .topbar-left { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
  .title { font-size: 1.45rem; font-weight: 800; letter-spacing: 0.05em; color: #f8fafc; }
  .badge {
    background: #1e293b; border: 1px solid #334155; padding: 5px 12px;
    border-radius: 20px; font-size: 0.85rem; color: #38bdf8; font-weight: 600;
  }
  .graff-logo-badge {
    font-size: 1.2rem; font-weight: 900; letter-spacing: 0.18em; color: #f8fafc;
    border: 1.5px solid #475569; padding: 5px 12px; border-radius: 6px;
  }
  .header-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 18px; }
  .card {
    background: #111827; border: 1px solid #1f2937; border-radius: 12px; padding: 16px;
  }
  .card-ref { border-color: #2563eb; }
  .card-header {
    font-size: 0.76rem; text-transform: uppercase; color: #94a3b8;
    margin-bottom: 6px; display: flex; justify-content: space-between; align-items: center;
  }
  .ref-select {
    width: 100%; background: #0f172a; color: #f8fafc; border: 1px solid #3b82f6;
    border-radius: 8px; padding: 8px 10px; font-size: 0.98rem; font-weight: 700;
    margin-bottom: 12px; cursor: pointer;
  }
  .driver-name { font-size: 1.3rem; font-weight: 700; margin-bottom: 12px; color: #ffffff; }
  .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }
  .stat-box { background: #0f172a; padding: 9px 6px; border-radius: 8px; text-align: center; }
  .stat-lbl { font-size: 0.7rem; color: #94a3b8; text-transform: uppercase; }
  .stat-val { font-size: 1.08rem; font-weight: 700; font-family: Consolas, monospace; margin-top: 3px; }

  .map-section {
    background: #111827; border: 1px solid #1e293b; border-radius: 14px;
    padding: 10px; margin-bottom: 16px;
  }
  #trackMap { width: 100%; height: 520px; }

  .total-gap-row {
    display: flex; align-items: center; gap: 16px; margin: 16px 0 22px 0; flex-wrap: wrap;
  }
  .total-gap-lbl { font-size: 1.2rem; font-weight: 700; color: #e2e8f0; }
  .total-gap-box {
    padding: 9px 30px; border-radius: 8px;
    font-size: 1.45rem; font-weight: 800; font-family: Consolas, monospace;
  }
  .gap-loss { background: rgba(239,68,68,0.18); border: 2px solid #ef4444; color: #fca5a5; }
  .gap-gain { background: rgba(34,197,94,0.18); border: 2px solid #22c55e; color: #86efac; }

  .bottom-grid { display: grid; grid-template-columns: 1.45fr 1fr; gap: 20px; align-items: stretch; margin-bottom: 24px; }
  .bottom-col { display: flex; flex-direction: column; }
  .section-title { font-size: 1.1rem; font-weight: 700; margin-bottom: 10px; color: #f8fafc; }
  .table-wrap {
    flex: 1; background: #111827; border: 1px solid #1f2937; border-radius: 12px;
    overflow-y: auto; max-height: 410px;
  }
  table { width: 100%; border-collapse: collapse; background: #111827; }
  th, td { padding: 9px 8px; text-align: center; border-bottom: 1px solid #1f2937; font-family: Consolas, monospace; font-size: 0.86rem; }
  th { background: #1e293b; color: #cbd5e1; font-size: 0.76rem; text-transform: uppercase; position: sticky; top: 0; }
  .theo-panel {
    flex: 1;
    background: linear-gradient(145deg, #111827 0%, #172033 100%);
    border: 1px solid #334155; border-radius: 12px; padding: 16px;
    display: flex; flex-direction: column; justify-content: space-between;
    max-height: 410px; box-sizing: border-box;
  }
  .theo-hero {
    background: rgba(56, 189, 248, 0.10); border: 1.5px solid #0284c7;
    border-radius: 10px; padding: 10px; text-align: center; margin: 8px 0;
  }
  .theo-time { font-size: 1.9rem; font-weight: 800; color: #38bdf8; font-family: Consolas, monospace; }

  .tel-section {
    background: #080b10; border: 1px solid #1e293b; border-radius: 14px; padding: 14px;
    margin-bottom: 24px;
  }
  .tel-toolbar {
    display: flex; justify-content: space-between; align-items: center;
    flex-wrap: wrap; gap: 10px; margin-bottom: 10px;
  }
  .tel-controls { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  .turn-select {
    background: #0f172a; color: #f8fafc; border: 1px solid #334155;
    border-radius: 6px; padding: 6px 10px; font-family: Consolas, monospace; font-size: 0.85rem;
    max-width: 100%;
  }
  .reset-btn {
    background: #1e293b; color: #38bdf8; border: 1px solid #334155;
    border-radius: 6px; padding: 6px 11px; font-size: 0.82rem; font-weight: 600; cursor: pointer;
  }
  .reset-btn:hover { background: #334155; }
  .reset-btn.active-touch { background: #2563eb; color: #fff; border-color: #60a5fa; }
  #telemetryChart { width: 100%; height: 780px; }

  .section-3d-box {
    background: #111827; border: 1px solid #1e293b; border-radius: 14px; padding: 14px;
  }

  @media (max-width: 820px) {
    body { padding: 12px; }
    .header-grid { grid-template-columns: 1fr; }
    .bottom-grid { grid-template-columns: 1fr; }
    .title { font-size: 1.2rem; }
    #trackMap { height: 370px; }
    #telemetryChart { height: 640px; }
    .theo-panel { max-height: none; }
  }
  __CSS_3D__
</style>
</head>
<body>
<div class="container">
  <div class="topbar">
    <div class="topbar-left">
      <div class="title">DRIVER PAGE</div>
      <div class="badge">__CIRCUIT_NAME__ — __SESSION_NAME__</div>
    </div>
    <div>__LOGO_HTML__</div>
  </div>

  <div class="header-grid">
    <div class="card">
      <div class="card-header">DRIVER — Tour __LAP_A_NUM__</div>
      <div class="driver-name">__DRV_A__</div>
      <div class="stats-grid">
        <div class="stat-box"><div class="stat-lbl">Lap Time</div><div class="stat-val">__LAP_A_TIME__</div></div>
        <div class="stat-box"><div class="stat-lbl">Fuel</div><div class="stat-val">__LAP_A_FUEL__ L</div></div>
        <div class="stat-box"><div class="stat-lbl">Tyres Age</div><div class="stat-val">__LAP_A_TYRE__ km</div></div>
      </div>
    </div>

    <div class="card card-ref">
      <div class="card-header">
        <span id="ref-lap-header">COMPARED TO</span>
        <span style="color:#60a5fa; font-size:0.72rem;">▼ Changer de référence</span>
      </div>
      <select id="refSelector" class="ref-select" onchange="updateComparison(this.value)">
        __OPTIONS_HTML__
      </select>
      <div class="stats-grid">
        <div class="stat-box"><div class="stat-lbl">Lap Time</div><div class="stat-val" id="ref-laptime">-</div></div>
        <div class="stat-box"><div class="stat-lbl">Fuel</div><div class="stat-val" id="ref-fuel">-</div></div>
        <div class="stat-box"><div class="stat-lbl">Tyres Age</div><div class="stat-val" id="ref-tyre">-</div></div>
      </div>
    </div>
  </div>

  <div class="map-section">
    <div id="trackMap"></div>
  </div>

  <div class="total-gap-row">
    <div class="total-gap-lbl">Total gap</div>
    <div id="totalGapBox" class="total-gap-box gap-loss">+0.000 s</div>
  </div>

  <div class="bottom-grid">
    <div class="bottom-col">
      <div class="section-title">Biggest gaps (de point de frein à point de frein suivant)</div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Virage</th>
              <th>Gap</th>
              <th>Vmin __DRV_A__</th>
              <th id="th-vmin-ref">Vmin Réf</th>
            </tr>
          </thead>
          <tbody id="gapsTableBody"></tbody>
        </table>
      </div>
    </div>

    <div class="bottom-col">
      <div class="section-title">Tour Théorique Idéal — __DRV_A__</div>
      <div class="theo-panel">
        <div style="font-size:0.76rem; color:#94a3b8; text-transform:uppercase; letter-spacing:0.06em;">
          Somme des meilleurs secteurs (__N_CLEAN__ tours valides)
        </div>
        <div class="theo-hero">
          <div style="font-size:0.72rem; color:#bae6fd; text-transform:uppercase;">Best Theoretical Lap</div>
          <div class="theo-time">__THEO_STR__</div>
          <div style="font-size:0.84rem; color:#4ade80; font-family:Consolas, monospace; margin-top:4px;">
            Gain potentiel vs Best réel : <strong>__GAIN_POT__ s</strong>
          </div>
        </div>
        __GAINS_LIST_HTML__
      </div>
    </div>
  </div>

  <div class="tel-section">
    <div class="tel-toolbar">
      <div>
        <div class="section-title" style="margin-bottom:2px;">DATA — Télémétrie Comparative</div>
        <div style="font-size:0.78rem; color:#94a3b8;">Utilise le menu par virage ou active le zoom tactile • Double-clic pour réinitialiser</div>
      </div>
      <div class="tel-controls">
        <select id="turnZoomSelect" class="turn-select" onchange="zoomToTurn(this.value)">
          <option value="all">Tour complet</option>
        </select>
        <button id="btnTouchZoom" class="reset-btn" onclick="toggleTouchZoom()">Zoom doigt : OFF</button>
        <button class="reset-btn" onclick="resetTelemetryZoom()">Reset Zoom</button>
      </div>
    </div>
    <div id="telemetryChart"></div>
  </div>

  <div class="section-3d-box">
    <div class="section-title" style="margin-bottom:4px;">3D — Comparaison Trajectoires & Relief (Nova NP02)</div>
    <div style="font-size:0.8rem; color:#94a3b8; margin-bottom:10px;">
      Vue synchronisée automatiquement sur le tour de référence sélectionné ci-dessus (Rose = __DRV_A__, Bleue = Référence)
    </div>
    __HTML_3D__
  </div>
</div>

<script>
const REF_DATA = __PAYLOAD_JSON__;
const DRIVER_A = __DRIVER_A_JSON__;
let currentRefObj = null;
const isMobileDevice = window.innerWidth <= 820 || ('ontouchstart' in window);
let touchZoomEnabled = !isMobileDevice;

if (touchZoomEnabled) {
  document.getElementById("btnTouchZoom").textContent = "Zoom doigt : ON";
  document.getElementById("btnTouchZoom").classList.add("active-touch");
}

function fmtSigned(val) {
  const sign = val >= 0 ? "+" : "";
  return sign + val.toFixed(3) + " s";
}

function toggleTouchZoom() {
  touchZoomEnabled = !touchZoomEnabled;
  const btn = document.getElementById("btnTouchZoom");
  btn.textContent = "Zoom doigt : " + (touchZoomEnabled ? "ON" : "OFF");
  btn.classList.toggle("active-touch", touchZoomEnabled);
  if (currentRefObj) {
    Plotly.relayout("telemetryChart", { dragmode: touchZoomEnabled ? "zoom" : false });
  }
}

function renderTelemetry(d, xRange = null) {
  const tel = d.tel;
  if (!tel || !tel.d || tel.d.length === 0) return;

  const colorA = "#ef4444";
  const colorB = "#22d3ee";
  const dist = tel.d;

  const traces = [
    { x: dist, y: tel.dt, xaxis: "x", yaxis: "y", name: "Delta (s)", line: { color: "#facc15", width: 1.8 }, hovertemplate: "%{y:+.3f} s<extra>Delta</extra>", showlegend: false },
    { x: dist, y: tel.gb, xaxis: "x2", yaxis: "y2", name: d.driver, line: { color: colorB, width: 1.6, shape: "hv" }, hovertemplate: "R%{y:.0f}<extra>" + d.driver + "</extra>", showlegend: false },
    { x: dist, y: tel.ga, xaxis: "x2", yaxis: "y2", name: DRIVER_A, line: { color: colorA, width: 1.8, shape: "hv" }, hovertemplate: "R%{y:.0f}<extra>" + DRIVER_A + "</extra>", showlegend: false },
    { x: dist, y: tel.sb, xaxis: "x3", yaxis: "y3", name: d.driver + " (Réf)", line: { color: colorB, width: 1.7 }, hovertemplate: "%{y:.1f} km/h<extra>" + d.driver + "</extra>" },
    { x: dist, y: tel.sa, xaxis: "x3", yaxis: "y3", name: DRIVER_A + " (Pilote)", line: { color: colorA, width: 2.0 }, hovertemplate: "%{y:.1f} km/h<extra>" + DRIVER_A + "</extra>" },
    { x: dist, y: tel.tb, xaxis: "x4", yaxis: "y4", name: d.driver, line: { color: colorB, width: 1.6 }, hovertemplate: "%{y:.1f} %<extra>" + d.driver + "</extra>", showlegend: false },
    { x: dist, y: tel.ta, xaxis: "x4", yaxis: "y4", name: DRIVER_A, line: { color: "#22c55e", width: 1.9 }, hovertemplate: "%{y:.1f} %<extra>" + DRIVER_A + "</extra>", showlegend: false },
    { x: dist, y: tel.bb, xaxis: "x5", yaxis: "y5", name: d.driver, line: { color: colorB, width: 1.6 }, hovertemplate: "%{y:.1f} " + tel.bUnit + "<extra>" + d.driver + "</extra>", showlegend: false },
    { x: dist, y: tel.ba, xaxis: "x5", yaxis: "y5", name: DRIVER_A, line: { color: colorA, width: 1.9 }, hovertemplate: "%{y:.1f} " + tel.bUnit + "<extra>" + DRIVER_A + "</extra>", showlegend: false },
    { x: dist, y: tel.stb, xaxis: "x6", yaxis: "y6", name: d.driver, line: { color: colorB, width: 1.6 }, hovertemplate: "%{y:.1f}°<extra>" + d.driver + "</extra>", showlegend: false },
    { x: dist, y: tel.sta, xaxis: "x6", yaxis: "y6", name: DRIVER_A, line: { color: colorA, width: 1.8 }, hovertemplate: "%{y:.1f}°<extra>" + DRIVER_A + "</extra>", showlegend: false }
  ];

  const shapes = [
    { type: "line", xref: "paper", yref: "y", x0: 0, x1: 1, y0: 0, y1: 0, line: { color: "#475569", width: 1, dash: "dot" } },
    { type: "line", xref: "paper", yref: "y6", x0: 0, x1: 1, y0: 0, y1: 0, line: { color: "#334155", width: 1, dash: "dot" } }
  ];
  const annotations = [];

  d.corners.forEach(c => {
    shapes.push({
      type: "line", xref: "x6", yref: "paper",
      x0: c.ApexDist, x1: c.ApexDist, y0: 0, y1: 1,
      line: { color: "rgba(148, 163, 184, 0.22)", width: 1, dash: "dash" }
    });
    annotations.push({
      x: c.ApexDist, y: 1.0, xref: "x6", yref: "paper",
      text: c.Turn, showarrow: false, yshift: 10,
      font: { color: "#94a3b8", size: 9, family: "Consolas" }
    });
  });

  const xr = xRange || [dist[0], dist[dist.length - 1]];
  const commonX = {
    range: xr, gridcolor: "#1a2230", showline: true, linecolor: "#2d3748", mirror: true,
    tickfont: { size: 9.5, color: "#cbd5e1", family: "Consolas" }
  };
  const commonY = (title, extra = {}) => Object.assign({
    title: { text: title, font: { size: 9.5, color: "#94a3b8", family: "Consolas" } },
    tickfont: { size: 8.5, color: "#94a3b8", family: "Consolas" },
    gridcolor: "#1a2230", zerolinecolor: "#273244", showline: true, linecolor: "#2d3748", mirror: true
  }, extra);

  const chartHeight = window.innerWidth <= 820 ? 640 : 780;
  const layout = {
    height: chartHeight,
    paper_bgcolor: "#080b10",
    plot_bgcolor: "#080b10",
    margin: { l: 52, r: 12, t: 34, b: 38 },
    hovermode: "x unified",
    dragmode: touchZoomEnabled ? "zoom" : false,
    shapes: shapes,
    annotations: annotations,
    legend: {
      orientation: "h", yanchor: "bottom", y: 1.025, xanchor: "right", x: 1.0,
      font: { color: "#f8fafc", size: 10.5, family: "Consolas" }, bgcolor: "rgba(15,23,42,0.7)"
    },
    xaxis: Object.assign({ matches: "x6", showticklabels: false }, commonX),
    xaxis2: Object.assign({ matches: "x6", showticklabels: false }, commonX),
    xaxis3: Object.assign({ matches: "x6", showticklabels: false }, commonX),
    xaxis4: Object.assign({ matches: "x6", showticklabels: false }, commonX),
    xaxis5: Object.assign({ matches: "x6", showticklabels: false }, commonX),
    xaxis6: Object.assign({ title: { text: "Distance (m)", font: { size: 10.5, color: "#cbd5e1", family: "Consolas" } } }, commonX),

    yaxis: commonY("Delta (s)", { domain: [0.89, 1.00] }),
    yaxis2: commonY("Gear", { domain: [0.77, 0.875], dtick: 1, range: [0.5, 6.8] }),
    yaxis3: commonY("Speed", { domain: [0.49, 0.755] }),
    yaxis4: commonY("Gaz (%)", { domain: [0.325, 0.475], range: [-4, 106] }),
    yaxis5: commonY("Frein (" + tel.bUnit + ")", { domain: [0.16, 0.31] }),
    yaxis6: commonY("Volant (°)", { domain: [0.00, 0.145] })
  };

  Plotly.react("telemetryChart", traces, layout, { responsive: true, scrollZoom: false, displayModeBar: !isMobileDevice });
}

function zoomToTurn(val) {
  if (!currentRefObj) return;
  if (val === "all") {
    resetTelemetryZoom();
    return;
  }
  const idx = parseInt(val, 10);
  const c = currentRefObj.corners[idx];
  if (!c) return;
  const maxD = currentRefObj.tel.d[currentRefObj.tel.d.length - 1];
  const d0 = Math.max(0, c.DistStart - 80);
  const d1 = Math.min(maxD, c.DistEnd + 80);
  Plotly.relayout("telemetryChart", { "xaxis6.range": [d0, d1] });
}

function resetTelemetryZoom() {
  if (!currentRefObj) return;
  document.getElementById("turnZoomSelect").value = "all";
  const maxD = currentRefObj.tel.d[currentRefObj.tel.d.length - 1];
  Plotly.relayout("telemetryChart", { "xaxis6.range": [0, maxD] });
}

__JS_3D__

function updateComparison(idx) {
  const d = REF_DATA[parseInt(idx, 10)];
  if (!d) return;
  currentRefObj = d;

  document.getElementById("ref-lap-header").textContent = "COMPARED TO — Tour " + d.lap;
  document.getElementById("ref-laptime").textContent = d.lapTime;
  document.getElementById("ref-fuel").textContent = d.fuel.toFixed(1) + " L";
  document.getElementById("ref-tyre").textContent = d.tyre.toFixed(1) + " km";
  document.getElementById("th-vmin-ref").textContent = "Vmin " + d.driver;

  const gapBox = document.getElementById("totalGapBox");
  gapBox.textContent = fmtSigned(d.totalGap);
  gapBox.className = "total-gap-box " + (d.totalGap > 0 ? "gap-loss" : "gap-gain");

  const sortedCorners = [...d.corners].sort((a, b) => b.GapSec - a.GapSec);
  const tbody = document.getElementById("gapsTableBody");
  tbody.innerHTML = sortedCorners.map((c, i) => {
    const color = c.GapSec > 0 ? "#f87171" : "#4ade80";
    return `<tr>
      <td>${i + 1}</td>
      <td><strong>${c.Turn}</strong></td>
      <td style="color:${color}; font-weight:700;">${fmtSigned(c.GapSec)}</td>
      <td>${c.MinSpeedA.toFixed(1)} km/h</td>
      <td>${c.MinSpeedB.toFixed(1)} km/h</td>
    </tr>`;
  }).join("");

  const turnSel = document.getElementById("turnZoomSelect");
  turnSel.innerHTML = '<option value="all">Tour complet</option>' + d.corners.map((c, i) =>
    `<option value="${i}">Zoom ${c.Turn} (${c.DistStart}m → ${c.DistEnd}m | ${fmtSigned(c.GapSec)})</option>`
  ).join("");

  const x = d.x;
  const y = d.y;
  const traces = [
    { x: x, y: y, mode: "lines", line: { color: "#1e293b", width: 14 }, hoverinfo: "skip", showlegend: false },
    { x: x, y: y, mode: "lines", line: { color: "#475569", width: 8 }, hoverinfo: "skip", showlegend: false },
    { x: x, y: y, mode: "lines", line: { color: "#94a3b8", width: 1.5 }, hoverinfo: "skip", showlegend: false },
    {
      x: [x[0]], y: [y[0]], mode: "markers+text",
      marker: { size: 12, color: "#ffffff", symbol: "line-ns", line: { color: "#ffffff", width: 3 } },
      text: [x[Math.min(5, x.length - 1)] >= x[0] ? "S/F ➔" : "⬅ S/F"], textposition: "bottom center",
      textfont: { color: "#e2e8f0", size: 11, family: "Consolas" },
      hoverinfo: "skip", showlegend: false
    }
  ];

  const cx = x.reduce((a, b) => a + b, 0) / x.length;
  const cy = y.reduce((a, b) => a + b, 0) / y.length;
  const minX = Math.min(...x), maxX = Math.max(...x);
  const minY = Math.min(...y), maxY = Math.max(...y);
  const spanX = maxX - minX || 100;
  const spanY = maxY - minY || 100;
  const offsetDist = Math.max(spanX, spanY) * 0.095;
  const annotations = [];

  d.corners.forEach(c => {
    const isLoss = c.GapSec > 0;
    const segColor = isLoss ? "#ef4444" : "#22c55e";
    const badgeBg = isLoss ? "#7f1d1d" : "#14532d";
    const mIdx = Math.min(c.mid_idx, x.length - 1);

    traces.push({
      x: x.slice(c.s_idx, c.e_idx + 1),
      y: y.slice(c.s_idx, c.e_idx + 1),
      mode: "lines",
      line: { color: segColor, width: 10 },
      hovertemplate: `<b>${c.Turn}</b> (${c.DistStart}m → ${c.DistEnd}m)<br>Gap: <b>${fmtSigned(c.GapSec)}</b><br>Vmin Pilote: ${c.MinSpeedA} km/h<br>Vmin Réf: ${c.MinSpeedB} km/h<extra></extra>`,
      showlegend: false
    });

    const vx = x[mIdx] - cx;
    const vy = y[mIdx] - cy;
    const norm = Math.hypot(vx, vy) || 1.0;
    annotations.push({
      x: x[mIdx], y: y[mIdx],
      ax: x[mIdx] + (vx / norm) * offsetDist,
      ay: y[mIdx] + (vy / norm) * offsetDist,
      xref: "x", yref: "y", axref: "x", ayref: "y",
      text: `<b>${c.Turn}</b>  ${c.GapSec >= 0 ? "+" : ""}${c.GapSec.toFixed(3)}s`,
      showarrow: true, arrowhead: 2, arrowsize: 0.8, arrowwidth: 1.5, arrowcolor: segColor,
      font: { color: "#ffffff", size: isMobileDevice ? 9.5 : 11, family: "Consolas" },
      bgcolor: badgeBg, bordercolor: segColor, borderwidth: 1.5, borderpad: 3
    });
  });

  const mapHeight = isMobileDevice ? 370 : 520;
  const layout = {
    paper_bgcolor: "#111827",
    plot_bgcolor: "#111827",
    margin: { l: 10, r: 10, t: 15, b: 15 },
    height: mapHeight,
    dragmode: false,
    annotations: annotations,
    xaxis: { visible: false, range: [minX - spanX * 0.20, maxX + spanX * 0.20], scaleanchor: "y", scaleratio: 1 },
    yaxis: { visible: false, range: [minY - spanY * 0.24, maxY + spanY * 0.24] }
  };

  Plotly.react("trackMap", traces, layout, { responsive: true, scrollZoom: false, displayModeBar: false });
  renderTelemetry(d);
  load3DComparison(d);
}

if (REF_DATA.length > 0) {
  updateComparison(0);
}
</script>
</body>
</html>"""

    html_content = (
        template.replace("__DRV_A__", drv_a)
        .replace("__DRIVER_A_JSON__", json.dumps(drv_a, ensure_ascii=False))
        .replace("__CIRCUIT_NAME__", circuit_name)
        .replace("__SESSION_NAME__", session_name)
        .replace("__LOGO_HTML__", graff_header_right)
        .replace("__LAP_A_NUM__", str(lap_a_info["OfficialLap"]))
        .replace("__LAP_A_TIME__", str(lap_a_info["LapTimeStr"]))
        .replace("__LAP_A_FUEL__", str(lap_a_info["Fuel_L"]))
        .replace("__LAP_A_TYRE__", str(lap_a_info["Tyre_km"]))
        .replace("__OPTIONS_HTML__", options_html)
        .replace("__N_CLEAN__", str(n_clean))
        .replace("__THEO_STR__", theo_str)
        .replace("__GAIN_POT__", f"{gain_pot:+.3f}")
        .replace("__GAINS_LIST_HTML__", gains_list_html)
        .replace("__PAYLOAD_JSON__", payload_json)
        .replace("__CSS_3D__", css_3d)
        .replace("__HTML_3D__", html_3d)
        .replace("__JS_3D__", js_3d)
    )
    return html_content


# -----------------------------------------------------------------------------
# INTERFACE PRINCIPALE (SIDEBAR + 2 ONGLETS)
# -----------------------------------------------------------------------------
graff_logo_path = find_graff_logo()
netlify_cfg = load_netlify_config()

with st.sidebar:
    if graff_logo_path:
        st.image(graff_logo_path, use_container_width=True)
    st.title("⚙️ Paramètres Session")
    circuit_name = st.text_input("Circuit", value="Portimão")
    session_name = st.selectbox("Session", ["PP1", "PP2", "FP1", "FP2", "Q", "R"], index=2)

    st.markdown("---")
    st.subheader("1. Fichier Chronométrage")
    lbl_file = st.file_uploader("Lap by Lap Organisateur (.xlsx / .csv)", type=["xlsx", "xls", "csv"])

    st.markdown("---")
    st.subheader("2. Télémétrie Voiture (MoTeC .csv ou VBOX .vbo)")
    motec_file_1 = st.file_uploader(
        "Data Voiture 1 — MoTeC (.csv) ou VBOX (.vbo)", type=["csv", "vbo"], key="m1"
    )
    car1_num = st.text_input("Numéro(s) Notre Voiture (surligné rose PDF, ex: 57 ou 57, 58)", value="57")

    motec_file_2 = st.file_uploader(
        "Data Voiture 2 — Référence optionnelle (.csv / .vbo)", type=["csv", "vbo"], key="m2"
    )
    car2_num = st.text_input("Numéro Voiture 2 (ex: 4)", value="")

    st.markdown("---")
    st.subheader("3. Fuel & Pneus — Voiture 1")
    col_f1, col_c1 = st.columns(2)
    fuel1_start = col_f1.number_input("Fuel départ (L)", min_value=0.0, max_value=150.0, value=65.0, step=1.0)
    fuel1_cons = col_c1.number_input("Conso / tour (L)", min_value=0.1, max_value=15.0, value=2.4, step=0.1)
    tyre1_start = st.number_input("Km pneus départ V1", min_value=0.0, max_value=2000.0, value=0.0, step=5.0)

    with st.expander("Ajouts Fuel / Changements Pneus V1"):
        st.caption("Ravitaillements V1")
        refuel1_df = st.data_editor(
            pd.DataFrame({"Tour": pd.Series(dtype="int"), "Litres ajoutés": pd.Series(dtype="float")}),
            num_rows="dynamic",
            key="ref1",
        )
        st.caption("Changements pneus V1")
        tyre1_df = st.data_editor(
            pd.DataFrame({"Tour": pd.Series(dtype="int"), "Nouveau km départ": pd.Series(dtype="float")}),
            num_rows="dynamic",
            key="tyr1",
        )

    if motec_file_2 is not None:
        st.markdown("---")
        st.subheader("4. Fuel & Pneus — Voiture 2")
        col_f2, col_c2 = st.columns(2)
        fuel2_start = col_f2.number_input("Fuel départ V2 (L)", min_value=0.0, max_value=150.0, value=65.0, step=1.0)
        fuel2_cons = col_c2.number_input("Conso / tour V2 (L)", min_value=0.1, max_value=15.0, value=2.4, step=0.1)
        tyre2_start = st.number_input("Km pneus départ V2", min_value=0.0, max_value=2000.0, value=0.0, step=5.0)
        with st.expander("Ajouts Fuel / Changements Pneus V2"):
            refuel2_df = st.data_editor(
                pd.DataFrame({"Tour": pd.Series(dtype="int"), "Litres ajoutés": pd.Series(dtype="float")}),
                num_rows="dynamic",
                key="ref2",
            )
            tyre2_df = st.data_editor(
                pd.DataFrame({"Tour": pd.Series(dtype="int"), "Nouveau km départ": pd.Series(dtype="float")}),
                num_rows="dynamic",
                key="tyr2",
            )
    else:
        fuel2_start, fuel2_cons, tyre2_start = 65.0, 2.4, 0.0
        refuel2_df = pd.DataFrame(columns=["Tour", "Litres ajoutés"])
        tyre2_df = pd.DataFrame(columns=["Tour", "Nouveau km départ"])

    st.markdown("---")
    with st.expander("🌐 Publication Web Pilotes (Netlify)"):
        netlify_site_id = st.text_input("Netlify Site ID", value=netlify_cfg["site_id"])
        netlify_token = st.text_input("Netlify Access Token", value=netlify_cfg["token"], type="password")
        if netlify_site_id != netlify_cfg["site_id"] or netlify_token != netlify_cfg["token"]:
            save_netlify_config(netlify_site_id, netlify_token)


our_cars_list = [x.strip() for x in car1_num.split(",") if x.strip()]
primary_car1 = our_cars_list[0] if our_cars_list else ""
df_lbl = parse_lap_by_lap(lbl_file) if lbl_file is not None else pd.DataFrame()

tab_stats, tab_data = st.tabs(["📊 Statistiques & Rythme Session", "🏎️ Data & Exploitation Pilote"])

# =============================================================================
# ONGLET 1 : STATISTIQUES DE SESSION
# =============================================================================
with tab_stats:
    render_graff_banner(
        title=f"STATISTIQUES & RYTHME SESSION — {circuit_name.upper()} ({session_name})",
        subtitle="Classements Pilotes, Voitures, Silver & Bronze (hors Pit/Out/FCY)",
        logo_path=graff_logo_path,
    )

    if df_lbl.empty:
        st.info("Importe le fichier **Lap by Lap de l'organisateur** dans la barre latérale pour afficher les classements de rythme.")
    else:
        (
            df_pilotes_best,
            df_pilotes_avg,
            df_cars_best,
            df_cars_avg,
            df_silver_best,
            df_silver_avg,
            df_bronze_best,
            df_bronze_avg,
        ) = compute_all_rankings(df_lbl)

        st.markdown(
            f"✅ **{len(df_lbl)} tours analysés** | **{len(df_pilotes_best)} pilotes classés** | "
            f"Voiture(s) surlignée(s) en rose : **{', '.join(our_cars_list) if our_cars_list else 'Aucune'}** (Top 50 sur PDF)"
        )

        col_pdf1, col_pdf2, _ = st.columns([1.3, 1.3, 2])
        with col_pdf1:
            pdf_drivers = generate_drivers_pdf(
                session_name,
                circuit_name,
                df_pilotes_best,
                df_pilotes_avg,
                df_silver_best,
                df_silver_avg,
                df_bronze_best,
                df_bronze_avg,
                highlight_cars=our_cars_list,
                logo_path=graff_logo_path,
            )
            st.download_button(
                label="📄 PDF Classements Pilotes (Top 50)",
                data=pdf_drivers,
                file_name=f"Classements_Pilotes_{circuit_name}_{session_name}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with col_pdf2:
            pdf_cars = generate_cars_pdf(
                session_name,
                circuit_name,
                df_cars_best,
                df_cars_avg,
                highlight_cars=our_cars_list,
                logo_path=graff_logo_path,
            )
            st.download_button(
                label="📄 PDF Classements Voitures",
                data=pdf_cars,
                file_name=f"Classements_Voitures_{circuit_name}_{session_name}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        st.markdown("### 1. Classement Pilotes — Meilleur Tour")
        render_clean_table(
            df_pilotes_best,
            ["Pos", "N°", "Pilote", "Cat.", "Tours Valides", "Meilleur Tour", "Écart"],
            highlight_cars=our_cars_list,
        )

        st.markdown("### 2. Classement Pilotes — Moyenne des 50% Meilleurs Tours")
        render_clean_table(
            df_pilotes_avg,
            ["Pos", "N°", "Pilote", "Cat.", "Tours Top 50%", "Moy. 50% Meilleurs", "Écart"],
            highlight_cars=our_cars_list,
        )

        st.markdown("### 3. Classement par Voiture — Meilleur Tour")
        render_clean_table(
            df_cars_best,
            ["Pos", "N°", "Équipage", "Cat.", "Meilleur tour", "Écart"],
            highlight_cars=our_cars_list,
        )

        st.markdown("### 4. Classement par Voiture — Moyenne des 50% Meilleurs Tours (Détail par pilote)")
        render_clean_table(
            df_cars_avg,
            ["Pos", "N°", "Meilleur tour", "Moy. des moy. 50%", "Rang tour", "Détail : moy. 50% par pilote (tours)"],
            highlight_cars=our_cars_list,
        )

        st.markdown("### 5. Classement SILVER (S) — Meilleur Tour")
        if df_silver_best.empty:
            st.write("Aucun pilote Silver détecté.")
        else:
            render_clean_table(
                df_silver_best,
                ["Pos", "N°", "Pilote", "Tours Valides", "Meilleur Tour", "Écart"],
                highlight_cars=our_cars_list,
                max_height=340,
            )

        st.markdown("### 6. Classement SILVER (S) — Moyenne des 50% Meilleurs Tours")
        if df_silver_avg.empty:
            st.write("Aucun pilote Silver détecté.")
        else:
            render_clean_table(
                df_silver_avg,
                ["Pos", "N°", "Pilote", "Tours Top 50%", "Moy. 50% Meilleurs", "Écart"],
                highlight_cars=our_cars_list,
                max_height=340,
            )

        st.markdown("### 7. Classement BRONZE (B) — Meilleur Tour")
        if df_bronze_best.empty:
            st.write("Aucun pilote Bronze détecté.")
        else:
            render_clean_table(
                df_bronze_best,
                ["Pos", "N°", "Pilote", "Tours Valides", "Meilleur Tour", "Écart"],
                highlight_cars=our_cars_list,
                max_height=340,
            )

        st.markdown("### 8. Classement BRONZE (B) — Moyenne des 50% Meilleurs Tours")
        if df_bronze_avg.empty:
            st.write("Aucun pilote Bronze détecté.")
        else:
            render_clean_table(
                df_bronze_avg,
                ["Pos", "N°", "Pilote", "Tours Top 50%", "Moy. 50% Meilleurs", "Écart"],
                highlight_cars=our_cars_list,
                max_height=340,
            )


# =============================================================================
# ONGLET 2 : DRIVER PAGE (2D + PUBLICATION NETLIFY + VUE 3D NOVA NP02)
# =============================================================================
with tab_data:
    presets_all = load_circuit_presets()
    circuit_key = circuit_name.strip().lower()
    saved_custom_corners = presets_all.get(circuit_key, None)
    preset_badge = (
        f" • Profil virages sauvegardé ({len(saved_custom_corners)} virages)" if saved_custom_corners else ""
    )

    render_graff_banner(
        title=f"DRIVER PAGE — {circuit_name.upper()} ({session_name})",
        subtitle=f"Exploitation télémétrie MoTeC (.csv) & VBOX (.vbo) • Lien mobile Netlify & Vue 3D Nova NP02{preset_badge}",
        logo_path=graff_logo_path,
    )

    if motec_file_1 is None:
        st.info(
            "Importe au moins le **fichier télémétrie de ta voiture (MoTeC .csv ou VBOX .vbo)** dans la barre latérale pour accéder à la Driver Page, à la vue 3D et générer le lien mobile."
        )
    else:
        meta1, df_m1_raw, df_m1_laps = parse_telemetry_file(motec_file_1)
        if df_m1_laps.empty:
            st.error("Aucun tour valide n'a pu être extrait du fichier Voiture 1. Vérifie les données GPS/vitesse du fichier.")
            st.stop()

        df_m1_laps = align_motec_with_lbl_and_stints(
            df_m1_laps,
            df_lbl,
            primary_car1,
            fuel1_start,
            fuel1_cons,
            refuel1_df.to_dict("records"),
            tyre1_start,
            tyre1_df.to_dict("records"),
        )
        all_sources = {"Voiture 1": (df_m1_raw, df_m1_laps)}

        if motec_file_2 is not None:
            meta2, df_m2_raw, df_m2_laps = parse_telemetry_file(motec_file_2)
            if not df_m2_laps.empty:
                df_m2_laps = align_motec_with_lbl_and_stints(
                    df_m2_laps,
                    df_lbl,
                    car2_num,
                    fuel2_start,
                    fuel2_cons,
                    refuel2_df.to_dict("records"),
                    tyre2_start,
                    tyre2_df.to_dict("records"),
                )
                all_sources["Voiture 2"] = (df_m2_raw, df_m2_laps)

        hdr_left, hdr_right = st.columns([3, 1])
        with hdr_left:
            src_fmt = meta1.get("Format", "Télémétrie")
            st.caption(f"Source Voiture 1 : **{src_fmt}** ({len(df_m1_laps)} tour(s) détecté(s))")
        with hdr_right:
            steer_threshold = st.slider(
                "Sensibilité auto virages (°)",
                min_value=4.0,
                max_value=25.0,
                value=10.0,
                step=1.0,
                disabled=(saved_custom_corners is not None),
            )

        col_drv, col_ref = st.columns(2)

        with col_drv:
            drivers_v1 = df_m1_laps["Driver"].unique().tolist()
            sel_driver_a = st.selectbox("DRIVER", drivers_v1, key="drv_a")
            laps_a_avail = df_m1_laps[df_m1_laps["Driver"] == sel_driver_a].sort_values("LapTimeSec")
            lap_a_labels = {
                row["BeaconLapIdx"]: f"Tour {row['OfficialLap']} — {row['LapTimeStr']}"
                for _, row in laps_a_avail.iterrows()
            }
            sel_lap_a_idx = st.selectbox(
                "Tour sélectionné (Pilote)", list(lap_a_labels.keys()), format_func=lambda k: lap_a_labels[k], key="lap_a"
            )
            row_a = df_m1_laps[df_m1_laps["BeaconLapIdx"] == sel_lap_a_idx].iloc[0]

            st.markdown(
                f"""
                <div class="driver-card">
                    <div style="display:flex; justify-content:space-between;">
                        <div><div class="metric-label">Lap time</div><div class="metric-value">{row_a['LapTimeStr']}</div></div>
                        <div><div class="metric-label">Fuel</div><div class="metric-value">{row_a['Fuel_L']} L</div></div>
                        <div><div class="metric-label">Tyres age</div><div class="metric-value">{row_a['Tyre_km']} km</div></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with col_ref:
            c_sub1, c_sub2 = st.columns([1, 2])
            with c_sub1:
                ref_car_choice = st.selectbox(
                    "Voiture Réf.", list(all_sources.keys()), index=len(all_sources) - 1, key="ref_car"
                )
            df_ref_raw, df_ref_laps = all_sources[ref_car_choice]
            with c_sub2:
                drivers_ref = df_ref_laps["Driver"].unique().tolist()
                sel_driver_b = st.selectbox("Compared to", drivers_ref, key="drv_b")

            laps_b_avail = df_ref_laps[df_ref_laps["Driver"] == sel_driver_b].sort_values("LapTimeSec")
            lap_b_labels = {
                row["BeaconLapIdx"]: f"Tour {row['OfficialLap']} — {row['LapTimeStr']}"
                for _, row in laps_b_avail.iterrows()
            }
            sel_lap_b_idx = st.selectbox(
                "Tour sélectionné (Référence)", list(lap_b_labels.keys()), format_func=lambda k: lap_b_labels[k], key="lap_b"
            )
            row_b = df_ref_laps[df_ref_laps["BeaconLapIdx"] == sel_lap_b_idx].iloc[0]

            st.markdown(
                f"""
                <div class="driver-card-ref">
                    <div style="display:flex; justify-content:space-between;">
                        <div><div class="metric-label">Lap time</div><div class="metric-value">{row_b['LapTimeStr']}</div></div>
                        <div><div class="metric-label">Fuel</div><div class="metric-value">{row_b['Fuel_L']} L</div></div>
                        <div><div class="metric-label">Tyres age</div><div class="metric-value">{row_b['Tyre_km']} km</div></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        analysis = compute_corner_analysis(
            df_m1_raw,
            row_a,
            df_ref_raw,
            row_b,
            steer_thresh=steer_threshold,
            custom_corners=saved_custom_corners,
        )

        if analysis is None:
            st.error("Impossible de calculer la comparaison sur ces deux tours (données GPS ou distance insuffisantes).")
        else:
            fig_map = build_track_map_figure(analysis, circuit_name=circuit_name)
            st.plotly_chart(fig_map, use_container_width=True)

            tot_gap = analysis["total_gap"]
            box_cls = "gap-box-pos" if tot_gap > 0 else "gap-box-neg"
            st.markdown(
                f"""
                <div class="total-gap-row">
                    <div class="total-gap-label">Total gap</div>
                    <div class="{box_cls}">{tot_gap:+.3f} s</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            personal_theo = compute_driver_personal_theoretical(
                df_m1_raw, df_m1_laps, sel_driver_a, analysis["corners"]
            )

            col_gaps, col_right_panel = st.columns([1.45, 1])

            with col_gaps:
                st.markdown("#### Biggest gaps (de point de frein à point de frein suivant)")
                df_worst = analysis["corners"].sort_values("GapSec", ascending=False).reset_index(drop=True)
                df_worst.insert(0, "#", df_worst.index + 1)
                df_display = df_worst[["#", "Turn", "GapSec", "MinSpeedA", "MinSpeedB"]].rename(
                    columns={
                        "Turn": "Virage",
                        "GapSec": "Gap (s)",
                        "MinSpeedA": f"Vmin {sel_driver_a} (km/h)",
                        "MinSpeedB": f"Vmin {sel_driver_b} (km/h)",
                    }
                )
                df_display["Gap (s)"] = df_display["Gap (s)"].apply(lambda x: f"{x:+.3f}")
                render_clean_table(df_display, list(df_display.columns), fixed_height=410)

            with col_right_panel:
                st.markdown(f"#### Tour Théorique Idéal — {sel_driver_a}")
                if personal_theo is not None:
                    gains_html = ""
                    if personal_theo["potential_gains"]:
                        rows_g = "".join(
                            f"<div style='display:flex; justify-content:space-between; padding:6px 10px; background:#0f172a; border-radius:6px; margin-bottom:5px; font-family:Consolas, monospace; font-size:0.86rem;'>"
                            f"<span><strong>{g['Turn']}</strong> <span style='color:#94a3b8;'>(vu au T{g['FromLap']})</span></span>"
                            f"<span style='color:#4ade80; font-weight:700;'>-{g['GainSec']:.3f} s</span></div>"
                            for g in personal_theo["potential_gains"]
                        )
                        gains_html = (
                            f"<div style='margin-top:8px; overflow-y:auto;'>"
                            f"<div style='font-size:0.74rem; text-transform:uppercase; color:#94a3b8; margin-bottom:6px;'>"
                            f"Top gains potentiels vs ton meilleur tour :</div>{rows_g}</div>"
                        )

                    st.markdown(
                        f"""
                        <div class="theo-hero-card">
                            <div>
                                <div style="font-size:0.76rem; color:#94a3b8; text-transform:uppercase; letter-spacing:0.06em;">
                                    Somme des meilleurs secteurs ({personal_theo['n_clean_laps']} tours valides)
                                </div>
                                <div class="theo-main-box">
                                    <div style="font-size:0.74rem; color:#bae6fd; text-transform:uppercase;">Best Theoretical Lap</div>
                                    <div class="theo-time-big">{personal_theo['theo_str']}</div>
                                    <div style="font-size:0.85rem; color:#4ade80; font-family:Consolas, monospace; margin-top:4px;">
                                        Gain potentiel vs Best réel : <strong>{personal_theo['gain_vs_best']:+.3f} s</strong>
                                    </div>
                                </div>
                            </div>
                            {gains_html}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            # -----------------------------------------------------------------
            # BOUTONS DE PUBLICATION NETLIFY (LIEN WHATSAPP) & EXPORT LOCAL
            # -----------------------------------------------------------------
            default_ref_id = f"{ref_car_choice}_{sel_driver_b}_{int(row_b['BeaconLapIdx'])}"
            html_str = generate_driver_html(
                circuit_name,
                session_name,
                sel_driver_a,
                row_a,
                df_m1_raw,
                all_sources,
                default_ref_key=default_ref_id,
                personal_theo_info=personal_theo,
                steer_thresh=steer_threshold,
                custom_corners=saved_custom_corners,
                logo_path=graff_logo_path,
            )
            safe_drv = re.sub(r"[^a-zA-Z0-9_-]", "_", sel_driver_a)
            safe_circ = re.sub(r"[^a-zA-Z0-9_-]", "_", circuit_name)
            export_filename = f"DriverPage_{safe_drv}_{safe_circ}_{session_name}.html"

            col_pub, col_dl = st.columns([1.6, 1])
            with col_pub:
                if st.button(
                    "🚀 Publier & Obtenir le lien Pilote (WhatsApp / Mobile)",
                    type="primary",
                    use_container_width=True,
                ):
                    with st.spinner("Mise en ligne de la page pilote (2D + Data + 3D) sur Netlify..."):
                        ok, res_msg = publish_to_netlify(
                            export_filename, html_str, netlify_site_id, netlify_token
                        )
                    if ok:
                        st.session_state["last_published_url"] = res_msg
                        st.session_state["last_published_driver"] = f"{sel_driver_a} ({circuit_name} {session_name})"
                    else:
                        st.error(res_msg)

            with col_dl:
                st.download_button(
                    label="💾 Télécharger .html (sauvegarde PC)",
                    data=html_str.encode("utf-8"),
                    file_name=export_filename,
                    mime="text/html",
                    use_container_width=True,
                )

            if st.session_state.get("last_published_url"):
                pub_url = st.session_state["last_published_url"]
                pub_lbl = st.session_state.get("last_published_driver", "")
                st.markdown(
                    f"""
                    <div class="netlify-link-box">
                        <div style="font-size:0.82rem; color:#6ee7b7; font-weight:700; text-transform:uppercase; margin-bottom:4px;">
                            ✅ Page en ligne prête pour mobile — {pub_lbl}
                        </div>
                        <div style="font-size:0.85rem; color:#cbd5e1;">
                            Copie le lien ci-dessous (icône à droite) et colle-le directement dans le groupe WhatsApp des pilotes :
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.code(pub_url, language=None)

            # -----------------------------------------------------------------
            # ÉDITEUR DE VIRAGES PAR CIRCUIT (SAUVEGARDABLE)
            # -----------------------------------------------------------------
            with st.expander("🛠️ Caler / Ajuster les numéros de virages pour ce circuit"):
                st.caption(
                    f"Ajuste les distances d'apex (en mètres depuis S/F, longueur totale du tour : {int(analysis['grid_d'][-1])} m) "
                    f"ou ajoute/supprime des lignes pour caler exactement les virages de **{circuit_name}** (ex: T1 à T15), puis clique sur Sauvegarder."
                )
                edit_df = analysis["corners"][["Turn", "ApexDist_m"]].copy()
                edited_corners = st.data_editor(
                    edit_df,
                    num_rows="dynamic",
                    use_container_width=True,
                    key=f"editor_{circuit_key}",
                )
                c_save, c_reset, _ = st.columns([1.2, 1.4, 2])
                with c_save:
                    if st.button("💾 Sauvegarder ce circuit", use_container_width=True):
                        valid_rows = edited_corners.dropna(subset=["Turn", "ApexDist_m"])
                        save_circuit_preset(circuit_name, valid_rows)
                        st.rerun()
                with c_reset:
                    if saved_custom_corners is not None:
                        if st.button("🔄 Revenir à la détection auto", use_container_width=True):
                            delete_circuit_preset(circuit_name)
                            st.rerun()

            st.markdown("---")

            # -----------------------------------------------------------------
            # SECTION DATA : TÉLÉMÉTRIE COMPARATIVE ZOOMABLE (FAÇON MOTEC)
            # -----------------------------------------------------------------
            tel_hdr_l, tel_hdr_r = st.columns([2.5, 1])
            with tel_hdr_l:
                st.markdown("### 📈 DATA — Télémétrie Comparative (Zoomable)")
                st.caption(
                    "Sélectionne une zone à la souris pour zoomer simultanément sur tous les canaux • Double-clique sur le graphique pour revenir au tour complet."
                )
            with tel_hdr_r:
                turn_zoom_options = ["Tour complet"] + [
                    f"{r['Turn']} ({int(r['DistStart'])}m → {int(r['DistEnd'])}m | {r['GapSec']:+.3f}s)"
                    for _, r in analysis["corners"].iterrows()
                ]
                selected_zoom = st.selectbox("Zoom rapide virage", turn_zoom_options, key="tel_zoom_turn")

            x_zoom_range = None
            if selected_zoom != "Tour complet":
                idx_t = turn_zoom_options.index(selected_zoom) - 1
                c_row = analysis["corners"].iloc[idx_t]
                max_dist = float(analysis["grid_d"][-1])
                x_zoom_range = [
                    max(0.0, float(c_row["DistStart"]) - 80.0),
                    min(max_dist, float(c_row["DistEnd"]) + 80.0),
                ]

            fig_tel = build_telemetry_traces_figure(
                analysis,
                drv_a_name=sel_driver_a,
                drv_b_name=sel_driver_b,
                x_range=x_zoom_range,
            )
            st.plotly_chart(fig_tel, use_container_width=True)

            st.markdown("---")

            # -----------------------------------------------------------------
            # SECTION 3D : COMPARAISON NOVA NP02 EN RELIEF (ROSE VS BLEUE)
            # -----------------------------------------------------------------
            st.markdown("### 🏎️ VUE 3D INTERACTIVE — Nova NP02 (Comparaison Tour & Relief)")
            st.caption(
                f"Synchronisé automatiquement sur **{sel_driver_a} (Tour {row_a['OfficialLap']})** [Rose] "
                f"vs **{sel_driver_b} (Tour {row_b['OfficialLap']})** [Bleue] • Utilise la molette ou le curseur pour zoomer, ou importe tes modèles .glb."
            )
            html_3d_widget = build_streamlit_3d_html(sel_driver_a, row_b, analysis)
            components.html(html_3d_widget, height=635, scrolling=False)