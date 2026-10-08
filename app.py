from __future__ import annotations

import importlib
from datetime import date, timedelta
from html import escape
from io import BytesIO

import streamlit as st

import nur_calculator as _nur_calculator

if not hasattr(_nur_calculator, "REGION_NUR_THRESHOLDS"):
    _nur_calculator = importlib.reload(_nur_calculator)

from nur_calculator import (
    REGION_NUR_THRESHOLDS,
    REGION_ORDER,
    CalculationResult,
    ImportResult,
    build_excel_report,
    calculate_nur,
    import_workbooks,
    top_worst_sites,
)
from ppt_report import build_powerpoint_report


st.set_page_config(page_title="Calcul NUR", page_icon="📡", layout="wide")

PERIODS = {
    "Jour (1 jour)": 1,
    "Semaine (7 jours)": 7,
    "Mois standard (30 jours)": 30,
}
CACHE_SCHEMA_VERSION = "site-names-v1"

COLUMN_NAMES = {
    "date": "Date",
    "technology": "Technologie",
    "region": "Région",
    "site_code": "Code site",
    "site_name": "Nom du site",
    "cell_name": "Cellule",
    "unavailability_seconds": "Indisponibilité (s)",
    "unavailability_minutes": "Indisponibilité (min)",
    "unavailability_hours": "Indisponibilité (h)",
    "source_row": "Ligne source",
    "cell_count": "Nombre de cellules",
    "technology_count": "Nombre de technologies",
    "site_count": "Nombre de sites",
    "period_days": "Nombre de jours",
    "nur": "NUR",
    "threshold": "Seuil NUR",
    "above_threshold": "Hors seuil",
    "threshold_gap": "Écart au seuil",
    "reason": "Motif",
    "category": "Catégorie",
}


@st.cache_data(show_spinner=False)
def load_workbooks_cached(
    sites_content: bytes, nur_content: bytes, schema_version: str
) -> ImportResult:
    del schema_version
    return import_workbooks(BytesIO(sites_content), BytesIO(nur_content))


@st.cache_data(show_spinner=False)
def calculate_cached(
    data: list[dict], start_date: date, period_days: int, schema_version: str
) -> CalculationResult:
    del schema_version
    return calculate_nur(data, start_date, period_days)


def format_value(key: str, value: object) -> str:
    if value is None:
        return ""
    if key == "nur":
        return f"{float(value):.2f}"
    if key == "threshold":
        return "—" if value is None else f"{float(value):.0f}"
    if key == "threshold_gap":
        return "—" if value is None else f"{float(value):+.2f}"
    if key == "above_threshold":
        return "Oui" if value else "Non"
    if key in {
        "unavailability_seconds",
        "unavailability_minutes",
        "unavailability_hours",
    }:
        return f"{float(value):,.2f}".replace(",", " ")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    return str(value)


def display_table(records: list[dict], limit: int = 500) -> None:
    if not records:
        st.info("Aucune donnée à afficher.")
        return
    columns = list(records[0])
    shown = records[:limit]
    header = "".join(
        f"<th>{escape(COLUMN_NAMES.get(column, column))}</th>" for column in columns
    )
    body_rows = []
    for record in shown:
        cells = "".join(
            f"<td>{escape(format_value(column, record.get(column)))}</td>"
            for column in columns
        )
        body_rows.append(f"<tr>{cells}</tr>")
    st.html(
        "<div class='nur-table-wrap'><table class='nur-table'>"
        f"<thead><tr>{header}</tr></thead><tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )
    if len(records) > limit:
        st.caption(
            f"{limit:,} lignes affichées sur {len(records):,}. "
            "Le rapport Excel contient toutes les lignes."
        )


def filtered_records(
    records: list[dict],
    selected_sites: list[str],
    selected_technologies: list[str],
    selected_regions: list[str] | None = None,
) -> list[dict]:
    return [
        row
        for row in records
        if (not selected_sites or row.get("site_code") in selected_sites)
        and (
            "technology" not in row
            or not selected_technologies
            or row.get("technology") in selected_technologies
        )
        and (not selected_regions or row.get("region") in selected_regions)
    ]


def display_top_sites(records: list[dict], limit: int = 15) -> None:
    top = sorted(records, key=lambda row: row["nur"], reverse=True)[:limit]
    if not top:
        st.info("Aucun site à classer.")
        return
    maximum = max(row["nur"] for row in top) or 1
    bars = []
    for row in top:
        width = row["nur"] / maximum * 100
        tone = "bad" if row.get("above_threshold") else "ok"
        region = escape(str(row.get("region") or ""))
        threshold = row.get("threshold")
        threshold_txt = f"  ·  seuil {threshold:.0f}" if threshold is not None else ""
        site_name = str(row.get("site_name") or "").strip()
        site_code = str(row.get("site_code") or "").strip()
        if site_name and site_name.casefold() != site_code.casefold():
            title = escape(site_name)
            details = f"{escape(site_code)}  ·  {region}{threshold_txt}"
        else:
            title = escape(site_code)
            details = f"{region}{threshold_txt}"
        bars.append(
            "<div class='bar-row'>"
            f"<span class='bar-label'><strong>{title}</strong>"
            f"<small>{details}</small></span>"
            f"<span class='bar-track'><span class='bar-fill {tone}' style='--bar-width:{width:.2f}%'></span></span>"
            f"<span class='bar-value {tone}'>{row['nur']:.1f}</span>"
            "</div>"
        )
    st.html("<div class='bar-chart'>" + "".join(bars) + "</div>")


def display_region_gauges(regions: list[dict]) -> None:
    if not regions:
        st.info("Aucune région à afficher.")
        return
    cards = []
    ordered = sorted(
        regions,
        key=lambda row: REGION_ORDER.index(row["region"])
        if row["region"] in REGION_ORDER
        else 99,
    )
    for row in ordered:
        threshold = row.get("threshold")
        above = bool(row.get("above_threshold"))
        tone = "bad" if above else "ok"
        ratio = 0.0
        if threshold:
            ratio = min(row["nur"] / max(threshold, 1), 1.35) / 1.35 * 100
        threshold_label = f"{threshold:.0f}" if threshold is not None else "—"
        cards.append(
            "<article class='region-card'>"
            f"<div class='region-head'><span>{escape(row['region'])}</span>"
            f"<strong class='{tone}'>{row['nur']:.1f}</strong></div>"
            f"<small>seuil {threshold_label}  ·  {row['site_count']} sites</small>"
            "<div class='bar-track'><span class='bar-fill "
            f"{tone}' style='--bar-width:{ratio:.2f}%'></span></div>"
            f"<em>{'Hors seuil' if above else 'Sous le seuil'}</em>"
            "</article>"
        )
    st.html("<section class='region-grid'>" + "".join(cards) + "</section>")


def display_period_tops(
    data: list[dict], start: date, schema_version: str, *, limit: int = 8
) -> None:
    columns = st.columns(3)
    labels = ((1, "Jour"), (7, "Semaine"), (30, "Mois"))
    for column, (days, title) in zip(columns, labels):
        result = calculate_cached(data, start, days, schema_version)
        top = top_worst_sites(result.sites, limit=limit)
        with column:
            st.markdown(f"**{title} ({days} j)**")
            st.caption(
                f"{result.start_date.strftime('%d/%m')} → {result.end_date.strftime('%d/%m')}"
                f"  ·  {len(result.sites_above_threshold)} hors seuil"
            )
            display_top_sites(top, limit=limit)


def format_duration(seconds: float) -> str:
    if seconds >= 3_600:
        return f"{seconds / 3_600:,.1f} h".replace(",", " ")
    return f"{seconds / 60:,.1f} min".replace(",", " ")


def display_duration_bars(
    records: list[dict],
    category_key: str,
    *,
    limit: int = 15,
    sort_by_value: bool = True,
) -> None:
    selected = (
        sorted(records, key=lambda row: row["unavailability_hours"], reverse=True)
        if sort_by_value
        else records
    )[:limit]
    if not selected:
        st.info("Aucune donnée à afficher.")
        return
    maximum = max(row["unavailability_hours"] for row in selected) or 1
    bars = []
    for row in selected:
        width = row["unavailability_hours"] / maximum * 100
        label = format_value(category_key, row[category_key])
        duration = format_duration(row["unavailability_seconds"])
        bars.append(
            "<div class='bar-row'>"
            f"<span class='bar-label'>{escape(label)}</span>"
            f"<span class='bar-track'><span class='bar-fill' style='--bar-width:{width:.2f}%'></span></span>"
            f"<span class='bar-value'>{escape(duration)}</span>"
            "</div>"
        )
    st.html("<div class='bar-chart'>" + "".join(bars) + "</div>")


def display_kpi_dashboard(calculation: CalculationResult) -> None:
    network = calculation.network[0]
    cards = [
        (
            "NUR TOTAL · TOUS LES SITES",
            f"{network['nur']:.2f}",
            "Indice réseau normalisé sur 100 000",
            "primary",
        ),
        (
            "SITES HORS SEUIL",
            str(len(calculation.sites_above_threshold)),
            f"sur {network['site_count']} sites analysés",
            "orange",
        ),
        (
            "COUVERTURE ANALYSÉE",
            f"{network['site_count']:,}".replace(",", " "),
            f"{network['cell_count']:,} cellules actives".replace(",", " "),
            "cyan",
        ),
        (
            "VOLUME DE DONNÉES",
            f"{calculation.selected_rows:,}".replace(",", " "),
            f"{calculation.period_days} jour(s) de mesure",
            "violet",
        ),
    ]
    html = []
    for label, value, note, tone in cards:
        html.append(
            f"<article class='kpi-card {tone}'>"
            "<span class='kpi-glow'></span>"
            f"<div class='kpi-label'>{escape(label)}</div>"
            f"<div class='kpi-value'>{escape(value)}</div>"
            f"<div class='kpi-note'>{escape(note)}</div>"
            "</article>"
        )
    st.html("<section class='kpi-grid'>" + "".join(html) + "</section>")


def display_technology_nur(technologies: list[dict]) -> None:
    cards = []
    colors = {"2G": "#7c817e", "3G": "#47665a", "4G": "#f36b16"}
    maximum = max((row["nur"] for row in technologies), default=1) or 1
    for row in technologies:
        width = row["nur"] / maximum * 100
        color = colors.get(row["technology"], "#a78bfa")
        cell_label = f"{row['cell_count']:,} cellules".replace(",", " ")
        cards.append(
            "<article class='tech-card'>"
            f"<div class='tech-icon' style='--tech-color:{color}'>{escape(row['technology'])}</div>"
            "<div class='tech-data'>"
            "<span>NUR TOTAL</span>"
            f"<strong>{row['nur']:.2f}</strong>"
            f"<small>{format_duration(row['unavailability_seconds'])} · {cell_label}</small>"
            f"<div class='tech-meter'><i style='--tech-color:{color};--meter-width:{width:.2f}%'></i></div>"
            "</div></article>"
        )
    st.html("<section class='tech-grid'>" + "".join(cards) + "</section>")


def display_technology_donut(technologies: list[dict]) -> None:
    total = sum(row["unavailability_hours"] for row in technologies) or 1
    colors = {"2G": "#7c817e", "3G": "#47665a", "4G": "#f36b16"}
    offset = 25.0
    circles = []
    legend = []
    for index, row in enumerate(technologies):
        share = row["unavailability_hours"] / total * 100
        color = colors.get(row["technology"], "#a78bfa")
        circles.append(
            f"<circle class='donut-segment segment-{index}' cx='90' cy='90' r='66' "
            f"stroke='{color}' stroke-dasharray='{share:.3f} {100-share:.3f}' "
            f"stroke-dashoffset='-{offset:.3f}' pathLength='100'>"
            f"<title>{escape(row['technology'])}: {share:.1f}%</title></circle>"
        )
        legend.append(
            "<div class='donut-legend-row'>"
            f"<i style='background:{color}'></i><span>{escape(row['technology'])}</span>"
            f"<strong>{share:.1f}%</strong></div>"
        )
        offset += share
    st.html(
        "<div class='viz-card donut-layout'>"
        "<div class='viz-heading'><span>RÉPARTITION TECHNOLOGIQUE</span>"
        "<strong>Part des heures-cellules</strong></div>"
        "<div class='donut-wrap'><svg viewBox='0 0 180 180'>"
        "<circle class='donut-bg' cx='90' cy='90' r='66'></circle>"
        + "".join(circles)
        + "</svg><div class='donut-center'><strong>100%</strong><span>RÉSEAU</span></div></div>"
        "<div class='donut-legend'>"
        + "".join(legend)
        + "</div></div>"
    )


def display_daily_area(records: list[dict]) -> None:
    if not records:
        return
    ordered = sorted(records, key=lambda row: row["date"])
    if len(ordered) == 1:
        row = ordered[0]
        st.html(
            "<div class='viz-card single-day-card'>"
            "<div class='viz-heading'><span>MESURE JOURNALIÈRE</span>"
            f"<strong>{row['date'].strftime('%d/%m/%Y')}</strong></div>"
            "<div style='display:grid;place-content:center;text-align:center;height:180px'>"
            f"<div style='font:700 2.7rem Space Grotesk;color:#0b1f33'>{row['nur']:.2f}</div>"
            "<div style='color:#f36b16;font-size:.68rem;font-weight:700;letter-spacing:.14em'>"
            "NUR RÉSEAU</div>"
            f"<div style='color:#94a3b8;margin-top:.55rem'>{format_duration(row['unavailability_seconds'])}"
            " d’indisponibilité cumulée</div></div></div>"
        )
        return
    values = [row["nur"] for row in ordered]
    minimum, maximum = min(values), max(values)
    span = maximum - minimum or 1
    count = max(len(values) - 1, 1)
    points = [
        (20 + index / count * 560, 175 - (value - minimum) / span * 125)
        for index, value in enumerate(values)
    ]
    point_text = " ".join(f"{x:.1f},{y:.1f}" for x, y in points)
    area_text = f"20,190 {point_text} {points[-1][0]:.1f},190"
    dots = "".join(
        f"<circle cx='{x:.1f}' cy='{y:.1f}' r='4'><title>"
        f"{ordered[index]['date'].strftime('%d/%m')}: NUR {ordered[index]['nur']:.4f}"
        "</title></circle>"
        for index, (x, y) in enumerate(points)
    )
    st.html(
        "<div class='viz-card area-card'>"
        "<div class='viz-heading'><span>ÉVOLUTION JOURNALIÈRE</span>"
        "<strong>Tendance du NUR réseau</strong></div>"
        "<svg class='area-chart' viewBox='0 0 600 210'>"
        "<defs><linearGradient id='areaGradient' x1='0' y1='0' x2='0' y2='1'>"
        "<stop offset='0%' stop-color='#f36b16' stop-opacity='.38'/>"
        "<stop offset='100%' stop-color='#f36b16' stop-opacity='0'/></linearGradient></defs>"
        f"<polygon points='{area_text}' fill='url(#areaGradient)'/>"
        f"<polyline points='{point_text}'/>{dots}</svg>"
        f"<div class='axis-caption'><span>{ordered[0]['date'].strftime('%d/%m/%Y')}</span>"
        f"<span>{ordered[-1]['date'].strftime('%d/%m/%Y')}</span></div></div>"
    )


def display_region_technology_heatmap(records: list[dict]) -> None:
    technologies = sorted({row["technology"] for row in records})
    totals: dict[tuple[str, str], float] = {}
    for row in records:
        key = (row["region"], row["technology"])
        totals[key] = totals.get(key, 0.0) + row["unavailability_hours"]
    regions = sorted(
        {row["region"] for row in records},
        key=lambda region: sum(
            totals.get((region, technology), 0) for technology in technologies
        ),
        reverse=True,
    )[:10]
    maximum = max(totals.values(), default=1) or 1
    header = "".join(f"<th>{escape(technology)}</th>" for technology in technologies)
    rows = []
    for region in regions:
        cells = []
        for technology in technologies:
            value = totals.get((region, technology), 0)
            intensity = 0.08 + value / maximum * 0.82
            cells.append(
                f"<td style='--heat:{intensity:.3f}'><strong>{value:,.0f}</strong>"
                "<small>heures</small></td>".replace(",", " ")
            )
        rows.append(f"<tr><th>{escape(region)}</th>{''.join(cells)}</tr>")
    st.html(
        "<div class='viz-card heatmap-card'><div class='viz-heading'>"
        "<span>MATRICE D’IMPACT</span><strong>Régions × technologies</strong></div>"
        f"<table class='heatmap'><thead><tr><th>Région</th>{header}</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table></div>"
    )


st.html(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
    :root { --ink:#f8fafc; --muted:#94a3b8; --panel:rgba(15,23,42,.72); --line:rgba(148,163,184,.14); --orange:#ff7900; --cyan:#22d3ee; --violet:#a78bfa; }
    .stApp {
        background:
          radial-gradient(circle at 12% 4%, rgba(255,121,0,.15), transparent 28rem),
          radial-gradient(circle at 88% 18%, rgba(34,211,238,.10), transparent 32rem),
          linear-gradient(145deg,#050816 0%,#090e1d 58%,#070b14 100%);
        color:var(--ink); font-family:'DM Sans',sans-serif;
    }
    [data-testid="stHeader"] { background:transparent; }
    [data-testid="stMainBlockContainer"] { max-width:1480px; padding-top:2.2rem; }
    [data-testid="stSidebar"] { background:rgba(4,7,17,.88); border-right:1px solid var(--line); backdrop-filter:blur(24px); }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] label { color:#e2e8f0 !important; }
    h1,h2,h3 { font-family:'Space Grotesk',sans-serif !important; letter-spacing:-.035em; }
    h1 { font-size:clamp(2rem,4vw,3.45rem) !important; background:linear-gradient(90deg,#fff 15%,#ffb36b 72%,#ff7900); -webkit-background-clip:text; color:transparent !important; }
    p, .stCaption { color:var(--muted) !important; }
    button, [data-testid="stDownloadButton"] button { border-radius:12px !important; transition:transform .25s ease,box-shadow .25s ease !important; }
    button:hover { transform:translateY(-2px); box-shadow:0 10px 28px rgba(255,121,0,.18); }
    [data-baseweb="tab-list"] { gap:.45rem; background:rgba(15,23,42,.62); border:1px solid var(--line); padding:.38rem; border-radius:16px; backdrop-filter:blur(16px); }
    [data-baseweb="tab"] { border-radius:11px; padding:.68rem 1rem; color:#94a3b8; }
    [aria-selected="true"] { background:linear-gradient(135deg,rgba(255,121,0,.22),rgba(255,121,0,.08)) !important; color:#fff !important; }
    [data-baseweb="tab-highlight"], [data-baseweb="tab-border"] { display:none; }
    .kpi-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:1rem; margin:1.35rem 0 1rem; }
    .kpi-card { position:relative; overflow:hidden; min-height:155px; padding:1.3rem; border:1px solid var(--line); border-radius:20px; background:linear-gradient(145deg,rgba(30,41,59,.88),rgba(15,23,42,.62)); backdrop-filter:blur(18px); animation:riseIn .65s both; box-shadow:0 20px 60px rgba(0,0,0,.18); }
    .kpi-card:nth-child(2){animation-delay:.08s}.kpi-card:nth-child(3){animation-delay:.16s}.kpi-card:nth-child(4){animation-delay:.24s}
    .kpi-card:before { content:""; position:absolute; inset:0; background:linear-gradient(110deg,transparent 25%,rgba(255,255,255,.055) 48%,transparent 70%); transform:translateX(-120%); animation:shine 5s 1.5s infinite; }
    .kpi-card.primary { grid-column:span 1; border-color:rgba(255,121,0,.38); background:linear-gradient(145deg,rgba(255,121,0,.20),rgba(41,22,10,.58)); }
    .kpi-glow { position:absolute; width:90px; height:90px; right:-20px; top:-25px; border-radius:50%; background:var(--violet); filter:blur(45px); opacity:.18; }
    .kpi-card.orange .kpi-glow,.kpi-card.primary .kpi-glow{background:var(--orange)}.kpi-card.cyan .kpi-glow{background:var(--cyan)}
    .kpi-label { color:#94a3b8; font-size:.72rem; font-weight:700; letter-spacing:.12em; }
    .kpi-value { font:700 clamp(1.65rem,2.4vw,2.45rem)/1.1 'Space Grotesk'; margin:.72rem 0 .55rem; color:#fff; }
    .kpi-note { color:#94a3b8; font-size:.76rem; }
    .tech-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:1rem; margin:0 0 1.7rem; }
    .tech-card { display:flex; align-items:center; gap:1rem; padding:1.05rem 1.15rem; border:1px solid var(--line); border-radius:17px; background:rgba(15,23,42,.62); backdrop-filter:blur(16px); animation:riseIn .6s both; }
    .tech-icon { display:grid; place-items:center; width:48px; height:48px; flex:0 0 48px; border-radius:14px; background:color-mix(in srgb,var(--tech-color) 18%,transparent); border:1px solid color-mix(in srgb,var(--tech-color) 48%,transparent); color:var(--tech-color); font:700 .9rem 'Space Grotesk'; }
    .tech-data { flex:1; min-width:0; }.tech-data span{display:block;color:#64748b;font-size:.65rem;font-weight:700;letter-spacing:.12em}.tech-data strong{font:700 1.35rem 'Space Grotesk';color:#fff}.tech-data small{display:block;color:#94a3b8;margin-top:.1rem}
    .tech-meter { height:3px; background:rgba(148,163,184,.13); margin-top:.6rem; overflow:hidden;border-radius:9px}.tech-meter i{display:block;height:100%;width:var(--meter-width);background:var(--tech-color);animation:growBar 1s ease both}
    .viz-card,.bar-chart,.nur-table-wrap { border:1px solid var(--line); background:var(--panel); backdrop-filter:blur(18px); box-shadow:0 24px 70px rgba(0,0,0,.18); }
    .viz-card { border-radius:20px; padding:1.25rem; min-height:250px; animation:riseIn .7s both; }
    .viz-heading span{display:block;color:#64748b;font-size:.65rem;font-weight:700;letter-spacing:.12em}.viz-heading strong{display:block;color:#f8fafc;font:600 1.05rem 'Space Grotesk';margin-top:.25rem}
    .donut-layout { display:grid; grid-template-columns:1fr 190px 1fr; align-items:center; gap:1rem; }.donut-wrap{position:relative;width:190px;height:190px}.donut-wrap svg{transform:rotate(-90deg)}.donut-bg{fill:none;stroke:rgba(148,163,184,.1);stroke-width:18}.donut-segment{fill:none;stroke-width:18;animation:drawDonut 1.2s ease both}.donut-center{position:absolute;inset:0;display:grid;place-content:center;text-align:center}.donut-center strong{font:700 1.5rem 'Space Grotesk';color:#fff}.donut-center span{font-size:.6rem;letter-spacing:.15em;color:#64748b}.donut-legend-row{display:grid;grid-template-columns:9px 1fr auto;gap:.6rem;align-items:center;margin:.7rem 0;color:#94a3b8}.donut-legend-row i{width:8px;height:8px;border-radius:50%}.donut-legend-row strong{color:#fff}
    .area-card{min-height:290px}.area-chart{width:100%;height:215px;overflow:visible}.area-chart polyline{fill:none;stroke:#ff7900;stroke-width:4;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:1200;animation:drawLine 1.5s ease both}.area-chart circle{fill:#050816;stroke:#ffb36b;stroke-width:3}.axis-caption{display:flex;justify-content:space-between;color:#64748b;font-size:.72rem}
    .heatmap-card{overflow-x:auto}.heatmap{width:100%;border-collapse:separate;border-spacing:6px;margin-top:1rem}.heatmap th{color:#94a3b8;font-size:.72rem;text-align:left;padding:.45rem}.heatmap td{min-width:90px;padding:.8rem;border-radius:10px;text-align:center;background:rgba(255,121,0,var(--heat));color:#fff;transition:transform .2s}.heatmap td:hover{transform:scale(1.06)}.heatmap td strong,.heatmap td small{display:block}.heatmap td small{font-size:.62rem;color:#cbd5e1}
    .nur-table-wrap { overflow:auto; max-height:620px; border-radius:16px; }
    .nur-table { border-collapse:collapse; width:100%; font-size:.82rem; }
    .nur-table th { position:sticky; top:0; background:#111827; color:#f8fafc; z-index:1; }
    .nur-table th,.nur-table td { padding:10px 13px; border-bottom:1px solid var(--line); text-align:left; white-space:nowrap; }
    .nur-table td{color:#cbd5e1}.nur-table tr:hover{background:rgba(255,121,0,.055)}
    .bar-chart { padding:1.15rem; border-radius:18px; }
    .bar-row { display:grid; grid-template-columns:170px 1fr 105px; gap:12px; align-items:center; margin:10px 0; }
    .bar-label,.bar-value { font-size:.8rem; color:#cbd5e1; }.bar-value{text-align:right;font-variant-numeric:tabular-nums;color:#f8fafc}
    .bar-track { height:10px; background:rgba(148,163,184,.1); border-radius:99px; overflow:hidden; }
    .bar-fill { display:block;height:100%;width:var(--bar-width);background:linear-gradient(90deg,#ff7900,#ffb86b);border-radius:99px;transform-origin:left;animation:growBar 1s cubic-bezier(.22,1,.36,1) both;box-shadow:0 0 16px rgba(255,121,0,.35)}
    @keyframes riseIn{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:none}}
    @keyframes growBar{from{width:0}to{width:var(--bar-width,var(--meter-width))}}
    @keyframes drawLine{from{stroke-dashoffset:1200}to{stroke-dashoffset:0}}
    @keyframes drawDonut{from{opacity:0;stroke-dashoffset:100}to{opacity:1}}
    @keyframes shine{0%,60%{transform:translateX(-120%)}80%,100%{transform:translateX(120%)}}
    @media(max-width:900px){.kpi-grid{grid-template-columns:1fr 1fr}.tech-grid{grid-template-columns:1fr}.donut-layout{grid-template-columns:1fr;justify-items:center}.viz-heading,.donut-legend{width:100%}.bar-row{grid-template-columns:110px 1fr 80px}}
    </style>
    """
)

st.html(
    """
    <style>
    :root {
      --ink:#191919; --muted:#6f716f; --soft:#a2a39f; --paper:#f4f1eb;
      --panel:#fffdf9; --line:#ded9d0; --orange:#f36b16; --sage:#47665a;
    }
    * { box-sizing:border-box; }
    .stApp {
      background:
        radial-gradient(circle at 92% 0%,rgba(243,107,22,.09),transparent 30rem),
        linear-gradient(180deg,#f8f6f1 0%,#f1eee7 100%);
      color:var(--ink); font-family:'DM Sans',sans-serif;
    }
    [data-testid="stMainBlockContainer"] { max-width:1440px; padding:2.6rem 2.25rem 4rem; }
    [data-testid="stHeader"] { background:rgba(248,246,241,.78); backdrop-filter:blur(14px); }
    [data-testid="stSidebar"] {
      background:#ebe7de; border-right:1px solid #d9d3c8; backdrop-filter:none;
    }
    [data-testid="stSidebar"] h2,[data-testid="stSidebar"] label { color:#242424 !important; }
    [data-testid="stSidebar"] p { color:#71736f !important; }
    h1,h2,h3 { color:#1b1b1b !important; }
    h1 {
      font-size:clamp(2.2rem,4vw,3.8rem) !important; letter-spacing:-.055em !important;
      background:none !important; -webkit-background-clip:initial !important;
    }
    p,.stCaption { color:#71736f !important; }
    [data-baseweb="tab-list"] {
      gap:.25rem; background:#e9e5dd; border:1px solid #d9d3c8;
      padding:.3rem; border-radius:13px; backdrop-filter:none;
    }
    [data-baseweb="tab"] { min-height:42px; border-radius:9px; color:#73746f; padding:.55rem .9rem; }
    [aria-selected="true"] {
      background:#fffdf9 !important; color:#1b1b1b !important;
      box-shadow:0 2px 8px rgba(32,29,24,.08) !important;
    }
    .kpi-grid {
      display:grid; grid-template-columns:repeat(12,minmax(0,1fr)); gap:14px;
      margin:1.45rem 0 14px; align-items:stretch;
    }
    .kpi-card {
      grid-column:span 3 !important; display:flex; flex-direction:column;
      justify-content:space-between; min-height:168px; height:100%; padding:1.25rem;
      border:1px solid #ded9d0; border-radius:16px; background:#fffdf9 !important;
      backdrop-filter:none; box-shadow:0 8px 28px rgba(50,42,32,.055);
      animation:cleanRise .5s both;
    }
    .kpi-card.primary {
      background:#1d1d1b !important; border-color:#1d1d1b;
      box-shadow:0 12px 32px rgba(24,24,22,.14);
    }
    .kpi-card:before,.kpi-glow { display:none; }
    .kpi-label { color:#85857f; font-size:.68rem; letter-spacing:.11em; }
    .kpi-value { color:#1b1b1b; font-size:clamp(1.65rem,2.25vw,2.25rem); margin:.7rem 0; }
    .kpi-note { color:#777873; font-size:.73rem; line-height:1.4; }
    .kpi-card.primary .kpi-value { color:#fff; }
    .kpi-card.primary .kpi-label { color:#ff9c5f; }
    .kpi-card.primary .kpi-note { color:#b7b7b1; }
    .tech-grid {
      display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px;
      align-items:stretch; margin:0 0 1.6rem;
    }
    .tech-card {
      min-height:120px; height:100%; align-items:flex-start; padding:1.15rem;
      border:1px solid #ded9d0; border-radius:16px; background:#fffdf9;
      backdrop-filter:none; box-shadow:none;
    }
    .tech-icon {
      width:46px; height:46px; flex:0 0 46px; border-radius:12px;
      background:#f1ede5; border:1px solid #ddd6ca; color:#292929;
    }
    .tech-data span { color:#8a8b85; }
    .tech-data strong { color:#1b1b1b; }
    .tech-data small { color:#73746f; line-height:1.35; min-height:20px; }
    .tech-meter { background:#ece7de; height:4px; }
    .viz-card,.bar-chart,.nur-table-wrap {
      border:1px solid #ded9d0; background:#fffdf9; backdrop-filter:none;
      box-shadow:0 8px 28px rgba(50,42,32,.045);
    }
    .viz-card { border-radius:16px; padding:1.25rem; }
    .viz-heading span { color:#999891; }
    .viz-heading strong { color:#1b1b1b; }
    .donut-bg { stroke:#ebe6dd; }
    .donut-center strong { color:#1b1b1b; }
    .donut-center span { color:#8a8b85; }
    .donut-legend-row { color:#72736e; }
    .donut-legend-row strong { color:#1b1b1b; }
    .area-chart polyline { stroke:#f36b16; }
    .area-chart circle { fill:#fffdf9; stroke:#f36b16; }
    .heatmap th { color:#777873; }
    .heatmap td { color:#fff; background:rgba(243,107,22,var(--heat)); }
    .heatmap td small { color:rgba(255,255,255,.8); }
    .nur-table-wrap { border-radius:14px; }
    .nur-table th { background:#252523; color:#fff; }
    .nur-table td { color:#454640; border-color:#ebe6dd; }
    .nur-table tr:hover { background:#faf2e9; }
    .bar-chart { border-radius:16px; padding:1.15rem; }
    .bar-row { grid-template-columns:280px 1fr 72px; }
    .bar-label strong { white-space:normal; }
    .bar-label { color:#565752; display:flex; flex-direction:column; gap:2px; line-height:1.15; }
    .bar-label small { color:#8a8b85; font-size:.68rem; }
    .bar-value { color:#1b1b1b; font-weight:600; }
    .bar-value.bad { color:#c41e3a; }
    .bar-value.ok { color:#0f7b4a; }
    .bar-track { background:#ece7de; height:8px; }
    .bar-fill { background:#ff7900; box-shadow:none; }
    .bar-fill.bad { background:#c41e3a; }
    .bar-fill.ok { background:#0f7b4a; }
    .region-grid {
      display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px;
      margin:0 0 1.4rem;
    }
    .region-card {
      background:#fffdf9; border:1px solid #ded9d0; border-radius:14px; padding:1rem 1.1rem;
    }
    .region-head { display:flex; justify-content:space-between; align-items:baseline; gap:8px; }
    .region-head span { font-weight:700; color:#0b1f33; }
    .region-head strong { font:700 1.35rem 'Space Grotesk'; }
    .region-head strong.bad { color:#c41e3a; }
    .region-head strong.ok { color:#0f7b4a; }
    .region-card small { display:block; color:#73746f; margin:.35rem 0 .55rem; }
    .region-card em { display:block; margin-top:.45rem; font-style:normal; font-size:.72rem; letter-spacing:.08em; text-transform:uppercase; color:#8a8b85; }
    .region-card .bar-track { margin-top:.2rem; }
    button,[data-testid="stDownloadButton"] button {
      border-radius:10px !important; border-color:#cdc6ba !important;
    }
    @keyframes cleanRise {
      from { opacity:0; transform:translateY(10px); }
      to { opacity:1; transform:none; }
    }
    @media(max-width:1100px) {
      .kpi-card { grid-column:span 6 !important; }
      .tech-grid, .region-grid { grid-template-columns:1fr; }
    }
    @media(max-width:680px) {
      [data-testid="stMainBlockContainer"] { padding:1.5rem .9rem 3rem; }
      .kpi-grid { grid-template-columns:1fr; }
      .kpi-card { grid-column:span 12 !important; min-height:145px; }
      .bar-row { grid-template-columns:95px 1fr 70px; gap:8px; }
    }
    </style>
    """
)

st.html(
    "<div style='color:#f36b16;font-size:.72rem;font-weight:700;letter-spacing:.18em;"
    "margin-bottom:.45rem'>NETWORK INTELLIGENCE CENTER</div>"
)
st.title("Observatoire NUR")
st.caption("NUR classé par région (préfixe site) · seuils · top mauvais sites par période.")

with st.sidebar:
    st.header("1. Fichiers")
    sites_file = st.file_uploader(
        "Base des sites déjà filtrée",
        type=["xlsx"],
        help="Les sites à analyser doivent se trouver dans la deuxième feuille.",
    )
    nur_file = st.file_uploader(
        "Fichier NUR 2G / 3G / 4G",
        type=["xlsx"],
        help="Le classeur doit contenir les onglets 2G, 3G et 4G.",
    )

if sites_file is None or nur_file is None:
    st.info("Importez les deux fichiers Excel pour lancer le calcul.")
    st.stop()

try:
    with st.spinner("Lecture et contrôle des fichiers…"):
        imported = load_workbooks_cached(
            sites_file.getvalue(), nur_file.getvalue(), CACHE_SCHEMA_VERSION
        )
except Exception as error:
    st.error(f"Impossible de lire les fichiers : {error}")
    st.stop()

if not imported.data:
    st.error("Aucune ligne NUR admissible n'a été trouvée après les contrôles.")
    display_table(imported.rejected)
    st.stop()

minimum_date = min(row["date"] for row in imported.data)
maximum_date = max(row["date"] for row in imported.data)

with st.sidebar:
    st.header("2. Période")
    period_label = st.selectbox("Type de période", list(PERIODS))
    period_days = PERIODS[period_label]
    selected_start = st.date_input(
        "Date de début",
        value=minimum_date,
        min_value=minimum_date,
        max_value=maximum_date,
        format="DD/MM/YYYY",
    )
    selected_end = selected_start + timedelta(days=period_days - 1)
    st.caption(
        f"Fenêtre : {selected_start.strftime('%d/%m/%Y')} au "
        f"{selected_end.strftime('%d/%m/%Y')}"
    )

calculation = calculate_cached(
    imported.data, selected_start, period_days, CACHE_SCHEMA_VERSION
)
if calculation.selected_rows == 0:
    st.warning("Aucune donnée n'est présente dans la période sélectionnée.")
    st.stop()

with st.spinner("Préparation des rapports Excel et PowerPoint…"):
    report = build_excel_report(calculation, imported.rejected)
    powerpoint = build_powerpoint_report(calculation, imported)

excel_name = (
    f"rapport_nur_{calculation.start_date.isoformat()}_"
    f"{calculation.period_days}j.xlsx"
)
ppt_name = (
    f"presentation_nur_{calculation.start_date.isoformat()}_"
    f"{calculation.period_days}j.pptx"
)

all_sites = sorted({row["site_code"] for row in calculation.cells})
all_technologies = sorted({row["technology"] for row in calculation.cells})
all_regions = [
    name for name in REGION_ORDER
    if any(row["region"] == name for row in calculation.sites)
]
other_regions = sorted(
    {row["region"] for row in calculation.sites if row["region"] not in REGION_ORDER}
)
all_regions.extend(other_regions)
with st.sidebar:
    st.header("3. Affichage")
    selected_technologies = st.multiselect(
        "Technologies", all_technologies, default=all_technologies
    )
    selected_regions = st.multiselect(
        "Régions", all_regions, default=all_regions
    )
    selected_sites = st.multiselect(
        "Sites (optionnel)", all_sites, placeholder="Tous les sites"
    )
    top_limit = st.slider("Nombre de sites dans les tops", 5, 30, 12)
    st.header("4. Rapports")
    st.caption("Excel détaillé et présentation PowerPoint de la période.")
    st.download_button(
        "Excel",
        data=report,
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="excel_sidebar",
    )
    st.download_button(
        "PowerPoint",
        data=powerpoint,
        file_name=ppt_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True,
        key="ppt_sidebar",
    )

filtered_sites = filtered_records(
    calculation.sites, selected_sites, selected_technologies, selected_regions
)

network_row = calculation.network[0]
display_kpi_dashboard(calculation)
display_technology_nur(calculation.technologies)
display_region_gauges(calculation.regions)

st.caption(
    f"Analyse du {calculation.start_date.strftime('%d/%m/%Y')} au "
    f"{calculation.end_date.strftime('%d/%m/%Y')} avec un dénominateur de "
    f"{calculation.period_days} jour(s) • Préfixes KN/KT/KO-KE/KC/KV/PO-EQ"
    f" • Base : « {imported.database_sheet_name} »."
)

st.subheader("Exporter les rapports")
st.caption("Les fichiers reprennent la période et les seuils régionaux ci-dessus.")
download_col_1, download_col_2 = st.columns(2)
with download_col_1:
    st.download_button(
        "Télécharger le rapport Excel",
        data=report,
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="excel_main",
    )
with download_col_2:
    st.download_button(
        "Télécharger la présentation PowerPoint",
        data=powerpoint,
        file_name=ppt_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True,
        key="ppt_main",
    )

tab_network, tab_worst, tab_charts, tab_sites, tab_technologies, tab_cells, tab_quality = st.tabs(
    [
        "Vue exécutive",
        "Top mauvais sites",
        "Exploration visuelle",
        "Sites",
        "Technologies",
        "Cellules",
        "Qualité",
    ]
)
with tab_network:
    st.subheader("NUR total de tous les sites")
    display_table(calculation.network)
    st.subheader("NUR total par région")
    display_table(calculation.regions)
    st.subheader("NUR total par technologie")
    display_table(calculation.technologies)

with tab_worst:
    st.subheader("10 pires sites par région")
    st.caption(
        "Classement du NUR le plus élevé dans chaque région. "
        "Rouge = au-dessus du seuil. "
        "Kinshasa 800 · Katanga 600 · Kasaï 500 · Kongo Central 300 · Kivu 1000 · Grand Nord 1000."
    )
    region_cols = st.columns(2)
    for index, region in enumerate(selected_regions or all_regions):
        with region_cols[index % 2]:
            threshold = REGION_NUR_THRESHOLDS.get(region)
            threshold_label = f"{threshold:.0f}" if threshold is not None else "—"
            st.markdown(f"**{region}** — seuil {threshold_label}")
            local = top_worst_sites(filtered_sites, region=region, limit=10)
            display_top_sites(local, limit=10)

    st.subheader("Période sélectionnée")
    st.caption("Classement national, tous sites confondus.")
    display_top_sites(filtered_sites, limit=top_limit)

    st.subheader("Hors seuil")
    hors = [
        row for row in filtered_sites if row.get("above_threshold")
    ]
    if hors:
        display_table(sorted(hors, key=lambda row: row["nur"], reverse=True))
    else:
        st.success("Aucun site au-dessus du seuil dans le filtre actuel.")

    st.subheader("Comparer jour / semaine / mois")
    st.caption("Même date de début, trois dénominateurs. Le top est national, non filtré.")
    display_period_tops(imported.data, selected_start, CACHE_SCHEMA_VERSION, limit=8)

with tab_charts:
    chart_col_1, chart_col_2 = st.columns(2)
    with chart_col_1:
        display_technology_donut(calculation.technologies)
    with chart_col_2:
        display_daily_area(calculation.daily)

    st.subheader("Empreinte régionale")
    st.caption("Somme des heures-cellules indisponibles pour chaque région.")
    display_duration_bars(calculation.regions, "region", limit=12)

    st.subheader("Où se concentre l’impact ?")
    display_region_technology_heatmap(calculation.site_technology)

    st.subheader("Sites les plus impactés (durée)")
    display_duration_bars(filtered_sites, "site_code", limit=15)

with tab_sites:
    records = filtered_records(
        calculation.sites, selected_sites, selected_technologies, selected_regions
    )
    display_table(sorted(records, key=lambda row: row["nur"], reverse=True))

with tab_technologies:
    st.subheader("Totaux par technologie")
    display_table(calculation.technologies)
    st.subheader("Détail par site et technologie")
    records = filtered_records(
        calculation.site_technology, selected_sites, selected_technologies, selected_regions
    )
    display_table(records)

with tab_cells:
    records = filtered_records(
        calculation.cells, selected_sites, selected_technologies, selected_regions
    )
    display_table(records)

with tab_quality:
    quality_columns = st.columns(4)
    quality_columns[0].metric("Sites admissibles en base", imported.eligible_site_count)
    quality_columns[1].metric("Feuille de base", imported.database_sheet_name)
    quality_columns[2].metric("Lignes NUR importées", imported.total_nur_rows)
    quality_columns[3].metric("Lignes exclues", len(imported.rejected))

    if not imported.rejected:
        st.success("Aucune ligne exclue.")
    else:
        category_counts: dict[str, int] = {}
        for row in imported.rejected:
            category = row["category"]
            category_counts[category] = category_counts.get(category, 0) + 1
        summary = [
            {"category": category, "Nombre": count}
            for category, count in sorted(category_counts.items())
        ]
        st.subheader("Résumé des exclusions")
        display_table(summary)
        with st.expander("Afficher le détail des exclusions"):
            display_table(imported.rejected)
