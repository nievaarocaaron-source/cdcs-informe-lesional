import sys
import os
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

# ==========================
# CONFIG (cambia SOLO esto)
# ==========================
CLUB_TITLE = "CDCS Academy – Informe lesional"
LOGO_PATH = "assets/logo_cdcs.png"

# Valores típicos en tu Excel (columna "Tramo temporada"):
# "Pretemporada", "1a Mitad", "2a Mitad"
TRAMO_OBJETIVO = "1a Mitad"

TOP_N = 5


def _col(df, *names):
    """Devuelve la primera columna existente por nombre (case-insensitive)."""
    lower_map = {c.lower().strip(): c for c in df.columns}
    for n in names:
        key = str(n).lower().strip()
        if key in lower_map:
            return lower_map[key]
    return None


def _series(df, col):
    if not col:
        return pd.Series(dtype=str)
    return df[col].astype(str).replace({"nan": ""}).fillna("")


# ---------- CHARTS ----------
def chart_barh_top(series: pd.Series, title: str, outpath: str, top_n: int = 8):
    s = series.replace("", "Sin especificar").dropna()
    if s.empty:
        _chart_empty(title, outpath); return
    counts = s.value_counts().head(top_n).sort_values(ascending=True)
    plt.figure(figsize=(8.5, 4.2))
    plt.title(title)
    plt.barh(counts.index.astype(str), counts.values)
    plt.tight_layout()
    plt.savefig(outpath, dpi=200)
    plt.close()


def chart_barv_top(series: pd.Series, title: str, outpath: str, top_n: int = 10):
    s = series.replace("", "Sin especificar").dropna()
    if s.empty:
        _chart_empty(title, outpath); return
    counts = s.value_counts().head(top_n)
    plt.figure(figsize=(8.5, 4.2))
    plt.title(title)
    plt.bar(counts.index.astype(str), counts.values)
    plt.xticks(rotation=25, ha="right")
    plt.tight_layout()
    plt.savefig(outpath, dpi=200)
    plt.close()


def chart_donut(series: pd.Series, title: str, outpath: str, top_n: int = 6):
    s = series.replace("", "Sin especificar").dropna()
    if s.empty:
        _chart_empty(title, outpath); return
    counts = s.value_counts().head(top_n)
    plt.figure(figsize=(7.2, 4.6))
    plt.title(title)
    plt.pie(
        counts.values,
        labels=counts.index.astype(str),
        autopct=lambda p: f"{p:.0f}%",
        startangle=90,
        counterclock=False,
        wedgeprops=dict(width=0.45)  # donut
    )
    plt.axis("equal")
    plt.tight_layout()
    plt.savefig(outpath, dpi=200)
    plt.close()


def chart_stacked_severity(series: pd.Series, title: str, outpath: str, order=None):
    s = series.replace("", "Sin especificar").dropna()
    if s.empty:
        _chart_empty(title, outpath); return

    counts = s.value_counts()
    if order:
        counts = counts.reindex([o for o in order if o in counts.index]).fillna(0).astype(int)

    plt.figure(figsize=(8.5, 2.8))
    plt.title(title)

    bottom = 0
    for label, val in counts.items():
        plt.bar(["Severidad"], [val], bottom=bottom, label=str(label))
        bottom += val

    plt.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), frameon=False)
    plt.tight_layout()
    plt.savefig(outpath, dpi=200)
    plt.close()


def _chart_empty(title: str, outpath: str):
    plt.figure(figsize=(8, 4))
    plt.title(title)
    plt.text(0.5, 0.5, "Sin datos", ha="center", va="center")
    plt.axis("off")
    plt.tight_layout()
    plt.savefig(outpath, dpi=200)
    plt.close()


# ---------- TABLES / KPI ----------
def _kpis(df: pd.DataFrame, col_player: str, col_rec: str, col_type: str):
    total = len(df)
    players = df[col_player].dropna().nunique() if col_player else 0

    recurr = 0
    pct_recurr = 0
    if col_rec and total:
        recurr = (df[col_rec].fillna("").astype(str).str.strip().str.lower() != "primera lesión").sum()
        pct_recurr = recurr / total

    muscular = 0
    pct_mus = 0
    if col_type and total:
        muscular = df[col_type].fillna("").astype(str).str.lower().str.contains("mus").sum()
        pct_mus = muscular / total

    return {
        "Tramo": TRAMO_OBJETIVO,
        "Nº lesiones (tramo)": total,
        "Nº jugadores lesionados": players,
        "% musculares": f"{pct_mus:.1%}",
        "% recurrencias": f"{pct_recurr:.1%}",
    }


def top_table(df: pd.DataFrame, col: str, n: int):
    if not col:
        return [["(No existe la columna en el Excel)", ""]]
    s = df[col].fillna("").astype(str).str.strip().replace("", "Sin especificar")
    top = s.value_counts().head(n)
    return [[idx, int(val)] for idx, val in top.items()]


def main(xlsx_path: str, out_pdf: str):
    df = pd.read_excel(xlsx_path, sheet_name="Base de datos", engine="openpyxl")

    # Column mapping (tolerante)
    col_tramo = _col(df, "Tramo temporada", "Tramo", "Periodo", "Mitad")
    col_player = _col(df, "Nombre jugador", "Jugador", "Nombre")
    col_loc = _col(df, "Localización anatómica", "Localizacion anatómica", "Localizacion anatomica")
    col_type = _col(df, "Tipo de lesión", "Tipo de lesion")
    col_sev = _col(df, "Grado / severidad", "Grado/severidad", "Grado severidad")
    col_mec = _col(df, "Mecanismo lesional", "Mecanismo")
    col_rec = _col(df, "Recurrencia")
    col_diag = _col(df, "Diagnóstico", "Diagnostico")
    col_cat = _col(df, "Categoría", "Categoria")
    col_pos = _col(df, "Posición", "Posicion")

    # ---------- FILTRO POR TRAMO ----------
    if not col_tramo:
        df = df.iloc[0:0].copy()
    else:
        tramo = df[col_tramo].fillna("").astype(str).str.strip()
        df = df[tramo == TRAMO_OBJETIVO].copy()

    # ---------- CHARTS (propuesta nueva) ----------
    os.makedirs("tmp_charts", exist_ok=True)
    c_loc = "tmp_charts/loc_barh.png"
    c_type = "tmp_charts/type_barv.png"
    c_mec = "tmp_charts/mec_donut.png"
    c_sev = "tmp_charts/sev_stacked.png"
    c_cat = "tmp_charts/cat_barh.png"

    chart_barh_top(_series(df, col_loc), "Localización anatómica (TOP)", c_loc, top_n=8)
    chart_barv_top(_series(df, col_type), "Tipo de lesión (TOP)", c_type, top_n=10)
    chart_donut(_series(df, col_mec), "Mecanismo lesional (distribución)", c_mec, top_n=6)
    chart_stacked_severity(_series(df, col_sev), "Severidad (apilado)", c_sev, order=["Leve", "Moderado", "Moderada", "Grave"])
    chart_barh_top(_series(df, col_cat), "Categoría (TOP)", c_cat, top_n=10)

    # ---------- PDF ----------
    styles = getSampleStyleSheet()
    os.makedirs(os.path.dirname(out_pdf), exist_ok=True)
    doc = SimpleDocTemplate(
        out_pdf,
        pagesize=A4,
        rightMargin=1.5*cm, leftMargin=1.5*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm
    )
    story = []

    # Logo centered (ancho fijo, alto proporcional)
    if os.path.exists(LOGO_PATH):
        logo = Image(LOGO_PATH)
        logo.drawWidth = 6 * cm
        logo.drawHeight = logo.imageHeight * (logo.drawWidth / logo.imageWidth)
        logo.hAlign = "CENTER"
        story.append(logo)
        story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(f"{CLUB_TITLE} – {TRAMO_OBJETIVO}", styles["Title"]))
    story.append(Paragraph(f"Generado automáticamente · {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 0.4*cm))

    # Page 1: KPI
    story.append(Paragraph("Resumen", styles["Heading2"]))
    kpi_data = [["KPI", "Valor"]] + [[k, str(v)] for k, v in _kpis(df, col_player, col_rec, col_type).items()]
    t = Table(kpi_data, colWidths=[9.5*cm, 6.0*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
    ]))
    story.append(t)
    story.append(PageBreak())

    # Page 2: Localización + Tipo (distintos)
    story.append(Paragraph("Distribución principal", styles["Heading2"]))
    story.append(Image(c_loc, width=17*cm, height=8*cm))
    story.append(Spacer(1, 0.3*cm))
    story.append(Image(c_type, width=17*cm, height=8*cm))
    story.append(PageBreak())

    # Page 3: Mecanismo (donut) + Severidad (stacked)
    story.append(Paragraph("Mecanismo y severidad", styles["Heading2"]))
    story.append(Image(c_mec, width=17*cm, height=9*cm))
    story.append(Spacer(1, 0.4*cm))
    story.append(Image(c_sev, width=17*cm, height=5*cm))
    story.append(PageBreak())

    # Page 4: Categoría (barh) + Posición (tabla TOP)
    story.append(Paragraph("Perfil del jugador", styles["Heading2"]))
    story.append(Image(c_cat, width=17*cm, height=8*cm))
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph("Posición (TOP 5)", styles["Heading3"]))
    pos_rows = [["Posición", "Nº"]] + top_table(df, col_pos, 5)
    pos_table = Table(pos_rows, colWidths=[12.5*cm, 3.0*cm])
    pos_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
    ]))
    story.append(pos_table)
    story.append(PageBreak())

    # Page 5: TOPs clínicos (solo tablas)
    story.append(Paragraph(f"TOPs clínicos (Top {TOP_N})", styles["Heading2"]))
    blocks = [
        ("Localización anatómica", col_loc),
        ("Tipo de lesión", col_type),
        ("Diagnóstico", col_diag),
        ("Mecanismo lesional", col_mec),
        ("Severidad", col_sev),
    ]
    for title, col in blocks:
        rows = [[title, "Nº"]] + top_table(df, col, TOP_N)
        tt = Table(rows, colWidths=[12.5*cm, 3.0*cm])
        tt.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ]))
        story.append(Paragraph(title, styles["Heading3"]))
        story.append(tt)
        story.append(Spacer(1, 0.35*cm))

    doc.build(story)
    print(f"OK -> {out_pdf}")


if __name__ == "__main__":
    xlsx = sys.argv[1] if len(sys.argv) > 1 else "data/REGISTRO LESIONAL.xlsx"
    out = sys.argv[2] if len(sys.argv) > 2 else "reports/informe.pdf"
    main(xlsx, out)
