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

# El tramo que quieres sacar en el PDF:
# Valores típicos: "Pretemporada", "1a Mitad", "2a Mitad"
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


def _save_bar(series: pd.Series, title: str, outpath: str, top_n: int = 12):
    series = series.dropna()
    if series.empty:
        plt.figure(figsize=(8, 4))
        plt.title(title)
        plt.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        plt.axis("off")
        plt.tight_layout()
        plt.savefig(outpath, dpi=200)
        plt.close()
        return
    counts = series.value_counts().head(top_n).sort_values(ascending=True)
    plt.figure(figsize=(8, 4))
    plt.title(title)
    plt.barh(counts.index.astype(str), counts.values)
    plt.tight_layout()
    plt.savefig(outpath, dpi=200)
    plt.close()


def _kpis(df: pd.DataFrame, col_player: str, col_days: str, col_rec: str, col_type: str):
    total = len(df)
    players = df[col_player].dropna().nunique() if col_player else 0
    days = pd.to_numeric(df[col_days], errors="coerce").fillna(0).sum() if col_days else 0
    avg_days = (days / total) if total else 0

    recurr = 0
    pct_recurr = 0
    if col_rec and total:
        recurr = (df[col_rec].fillna("").str.strip().str.lower() != "primera lesión").sum()
        pct_recurr = recurr / total

    muscular = 0
    pct_mus = 0
    if col_type and total:
        muscular = df[col_type].fillna("").str.lower().str.contains("mus").sum()
        pct_mus = muscular / total

    return {
        "Tramo": TRAMO_OBJETIVO,
        "Nº lesiones (tramo)": total,
        "Nº jugadores lesionados": players,
        "Días totales de baja": int(days),
        "Media días / lesión": round(avg_days, 1),
        "% musculares": f"{pct_mus:.1%}",
        "% recurrencias": f"{pct_recurr:.1%}",
    }


def _top_rows(df: pd.DataFrame, col: str, n: int):
    if not col:
        return [["(No existe la columna en el Excel)", ""]]
    s = df[col].fillna("").replace("", "Sin especificar")
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
    col_days = _col(df, "Días de baja", "Dias de baja", "Dias baja")

    # ---------- FILTRO POR TRAMO ----------
    if not col_tramo:
        # Sin tramo -> vacío para forzar consistencia
        df = df.iloc[0:0].copy()
    else:
        # Normalizamos texto
        tramo = df[col_tramo].fillna("").astype(str).str.strip()
        df = df[tramo == TRAMO_OBJETIVO].copy()

    # ---------- CHARTS ----------
    os.makedirs("tmp_charts", exist_ok=True)
    c_tipo = "tmp_charts/tipo.png"
    c_zona = "tmp_charts/zona.png"
    c_mec  = "tmp_charts/mecanismo.png"
    c_sev  = "tmp_charts/severidad.png"
    c_cat  = "tmp_charts/categoria.png"
    c_pos  = "tmp_charts/posicion.png"

    _save_bar(df[col_type] if col_type else pd.Series(dtype=str), "Lesiones por tipo", c_tipo)
    _save_bar(df[col_loc] if col_loc else pd.Series(dtype=str), "Lesiones por localización anatómica", c_zona)
    _save_bar(df[col_mec] if col_mec else pd.Series(dtype=str), "Lesiones por mecanismo lesional", c_mec)
    _save_bar(df[col_sev] if col_sev else pd.Series(dtype=str), "Lesiones por severidad", c_sev)
    _save_bar(df[col_cat] if col_cat else pd.Series(dtype=str), "Lesiones por categoría", c_cat)
    _save_bar(df[col_pos] if col_pos else pd.Series(dtype=str), "Lesiones por posición", c_pos)

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

    # Logo centered
    if os.path.exists(LOGO_PATH):
        logo = Image(LOGO_PATH, width=6*cm, height=6*cm)
        logo.hAlign = "CENTER"
        story.append(logo)
        story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph(CLB_TITLE := f"{CLUB_TITLE} – {TRAMO_OBJETIVO}", styles["Title"]))
    story.append(Paragraph(f"Generado automáticamente · {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 0.4*cm))

    story.append(Paragraph("Resumen", styles["Heading2"]))
    kpi_table = [["KPI", "Valor"]] + [[k, str(v)] for k, v in _kpis(df, col_player, col_days, col_rec, col_type).items()]
    t = Table(kpi_table, colWidths=[9.5*cm, 6.0*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph("Gráficos", styles["Heading2"]))
    for img in [c_tipo, c_zona, c_mec, c_sev, c_cat, c_pos]:
        story.append(Image(img, width=17*cm, height=8*cm))
        story.append(Spacer(1, 0.25*cm))

    story.append(PageBreak())
    story.append(Paragraph(f"TOP {TOP_N}", styles["Heading2"]))

    blocks = [
        ("Localización anatómica", col_loc),
        ("Tipo de lesión", col_type),
        ("Diagnóstico", col_diag),
        ("Categoría", col_cat),
        ("Posición", col_pos),
        ("Mecanismo lesional", col_mec),
        ("Severidad", col_sev),
    ]
    for title, col in blocks:
        rows = [[f"{title} (Top {TOP_N})", "Nº"]] + _top_rows(df, col, TOP_N)
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
