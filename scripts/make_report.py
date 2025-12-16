import sys
import os
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image,
    Table, TableStyle, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet

# ==========================
# CONFIG (cambia SOLO esto)
# ==========================
CLUB_TITLE = "CDCS Academy – Primera mitad de la temporada"
LOGO_PATH = "assets/logo_cdcs.png"

# Primera mitad (AJUSTA FECHAS A TU TEMPORADA)
FIRST_HALF_START = "2025-07-01"
FIRST_HALF_END   = "2025-12-31"

TOP_N = 5  # TOPs para tablas

def _safe_dt(series):
    return pd.to_datetime(series, errors="coerce", dayfirst=True)

def _save_bar_chart(series: pd.Series, title: str, outpath: str, top_n: int = 12):
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

def _save_line_by_month(df: pd.DataFrame, date_col: str, title: str, outpath: str):
    d = df[df[date_col].notna()].copy()
    if d.empty:
        plt.figure(figsize=(8, 4))
        plt.title(title)
        plt.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        plt.axis("off")
        plt.tight_layout()
        plt.savefig(outpath, dpi=200)
        plt.close()
        return
    d[date_col] = _safe_dt(d[date_col])
    d = d[d[date_col].notna()]
    d["Mes"] = d[date_col].dt.to_period("M").dt.to_timestamp()
    m = d.groupby("Mes").size()
    plt.figure(figsize=(8, 4))
    plt.title(title)
    plt.plot(m.index, m.values, marker="o")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    plt.savefig(outpath, dpi=200)
    plt.close()

def _kpi_block(df: pd.DataFrame):
    total = len(df)
    jugadores = df["Nombre jugador"].dropna().nunique() if "Nombre jugador" in df.columns else 0
    dias_baja = pd.to_numeric(df.get("Días de baja", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    media_dias = (dias_baja / total) if total else 0

    recurr = 0
    pct_recurr = 0
    if "Recurrencia" in df.columns and total:
        recurr = (df["Recurrencia"].fillna("").str.strip().str.lower() != "primera lesión").sum()
        pct_recurr = recurr / total

    muscular = 0
    pct_mus = 0
    if "Tipo de lesión" in df.columns and total:
        muscular = df["Tipo de lesión"].fillna("").str.lower().str.contains("mus").sum()
        pct_mus = muscular / total

    return {
        "Nº lesiones (periodo)": total,
        "Nº jugadores lesionados": jugadores,
        "Días totales de baja": int(dias_baja),
        "Media días / lesión": round(media_dias, 1),
        "% musculares": f"{pct_mus:.1%}",
        "% recurrencias": f"{pct_recurr:.1%}",
    }

def _top_rows(df: pd.DataFrame, col: str, n: int):
    if col not in df.columns:
        return [["(No existe la columna en el Excel)", ""]]
    s = df[col].fillna("").replace("", "Sin especificar")
    top = s.value_counts().head(n)
    return [[idx, int(val)] for idx, val in top.items()]

def main(xlsx_path: str, out_pdf: str):
    df = pd.read_excel(xlsx_path, sheet_name="Base de datos", engine="openpyxl")

    # filas válidas: con Fecha lesión
    if "Fecha lesión" in df.columns:
        df = df[df["Fecha lesión"].notna()].copy()
        df["Fecha lesión"] = _safe_dt(df["Fecha lesión"])

    # filtro primera mitad
    start = pd.Timestamp(FIRST_HALF_START)
    end = pd.Timestamp(FIRST_HALF_END)
    if "Fecha lesión" in df.columns:
        df = df[(df["Fecha lesión"] >= start) & (df["Fecha lesión"] <= end)].copy()

    os.makedirs("tmp_charts", exist_ok=True)
    c_tipo = "tmp_charts/tipo.png"
    c_zona = "tmp_charts/zona.png"
    c_mec  = "tmp_charts/mecanismo.png"
    c_sev  = "tmp_charts/severidad.png"
    c_mes  = "tmp_charts/mes.png"

    _save_bar_chart(df.get("Tipo de lesión", pd.Series(dtype=str)), "Lesiones por tipo", c_tipo)
    _save_bar_chart(df.get("Localización anatómica", pd.Series(dtype=str)), "Lesiones por localización anatómica", c_zona)
    _save_bar_chart(df.get("Mecanismo lesional", pd.Series(dtype=str)), "Lesiones por mecanismo lesional", c_mec)
    _save_bar_chart(df.get("Grado / severidad", pd.Series(dtype=str)), "Lesiones por severidad", c_sev)
    if "Fecha lesión" in df.columns:
        _save_line_by_month(df, "Fecha lesión", "Lesiones por mes (periodo)", c_mes)
    else:
        _save_bar_chart(pd.Series(dtype=str), "Lesiones por mes (periodo)", c_mes)

    styles = getSampleStyleSheet()
    os.makedirs(os.path.dirname(out_pdf), exist_ok=True)
    doc = SimpleDocTemplate(out_pdf, pagesize=A4,
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    story = []

    # Logo
    if os.path.exists(LOGO_PATH):
        logo = Image(LOGO_PATH, width=6*cm, height=6*cm)
        logo.hAlign = "LEFT"
        story.append(logo)
        story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph(CLUB_TITLE, styles["Title"]))
    story.append(Paragraph(f"Periodo: {start.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}", styles["Normal"]))
    story.append(Paragraph(f"Generado automáticamente · {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 0.4*cm))

    # KPIs
    story.append(Paragraph("Resumen", styles["Heading2"]))
    kpis = _kpi_block(df)
    kpi_table = [["KPI", "Valor"]] + [[k, str(v)] for k, v in kpis.items()]
    t = Table(kpi_table, colWidths=[9.5*cm, 6.0*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.5*cm))

    # Charts
    story.append(Paragraph("Gráficos", styles["Heading2"]))
    for img in [c_tipo, c_zona, c_mec, c_sev, c_mes]:
        story.append(Image(img, width=17*cm, height=8*cm))
        story.append(Spacer(1, 0.25*cm))

    # TOPs
    story.append(PageBreak())
    story.append(Paragraph(f"TOP {TOP_N}", styles["Heading2"]))

    blocks = [
        ("Localización anatómica", "Localización anatómica"),
        ("Tipo de lesión", "Tipo de lesión"),
        ("Diagnóstico", "Diagnóstico"),
        ("Categoría", "Categoría"),
        ("Posición", "Posición"),
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
