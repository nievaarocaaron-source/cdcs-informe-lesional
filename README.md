# CDCS – Informe Lesional automático (PDF)

Cada vez que subas/actualices el Excel en `data/REGISTRO LESIONAL.xlsx`, GitHub Actions generará un PDF listo para entregar al club.

## Pasos rápidos
1) Sube tu Excel a `data/REGISTRO LESIONAL.xlsx`.
2) Sube el logo a `assets/logo_cdcs.png`.
3) Ve a **Actions** → último run → **Artifacts** → descarga `informe-lesional`.

## Ajustes
Edita `scripts/make_report.py`:
- `CLUB_TITLE`
- `FIRST_HALF_START` / `FIRST_HALF_END`
- `TOP_N`
