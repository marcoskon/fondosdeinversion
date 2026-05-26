#!/usr/bin/env python3
"""
CAFCI Dashboard — Procesador de Planilla Diaria
================================================
Uso:
    python process_excel.py <ruta_al_excel>
    python process_excel.py "20260522_Planilla_Diaria_A.xlsx"

Genera:  ../data/cafci_data.json
         (o data/cafci_data.json si se ejecuta desde la raíz del repo)
"""

import sys
import json
import os
import re
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas no está instalado. Ejecutá: pip install pandas openpyxl")
    sys.exit(1)

def get_base_name(name: str) -> str:
    """Elimina sufijos de clase ('- Clase A', '- Clase B', etc.)"""
    return re.sub(r'\s*[-–]\s*Clase\b[\s\S]*$', '', name, flags=re.IGNORECASE).strip()

def annualize(pct, days):
    """Convierte variación acumulada en TNA"""
    if pct is None:
        return None
    try:
        return round(((1 + pct / 100) ** (365 / days) - 1) * 100, 2)
    except Exception:
        return None

def process(excel_path: str) -> list:
    df = pd.read_excel(excel_path, header=None)
    data = df.iloc[10:].copy()
    data.columns = range(len(data.columns))

    # Mapear categoría a cada fila
    category_map = {}
    current_cat = 'Sin Clasificar'
    for idx, row in data.iterrows():
        val0 = row[0]
        val1 = row[1]
        is_section = (
            (val1 is None or (hasattr(val1, '__float__') and pd.isna(val1)))
            and pd.notna(val0)
            and str(val0).strip()
            and not str(val0).strip().startswith('(')
            and not str(val0).strip().startswith('Advertencia')
            and len(str(val0).strip()) > 5
        )
        if is_section:
            current_cat = str(val0).strip()
        elif pd.notna(val1):
            category_map[idx] = current_cat

    funds_df = data[data[1].notna()].copy()
    funds_df['categoria'] = funds_df.index.map(category_map)

    # Categorías a ignorar
    skip = {
        'En Proceso de Liquidacion por Pago Parcial y Especies',
        'En Proceso de Liquidacion por Pago Total',
        'Solicitud en tramite',
        'Clases en Dolar Estadounidense',
    }

    records = []
    for _, row in funds_df.iterrows():
        cat = row.get('categoria', 'Sin Clasificar')
        if cat in skip:
            continue

        def safe_float(v):
            try:
                f = float(v)
                return None if pd.isna(f) else round(f, 6)
            except Exception:
                return None

        def safe_str(v):
            return str(v).strip() if pd.notna(v) else None

        pat = safe_float(row[14])
        if not pat or pat <= 0:
            continue

        vy  = safe_float(row[10])   # YTD desde 30/12
        vm  = safe_float(row[9])    # MTD desde 30/04
        vd  = safe_float(row[7])    # Diaria
        v12 = safe_float(row[11])   # 12M

        records.append({
            'f':   str(row[0]).strip(),
            'cat': cat,
            'mon': safe_str(row[1]),
            'hor': safe_str(row[3]),
            'pat': round(pat / 1e6, 4),   # millones ARS
            'ms':  safe_float(row[16]),
            'vd':  vd,
            'vm':  vm,
            'vy':  vy,
            'v12': v12,
            'tm':  annualize(vm,  22),    # TNA MTD (22 días aprox)
            'ty':  annualize(vy, 143),    # TNA YTD (143 días aprox)
            'sg':  safe_str(row[23]),
        })

    return records

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    excel_path = sys.argv[1]
    if not os.path.exists(excel_path):
        print(f"ERROR: No se encontró el archivo '{excel_path}'")
        sys.exit(1)

    print(f"Procesando: {excel_path}")
    records = process(excel_path)
    print(f"Fondos activos encontrados: {len(records)}")

    # Detectar fecha del archivo desde el nombre o usar hoy
    date_match = re.search(r'(\d{8})', os.path.basename(excel_path))
    if date_match:
        raw = date_match.group(1)
        file_date = f"{raw[6:8]}/{raw[4:6]}/{raw[0:4]}"
    else:
        file_date = datetime.today().strftime('%d/%m/%Y')

    output = {
        'fecha': file_date,
        'generado': datetime.now().isoformat(timespec='seconds'),
        'total_fondos': len(records),
        'fondos': records,
    }

    # Guardar en data/cafci_data.json relativo al repo
    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(script_dir, '..', 'data')
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, 'cafci_data.json')

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, separators=(',', ':'))

    size_kb = os.path.getsize(out_path) / 1024
    print(f"Guardado en: {out_path}  ({size_kb:.0f} KB)")
    print(f"Fecha de datos: {file_date}")

if __name__ == '__main__':
    main()
