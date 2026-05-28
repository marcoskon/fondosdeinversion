#!/usr/bin/env python3
"""
CAFCI Dashboard — Procesador de Planilla Diaria
================================================
Uso:
    python process_excel.py <ruta_al_excel>

Genera:
    data/cafci_data.json    — métricas del día (reemplaza)
    data/cafci_history.json — serie histórica VCP (acumula día a día)
"""

import sys, json, os, re
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("ERROR: pip install pandas openpyxl")
    sys.exit(1)

# ── Helpers ──────────────────────────────────────────────────────────────────

def get_base_name(name: str) -> str:
    return re.sub(r'\s*[-–]\s*Clase\b[\s\S]*$', '', name, flags=re.IGNORECASE).strip()

def annualize(pct, days):
    if pct is None: return None
    try: return round(((1 + pct/100)**(365/days) - 1)*100, 2)
    except: return None

def safe_float(v):
    try:
        f = float(v)
        return None if pd.isna(f) else round(f, 6)
    except: return None

def safe_str(v):
    return str(v).strip() if pd.notna(v) else None

# ── Procesar Excel ────────────────────────────────────────────────────────────

def process(excel_path: str):
    df = pd.read_excel(excel_path, header=None)
    data = df.iloc[10:].copy()
    data.columns = range(len(data.columns))

    # Mapear categoría
    category_map = {}
    current_cat = 'Sin Clasificar'
    for idx, row in data.iterrows():
        val0, val1 = row[0], row[1]
        is_section = (
            (val1 is None or (isinstance(val1, float) and pd.isna(val1)))
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

    skip = {
        'En Proceso de Liquidacion por Pago Parcial y Especies',
        'En Proceso de Liquidacion por Pago Total',
        'Solicitud en tramite',
        'Clases en Dolar Estadounidense',
    }

    records = []
    for _, row in data[data[1].notna()].iterrows():
        cat = category_map.get(_, 'Sin Clasificar')
        if cat in skip: continue

        pat = safe_float(row[14])
        if not pat or pat <= 0: continue

        vcp  = safe_float(row[5])   # VCP actual (cuotaparte del día)
        vy   = safe_float(row[10])
        vm   = safe_float(row[9])
        vd   = safe_float(row[7])
        v12  = safe_float(row[11])

        records.append({
            'f':   str(row[0]).strip(),
            'cat': cat,
            'mon': safe_str(row[1]),
            'hor': safe_str(row[3]),
            'pat': round(pat / 1e6, 4),
            'ms':  safe_float(row[16]),
            'vcp': vcp,                  # VCP del día
            'vd':  vd,
            'vm':  vm,
            'vy':  vy,
            'v12': v12,
            'tm':  None,  # calculado en main() con días reales
            'ty':  None,  # calculado en main() con días reales
            'sg':  safe_str(row[23]),
        })

    return records

# ── Actualizar historial de VCP ────────────────────────────────────────────────

def update_history(records, date_str: str, history_path: str):
    """
    Agrega los VCP del día al historial acumulado.
    Para fondos con múltiples clases, guarda el VCP de la clase con mayor patrimonio.
    Estructura: { "base_name||cat": { "dates": [...], "vcps": [...] } }
    """
    # Cargar historial existente
    if os.path.exists(history_path):
        with open(history_path, encoding='utf-8') as f:
            history = json.load(f)
        funds_hist = history.get('funds', {})
    else:
        funds_hist = {}

    # Agrupar por fondo base: quedarse con la clase de mayor patrimonio
    best_class = {}  # key -> {vcp, pat}
    for r in records:
        if not r['vcp'] or not r['pat']: continue
        key = get_base_name(r['f']) + '||' + r['cat']
        if key not in best_class or r['pat'] > best_class[key]['pat']:
            best_class[key] = {'vcp': r['vcp'], 'pat': r['pat']}

    # Agregar fecha de hoy a cada fondo
    updated = 0
    for key, info in best_class.items():
        vcp = round(info['vcp'], 4)
        if key not in funds_hist:
            funds_hist[key] = {'dates': [], 'vcps': []}

        fund_data = funds_hist[key]

        # Evitar duplicar si ya existe esta fecha
        if date_str in fund_data['dates']:
            idx = fund_data['dates'].index(date_str)
            fund_data['vcps'][idx] = vcp
        else:
            fund_data['dates'].append(date_str)
            fund_data['vcps'].append(vcp)
            updated += 1

    # Mantener solo últimos 500 días hábiles (~2 años)
    MAX_DAYS = 500
    for key in funds_hist:
        d = funds_hist[key]
        if len(d['dates']) > MAX_DAYS:
            d['dates'] = d['dates'][-MAX_DAYS:]
            d['vcps']  = d['vcps'][-MAX_DAYS:]

    history_out = {
        'updated':     date_str,
        'total_funds': len(funds_hist),
        'funds':       funds_hist,
    }

    with open(history_path, 'w', encoding='utf-8') as f:
        json.dump(history_out, f, ensure_ascii=False, separators=(',', ':'))

    size_kb = os.path.getsize(history_path) / 1024
    print(f"Historial actualizado: {len(funds_hist)} fondos, {updated} entradas nuevas ({size_kb:.0f} KB)")
    return history_out

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    excel_path = sys.argv[1]
    if not os.path.exists(excel_path):
        print(f"ERROR: No se encontró '{excel_path}'")
        sys.exit(1)

    print(f"Procesando: {excel_path}")
    records = process(excel_path)
    print(f"Fondos activos: {len(records)}")

    # Detectar fecha
    date_match = re.search(r'(\d{8})', os.path.basename(excel_path))
    if date_match:
        raw = date_match.group(1)
        date_str  = raw
        file_date = f"{raw[6:8]}/{raw[4:6]}/{raw[0:4]}"
        file_date_obj = datetime(int(raw[0:4]), int(raw[4:6]), int(raw[6:8])).date()
    else:
        today         = datetime.today()
        date_str      = today.strftime('%Y%m%d')
        file_date     = today.strftime('%d/%m/%Y')
        file_date_obj = today.date()

    # Calcular días reales para anualizacion
    import calendar as _cal
    from datetime import date as _date
    ytd_start    = _date(file_date_obj.year - 1, 12, 31)
    days_ytd     = (file_date_obj - ytd_start).days

    prev_month      = file_date_obj.month - 1 if file_date_obj.month > 1 else 12
    prev_month_year = file_date_obj.year if file_date_obj.month > 1 else file_date_obj.year - 1
    last_day_prev   = _cal.monthrange(prev_month_year, prev_month)[1]
    mtd_start       = _date(prev_month_year, prev_month, last_day_prev)
    days_mtd        = (file_date_obj - mtd_start).days

    print(f"Días YTD: {days_ytd} (desde {ytd_start})")
    print(f"Días MTD: {days_mtd} (desde {mtd_start})")

    # Calcular TNA con días reales
    for r in records:
        r['tm'] = annualize(r['vm'], days_mtd)
        r['ty'] = annualize(r['vy'], days_ytd)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir    = os.path.join(script_dir, '..', 'data')
    os.makedirs(out_dir, exist_ok=True)

    # ── 1. Guardar cafci_data.json (métricas del día) ──────────────────────
    # Quitar campo vcp del output principal (solo va al historial)
    records_out = [{k: v for k, v in r.items() if k != 'vcp'} for r in records]

    data_out = {
        'fecha':        file_date,
        'generado':     datetime.now().isoformat(timespec='seconds'),
        'total_fondos': len(records_out),
        'fondos':       records_out,
    }
    data_path = os.path.join(out_dir, 'cafci_data.json')
    with open(data_path, 'w', encoding='utf-8') as f:
        json.dump(data_out, f, ensure_ascii=False, separators=(',', ':'))
    print(f"Datos del día: {data_path} ({os.path.getsize(data_path)/1024:.0f} KB)")

    # ── 2. Actualizar cafci_history.json (serie histórica) ─────────────────
    history_path = os.path.join(out_dir, 'cafci_history.json')
    update_history(records, date_str, history_path)

    print(f"Fecha: {file_date}")

if __name__ == '__main__':
    main()
