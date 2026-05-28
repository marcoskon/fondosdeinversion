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
    """TNA sin capitalización: (rendimiento% / días) × 365"""
    if pct is None or not days: return None
    try: return round((pct / days) * 365, 2)
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
            't12': None,  # TNA 12M, calculado en main()
            'sg':  safe_str(row[23]),
        })

    # Extraer la fecha real del VCP del primer registro válido (col 4 = fecha hábil)
    vcp_date_str = None
    for _, row in data[data[1].notna()].iterrows():
        raw_fecha = safe_str(row[4])
        if raw_fecha and raw_fecha != 'None':
            vcp_date_str = raw_fecha  # formato DD/MM/YY o DD/MM/YYYY
            break

    return records, vcp_date_str

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
    records, vcp_date_raw = process(excel_path)
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

    # Leer fechas de referencia directamente del header del Excel
    # Fila 8 (índice 8) contiene las fechas exactas que usa CAFCI
    # Col 9 = MTD (ej: 30/04/26), Col 10 = YTD (ej: 30/12/25), Col 11 = 12M
    from datetime import date as _date
    def parse_cafci_date(val):
        """Parsea fecha DD/MM/YY del header del Excel"""
        try:
            s = str(val).strip()
            d, m, y = s.split('/')
            year = 2000 + int(y) if len(y) == 2 else int(y)
            return _date(year, int(m), int(d))
        except:
            return None

    # ── Leer fechas de referencia desde el propio Excel ──────────────────────
    # Row 8 tiene las fechas base de comparación (MTD, YTD, 12M)
    # Col 4 de los fondos tiene la fecha real del VCP (día hábil anterior)
    df_main = pd.read_excel(excel_path, header=None)
    header_row = df_main.iloc[8]

    def parse_cafci_date(val):
        if val is None or (isinstance(val, float) and pd.isna(val)): return None
        try:
            s = str(val).strip()
            parts = s.replace('/', '-').split('-')
            d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 100: y += 2000
            from datetime import date as _date
            return _date(y, m, d)
        except: return None

    # Fecha del VCP: columna 4 del primer registro de fondo
    vcp_date = None
    for _, row in df_main.iloc[10:].iterrows():
        if pd.notna(row[1]):
            vcp_date = parse_cafci_date(row[4])
            break

    # Fechas base de comparación del header
    ref_mtd = parse_cafci_date(header_row[9])   # último día hábil mes anterior
    ref_ytd = parse_cafci_date(header_row[10])  # último día hábil dic año anterior
    ref_12m = parse_cafci_date(header_row[11])  # mismo período año anterior

    # Fallback: calcular desde la fecha del archivo
    import calendar as _cal
    from datetime import date as _date2
    base_date = vcp_date or file_date_obj

    if not ref_ytd:
        ref_ytd = _date2(base_date.year - 1, 12, 31)
    if not ref_mtd:
        pm = base_date.month - 1 if base_date.month > 1 else 12
        py = base_date.year if base_date.month > 1 else base_date.year - 1
        ref_mtd = _date2(py, pm, _cal.monthrange(py, pm)[1])

    # Días = desde fecha de referencia hasta fecha del VCP (día hábil anterior)
    days_ytd = (base_date - ref_ytd).days
    days_mtd = (base_date - ref_mtd).days
    days_12m = (base_date - ref_12m).days if ref_12m else 365

    print(f"Fecha VCP (día hábil anterior): {base_date}")
    print(f"Ref. YTD: {ref_ytd} → {days_ytd} días corridos")
    print(f"Ref. MTD: {ref_mtd} → {days_mtd} días corridos")
