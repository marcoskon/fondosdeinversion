#!/usr/bin/env python3
"""
CAFCI Dashboard — Procesador de Planilla Diaria
Uso: python process_excel.py <ruta_al_excel>
Genera: data/cafci_data.json  y  data/cafci_history.json
"""
import sys, json, os, re, calendar
from datetime import datetime, date

try:
    import pandas as pd
except ImportError:
    print("ERROR: pip install pandas openpyxl"); sys.exit(1)

def get_base_name(name):
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

def parse_date(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return None
    try:
        s = str(val).strip()
        parts = s.replace('/', '-').split('-')
        d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
        if y < 100: y += 2000
        return date(y, m, d)
    except: return None

def process(excel_path):
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
        records.append({
            'f':   str(row[0]).strip(),
            'cat': cat,
            'mon': safe_str(row[1]),
            'hor': safe_str(row[3]),
            'pat': round(pat / 1e6, 4),
            'ms':  safe_float(row[16]),
            'vcp': safe_float(row[5]),
            'vd':  safe_float(row[7]),
            'vm':  safe_float(row[9]),
            'vy':  safe_float(row[10]),
            'v12': safe_float(row[11]),
            'tm':  None,
            'ty':  None,
            'sg':  safe_str(row[23]),
        })

    # Leer fecha VCP y referencias del header
    vcp_date = parse_date(data[data[1].notna()].iloc[0][4]) if len(data[data[1].notna()]) > 0 else None
    header   = df.iloc[8]
    ref_mtd  = parse_date(header[9])
    ref_ytd  = parse_date(header[10])
    ref_12m  = parse_date(header[11])

    return records, vcp_date, ref_mtd, ref_ytd, ref_12m

def update_history(records, date_str, history_path):
    if os.path.exists(history_path):
        with open(history_path, encoding='utf-8') as f:
            history = json.load(f)
        funds_hist = history.get('funds', {})
    else:
        funds_hist = {}

    best_class = {}
    for r in records:
        if not r.get('vcp') or not r['pat']: continue
        key = get_base_name(r['f']) + '||' + r['cat']
        if key not in best_class or r['pat'] > best_class[key]['pat']:
            best_class[key] = {'vcp': r['vcp'], 'pat': r['pat']}

    updated = 0
    for key, info in best_class.items():
        vcp = round(info['vcp'], 4)
        if key not in funds_hist:
            funds_hist[key] = {'dates': [], 'vcps': []}
        fd = funds_hist[key]
        if date_str in fd['dates']:
            fd['vcps'][fd['dates'].index(date_str)] = vcp
        else:
            fd['dates'].append(date_str)
            fd['vcps'].append(vcp)
            updated += 1

    MAX_DAYS = 500
    for key in funds_hist:
        d = funds_hist[key]
        if len(d['dates']) > MAX_DAYS:
            d['dates'] = d['dates'][-MAX_DAYS:]
            d['vcps']  = d['vcps'][-MAX_DAYS:]

    out = {'updated': date_str, 'total_funds': len(funds_hist), 'funds': funds_hist}
    with open(history_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    print(f"Historial: {len(funds_hist)} fondos, {updated} entradas nuevas ({os.path.getsize(history_path)//1024} KB)")

def main():
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)

    excel_path = sys.argv[1]
    if not os.path.exists(excel_path):
        print(f"ERROR: No se encontró '{excel_path}'"); sys.exit(1)

    print(f"Procesando: {excel_path}")
    records, vcp_date, ref_mtd, ref_ytd, ref_12m = process(excel_path)
    print(f"Fondos activos: {len(records)}")

    # Fecha desde el nombre del archivo
    date_match = re.search(r'(\d{8})', os.path.basename(excel_path))
    if date_match:
        raw = date_match.group(1)
        file_date     = f"{raw[6:8]}/{raw[4:6]}/{raw[0:4]}"
        file_date_obj = date(int(raw[0:4]), int(raw[4:6]), int(raw[6:8]))
        date_str      = raw
    else:
        today         = datetime.today()
        file_date     = today.strftime('%d/%m/%Y')
        file_date_obj = today.date()
        date_str      = today.strftime('%Y%m%d')

    # Usar fecha VCP del Excel (más precisa que el nombre del archivo)
    base_date = vcp_date or file_date_obj
    file_date = f"{base_date.day:02d}/{base_date.month:02d}/{base_date.year}"
    date_str  = base_date.strftime('%Y%m%d')

    # Fallback referencias
    if not ref_ytd:
        ref_ytd = date(base_date.year - 1, 12, 31)
    if not ref_mtd:
        pm = base_date.month - 1 if base_date.month > 1 else 12
        py = base_date.year if base_date.month > 1 else base_date.year - 1
        ref_mtd = date(py, pm, calendar.monthrange(py, pm)[1])

    days_ytd = (base_date - ref_ytd).days
    days_mtd = (base_date - ref_mtd).days
    days_12m = (base_date - ref_12m).days if ref_12m else 365

    print(f"Fecha VCP: {base_date} | YTD: {days_ytd} días | MTD: {days_mtd} días")

    # Calcular TNA y quitar VCP del output principal
    records_out = []
    for r in records:
        r['tm'] = annualize(r['vm'], days_mtd)
        r['ty'] = annualize(r['vy'], days_ytd)
        records_out.append({k: v for k, v in r.items() if k != 'vcp'})

    # Salidas
    script_dir   = os.path.dirname(os.path.abspath(__file__))
    out_dir      = os.path.join(script_dir, '..', 'data')
    os.makedirs(out_dir, exist_ok=True)

    data_out = {
        'fecha':        file_date,
        'generado':     datetime.now().isoformat(timespec='seconds'),
        'total_fondos': len(records_out),
        'fondos':       records_out,
    }
    data_path = os.path.join(out_dir, 'cafci_data.json')
    with open(data_path, 'w', encoding='utf-8') as f:
        json.dump(data_out, f, ensure_ascii=False, separators=(',', ':'))
    print(f"Datos del día: {data_path} ({os.path.getsize(data_path)//1024} KB)")

    history_path = os.path.join(out_dir, 'cafci_history.json')
    update_history(records, date_str, history_path)
    print(f"Fecha: {file_date}")

if __name__ == '__main__':
    main()
