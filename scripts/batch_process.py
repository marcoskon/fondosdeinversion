#!/usr/bin/env python3
"""
CAFCI Dashboard — Procesamiento masivo de planillas históricas
==============================================================
Uso:
    python batch_process.py <carpeta_con_excels>

Ejemplo:
    python batch_process.py C:/Users/Marcos/Downloads/planillas/

Procesa todos los .xlsx de la carpeta en orden cronológico y genera:
    data/cafci_data.json         — datos del último día
    data/cafci_history.json      — historial de VCPs
    data/cafci_aum_history.json  — historial de patrimonios

Después subís solo esos 3 archivos a GitHub.
"""

import sys, os, glob, re
from datetime import datetime

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    folder = sys.argv[1].rstrip('/\\')
    if not os.path.isdir(folder):
        print(f"ERROR: '{folder}' no es una carpeta válida")
        sys.exit(1)

    # Buscar todos los .xlsx
    files = glob.glob(os.path.join(folder, '*.xlsx'))
    if not files:
        print(f"ERROR: No se encontraron archivos .xlsx en '{folder}'")
        sys.exit(1)

    # Ordenar por fecha del nombre (YYYYMMDD_...)
    def get_date(path):
        m = re.search(r'(\d{8})', os.path.basename(path))
        return m.group(1) if m else '00000000'

    files_sorted = sorted(files, key=get_date)

    print(f"\n{'='*60}")
    print(f"  CAFCI Batch Processor")
    print(f"{'='*60}")
    print(f"  Carpeta: {folder}")
    print(f"  Archivos encontrados: {len(files_sorted)}")
    print(f"  Rango: {get_date(files_sorted[0])} → {get_date(files_sorted[-1])}")
    print(f"{'='*60}\n")

    # Importar el módulo process_excel
    script_dir = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, script_dir)
    import process_excel as pe

    out_dir = os.path.join(script_dir, '..', 'data')
    os.makedirs(out_dir, exist_ok=True)

    errors = []
    processed = 0

    for i, excel_path in enumerate(files_sorted, 1):
        fname = os.path.basename(excel_path)
        date_str_raw = get_date(excel_path)

        print(f"[{i:3d}/{len(files_sorted)}] {fname}", end=' ... ', flush=True)

        try:
            # Procesar el Excel
            records, vcp_date, ref_mtd, ref_ytd, ref_12m = pe.process(excel_path)

            if not records:
                print("⚠ Sin fondos activos")
                continue

            # Fechas
            from datetime import date, datetime as dt2
            import calendar
            date_match = re.search(r'(\d{8})', fname)
            if date_match:
                raw = date_match.group(1)
                file_date_obj = date(int(raw[0:4]), int(raw[4:6]), int(raw[6:8]))
            else:
                file_date_obj = dt2.today().date()

            base_date = vcp_date or file_date_obj
            file_date = f"{file_date_obj.day:02d}/{file_date_obj.month:02d}/{file_date_obj.year}"
            hist_date_str = base_date.strftime('%Y%m%d')

            if not ref_ytd:
                ref_ytd = date(base_date.year - 1, 12, 31)
            if not ref_mtd:
                pm = base_date.month - 1 if base_date.month > 1 else 12
                py = base_date.year if base_date.month > 1 else base_date.year - 1
                ref_mtd = date(py, pm, calendar.monthrange(py, pm)[1])

            days_ytd = (base_date - ref_ytd).days
            days_mtd = (base_date - ref_mtd).days

            # Calcular TNA
            for r in records:
                r['tm'] = pe.annualize(r['vm'], days_mtd)
                r['ty'] = pe.annualize(r['vy'], days_ytd)

            records_out = [{k: v for k, v in r.items() if k != 'vcp'} for r in records]

            # Solo guardar cafci_data.json del último archivo
            if i == len(files_sorted):
                data_out = {
                    'fecha': file_date,
                    'generado': dt2.now().isoformat(timespec='seconds'),
                    'total_fondos': len(records_out),
                    'fondos': records_out,
                }
                import json
                data_path = os.path.join(out_dir, 'cafci_data.json')
                with open(data_path, 'w', encoding='utf-8') as f:
                    json.dump(data_out, f, ensure_ascii=False, separators=(',', ':'))

            # Actualizar historiales
            pe.update_history(records, hist_date_str, os.path.join(out_dir, 'cafci_history.json'))
            pe.update_aum_history(records, hist_date_str, os.path.join(out_dir, 'cafci_aum_history.json'))

            processed += 1
            print(f"✓ {len(records)} fondos | {file_date}")

        except Exception as e:
            errors.append((fname, str(e)))
            print(f"✗ ERROR: {e}")
            continue

    # Resumen final
    print(f"\n{'='*60}")
    print(f"  COMPLETADO: {processed}/{len(files_sorted)} archivos procesados")
    if errors:
        print(f"  ERRORES ({len(errors)}):")
        for fname, err in errors:
            print(f"    - {fname}: {err}")
    print(f"\n  Archivos generados en /data/:")
    for fname in ['cafci_data.json', 'cafci_history.json', 'cafci_aum_history.json']:
        path = os.path.join(out_dir, fname)
        if os.path.exists(path):
            size = os.path.getsize(path) / 1024
            print(f"    ✓ {fname} ({size:.0f} KB)")
    print(f"\n  Próximo paso: subí los 3 archivos de /data/ a tu repo de GitHub.")
    print(f"{'='*60}\n")

if __name__ == '__main__':
    main()
