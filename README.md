# 📊 CAFCI Dashboard

Dashboard interactivo de Fondos Comunes de Inversión argentinos, con datos de [CAFCI](https://www.cafci.org.ar/).

🌐 **Ver dashboard:** `https://<tu-usuario>.github.io/<nombre-del-repo>/`

---

## 🗂 Estructura del repositorio

```
cafci-dashboard/
├── index.html              ← Dashboard (no tocar)
├── data/
│   └── cafci_data.json     ← Datos diarios (se actualiza a diario)
├── scripts/
│   └── process_excel.py    ← Script para procesar la planilla CAFCI
└── .github/
    └── workflows/
        └── pages.yml       ← Auto-deploy a GitHub Pages
```

---

## 🚀 Configuración inicial (una sola vez)

### 1. Crear el repositorio en GitHub
```bash
git init cafci-dashboard
cd cafci-dashboard
# Copiar todos los archivos acá
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/<tu-usuario>/cafci-dashboard.git
git push -u origin main
```

### 2. Activar GitHub Pages
- Ir a **Settings → Pages**
- En *Source*, seleccionar **GitHub Actions**
- Guardar

El dashboard queda disponible en `https://<tu-usuario>.github.io/cafci-dashboard/`

### 3. Instalar dependencias Python (solo una vez)
```bash
pip install pandas openpyxl
```

---

## 📅 Actualización diaria

Cada día que CAFCI publique la planilla:

### Opción A — Línea de comandos
```bash
# 1. Descargar la planilla de https://www.cafci.org.ar/
# 2. Procesar el Excel
python scripts/process_excel.py ruta/al/archivo/20260523_Planilla_Diaria_A.xlsx

# 3. Subir a GitHub
git add data/cafci_data.json
git commit -m "Datos al 23/05/2026"
git push
```

### Opción B — GitHub Desktop
1. Correr el script Python localmente
2. Abrir GitHub Desktop → los cambios en `data/cafci_data.json` aparecen automáticamente
3. Escribir el mensaje del commit (ej: "Datos 23/05/2026") y hacer **Commit → Push**

El deploy a GitHub Pages se dispara solo al hacer push (tarda ~30 segundos).

---

## 📌 Notas

- El archivo `data/cafci_data.json` pesa ~540 KB — dentro del límite de GitHub (100 MB por archivo).
- GitHub Pages tiene un límite de 1 GB de almacenamiento y 100 GB de bandwidth/mes (más que suficiente).
- El dashboard funciona en cualquier navegador moderno. No necesita servidor propio.
- Para verlo localmente sin subirlo a GitHub, usar un servidor HTTP simple:
  ```bash
  python -m http.server 8080
  # Abrir http://localhost:8080
  ```
  ⚠️ No abrir `index.html` directamente con doble clic — el fetch no funciona con `file://`.

---

## 🔧 Personalización

### Cambiar las fechas de los períodos
En `scripts/process_excel.py`, ajustar los días de cada período:
```python
'tm': annualize(vm, 22),   # días del MTD
'ty': annualize(vy, 143),  # días del YTD
```

### Agregar nuevas categorías a ignorar
En `scripts/process_excel.py`, agregar a la lista `skip`:
```python
skip = {
    'En Proceso de Liquidacion por Pago Parcial y Especies',
    ...
    'Nueva Categoria a Ignorar',
}
```
