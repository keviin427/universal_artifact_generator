# Universal Artifact Generator API

API desarrollada con **FastAPI** que permite generar y descargar archivos (Excel, Word, PDF, PowerPoint, SVG, CSV para Power BI) y ejecutar modelos de Machine Learning. Ideal para integrarse con GPTs personalizados, interfaces web y procesos de automatización documental.

---

## 🚀 Funcionalidades

- 📊 **Excel** - Tablas, fórmulas, cálculos automatizados
- 📝 **Word** - Documentos con secciones y encabezados
- 📽 **PowerPoint** - Presentaciones con viñetas
- 📄 **PDF** - Informes en formato portable
- 🖼 **Imagen SVG** - Branding o gráficos
- 🎨 **Canva SVG** - Plantillas creativas
- 📈 **Power BI CSV** - Dataset estructurado
- 🤖 **ML Entrenamiento** - Modelos supervisados
- 🔮 **ML Predicciones** - Clasificaciones

---

## ⚙️ Instalación

```bash
git clone https://github.com/tuusuario/universal-artifact-api.git
cd universal-artifact-api
pip install -r requirements_full.txt
```

---

## 🧪 Ejecutar localmente

```bash
uvicorn main:app --reload
```

Abrir Swagger UI en:
```
http://localhost:8000/docs
```

---

## 📤 Ejemplos de uso

### 🧾 1. Generar Excel

**Endpoint:** `/generate_excel`
```json
{
  "titulo": "Ventas Q3",
  "headers": ["Producto", "Cantidad", "Precio"],
  "rows": [["Zapatos", 20, 50], ["Camisas", 35, 30]],
  "formulas": {
    "D2": "=B2*C2",
    "D3": "=B3*C3"
  }
}
```

---

### 📄 2. Generar Word

**Endpoint:** `/generate_word`
```json
{
  "titulo": "Informe mensual",
  "secciones": [
    "Introducción",
    "Resumen Ejecutivo",
    "Resultados"
  ]
}
```

---

### 🎥 3. Generar PowerPoint

**Endpoint:** `/generate_ppt`
```json
{
  "titulo": "Estrategia 2025",
  "bullets": ["Objetivos", "Acciones", "KPIs"]
}
```

---

### 📈 4. Dataset Power BI

**Endpoint:** `/generate_powerbi`
```json
{
  "headers": ["Fecha", "Venta", "Región"],
  "rows": [["2023-09-01", 1200, "Norte"], ["2023-09-02", 800, "Sur"]]
}
```

---

### 🧠 5. ML Entrenamiento

**Endpoint:** `/train_model`
```json
{
  "features": [[5.1, 3.5, 1.4, 0.2], [7.0, 3.2, 4.7, 1.4]],
  "labels": [0, 1]
}
```

---

### 🔮 6. ML Predicción

**Endpoint:** `/predict_model`
```json
{
  "features": [[6.7, 3.1, 4.4, 1.4]]
}
```

---

## 🛠 Arquitectura híbrida y offloading

Si detectas en el input términos como `"merge"`, `"agrupado"`, `"dashboard"` o `"KPIs"` el GPT debe derivar a estos endpoints:

| Tarea Avanzada        | Endpoint sugerido         |
|-----------------------|---------------------------|
| Agrupar o unir Excel  | `/analyze_excel` (externo)|
| Ejecutar lógica Python| `/run_python_job`         |
| Construir dashboard   | `/build_dashboard`        |

---

## 📁 Estructura del Proyecto

```
📦 universal-artifact-api
├── main.py
├── openapi.json
├── requirements_full.txt
├── render.yaml
├── README.md
└── resultados/
```

---

## 🌐 Despliegue en Render

1. Crear nuevo servicio Web Service
2. Seleccionar repositorio GitHub
3. Configurar:
   - Start Command: `uvicorn main:app --host 0.0.0.0 --port 10000`
   - Python Version: 3.11
   - Build Command: `pip install -r requirements_full.txt`

---

## ✅ Probado con:

- GPT personalizados en ChatGPT Plus
- Swagger/OpenAPI 3.1
- Hosting en Render

---

