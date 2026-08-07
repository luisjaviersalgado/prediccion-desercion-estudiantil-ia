# Predicción temprana del riesgo de deserción estudiantil universitaria
Actividad: Creación de un Modelo de Inteligencia Artificial Avanzada — Estrategia ABR

## Contenido
- `main.py` — Pipeline completo: genera el dataset sintético, entrena los 3 modelos
  (Regresión Logística, Random Forest, Red Neuronal MLP), evalúa métricas de desempeño
  y calcula el análisis de equidad (fairness) por género y estrato socioeconómico.
  Genera todas las figuras usadas en el informe y la presentación.
- `generar_informe.py` — Construye el informe Word (10-15 páginas) a partir de los
  resultados de `main.py` (requiere haber ejecutado main.py antes, usa `figures/` y `data/`).
- `generar_pptx.js` — Construye la presentación PowerPoint (10 diapositivas) con Node.js
  y pptxgenjs (requiere haber ejecutado main.py antes).
- `dataset_desercion.csv` — Dataset sintético generado (2400 estudiantes).
- `tabla_metricas.csv` — Tabla de métricas de desempeño de los 3 modelos.
- `resumen_resultados.json` — Resumen completo de resultados y métricas de equidad.

## Cómo reproducir
```bash
pip install scikit-learn pandas matplotlib numpy --break-system-packages
python3 main.py                 # genera data/ y figures/
python3 generar_informe.py      # genera outputs/Informe_...docx (requiere python-docx)
node generar_pptx.js            # genera outputs/Presentacion_...pptx (requiere pptxgenjs)
```

## Resumen de resultados
- Dataset: 2400 estudiantes sintéticos, tasa de deserción 42.3%
- Mejor AUC-ROC: Regresión Logística (0.677), seguida de Random Forest (0.674) y MLP (0.658)
- Mejor sensibilidad (recall): Red Neuronal MLP (88.7%), con umbral de decisión ajustado a 0.25
- Análisis de equidad (MLP): Disparate Impact Ratio de 0.91 (género) y 0.88 (estrato);
  Equal Opportunity Difference baja en ambos (~0.04); Equalized Odds Difference más alta
  por mayor tasa de falsos positivos en estrato bajo y género femenino.

Autor: Luis Javier Salgado Guzmán — Ingeniería de Software, séptimo semestre
