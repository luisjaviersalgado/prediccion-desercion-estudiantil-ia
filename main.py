# -*- coding: utf-8 -*-
"""
Proyecto: Predicción temprana del riesgo de deserción estudiantil universitaria
mediante una red neuronal (Perceptrón Multicapa).

Actividad: Creación de un modelo de inteligencia artificial avanzada
Estrategia: Aprendizaje Basado en Retos (ABR)

Este script:
  1. Genera un conjunto de datos sintético pero realista de estudiantes
     universitarios (dado que no se dispone de acceso a datasets externos
     en este entorno de ejecución offline).
  2. Preprocesa los datos (codificación, escalado, partición train/val/test).
  3. Entrena tres modelos: Regresión Logística (línea base), Random Forest
     (línea base de ensamble) y un Perceptrón Multicapa -MLP- (modelo
     avanzado / red neuronal profunda superficial).
  4. Evalúa el desempeño con métricas estándar (exactitud, precisión,
     sensibilidad, F1, AUC-ROC) y genera visualizaciones.
  5. Realiza un análisis de equidad (fairness) por género y por estrato
     socioeconómico, calculando Disparate Impact, Equal Opportunity
     Difference y Equalized Odds Difference.
  6. Exporta todas las figuras y una tabla de métricas para el informe.

Autor: Luis Javier Salgado Guzmán - Ingeniería de Software
"""

import json
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.neural_network import MLPClassifier
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    roc_auc_score, roc_curve, confusion_matrix, ConfusionMatrixDisplay
)

warnings.filterwarnings("ignore")
RNG = np.random.default_rng(42)

FIG_DIR = "figures"
DATA_DIR = "data"

plt.rcParams.update({
    "figure.dpi": 150,
    "font.size": 11,
    "axes.spines.top": False,
    "axes.spines.right": False,
})

COLOR_PRIMARY = "#2E5EAA"
COLOR_SECONDARY = "#D9822B"
COLOR_NEUTRAL = "#6B7280"
COLOR_OK = "#2E9E6B"
COLOR_BAD = "#C0392B"


# ---------------------------------------------------------------------------
# 1. GENERACIÓN DEL CONJUNTO DE DATOS SINTÉTICO
# ---------------------------------------------------------------------------
def generar_datos(n=2400):
    """Genera un dataset sintético de estudiantes universitarios con un
    riesgo de deserción generado a partir de una combinación realista de
    factores académicos, de asistencia y socioeconómicos, siguiendo los
    factores de riesgo reportados en la literatura (desempeño académico,
    asistencia/participación y variables socioeconómicas)."""

    genero = RNG.choice(["Femenino", "Masculino"], size=n, p=[0.52, 0.48])
    estrato = RNG.choice([1, 2, 3, 4, 5, 6], size=n,
                          p=[0.12, 0.23, 0.30, 0.20, 0.10, 0.05])
    edad = np.clip(RNG.normal(20.5, 2.2, n), 17, 35).round(0)
    semestre = RNG.integers(1, 11, n)

    # La probabilidad de trabajar decrece con el estrato (más común en
    # estratos bajos por razones económicas).
    prob_trabaja = np.clip(0.55 - 0.06 * (estrato - 1), 0.10, 0.6)
    trabaja = RNG.binomial(1, prob_trabaja)

    horas_estudio = np.clip(
        RNG.normal(14, 6, n) - trabaja * RNG.normal(4, 2, n), 0, 40
    )

    # El apoyo financiero (becas) es más frecuente en estratos bajos
    # (programas de fomento) y algo menos en estratos altos.
    prob_beca = np.clip(0.45 - 0.03 * (estrato - 1), 0.08, 0.5)
    apoyo_financiero = RNG.binomial(1, prob_beca)

    distancia_km = np.clip(RNG.exponential(9, n), 0, 60)

    participacion_extra = RNG.binomial(1, 0.30, n)

    # Promedio académico anterior (escala 0-5), correlacionado con horas de
    # estudio y con ruido individual.
    promedio_anterior = np.clip(
        2.6 + 0.045 * horas_estudio + RNG.normal(0, 0.5, n), 0.5, 5.0
    )

    asistencia_pct = np.clip(
        78 + 4 * (participacion_extra) - 3 * trabaja + RNG.normal(0, 10, n),
        30, 100
    )

    materias_reprobadas = np.clip(
        RNG.poisson(np.clip(2.4 - 0.35 * promedio_anterior, 0.05, None)), 0, 6
    )

    creditos_matriculados = RNG.integers(12, 22, n)

    # --- Modelo generador del riesgo real de deserción --------------------
    # Combinación lineal de factores de riesgo conocidos + un pequeño sesgo
    # estructural residual asociado al estrato socioeconómico (para poder
    # ilustrar, de forma controlada, un caso de inequidad algorítmica en la
    # Fase de evaluación de equidad).
    z = (
        -0.55 * (promedio_anterior - 3.0)
        - 0.028 * (asistencia_pct - 75)
        + 0.33 * materias_reprobadas
        + 0.55 * trabaja
        - 0.40 * apoyo_financiero
        + 0.020 * distancia_km
        - 0.30 * participacion_extra
        + 0.05 * (semestre <= 2).astype(float)      # riesgo de deserción temprana
        - 0.12 * (estrato - 3.5) * 0.0               # sin efecto directo deliberado
        + 0.18 * (estrato <= 2).astype(float)        # sesgo estructural residual
        - 0.85
        + RNG.normal(0, 0.65, n)
    )
    prob_desercion = 1 / (1 + np.exp(-z))
    desercion = RNG.binomial(1, prob_desercion)

    df = pd.DataFrame({
        "edad": edad,
        "genero": genero,
        "estrato": estrato,
        "semestre": semestre,
        "trabaja": trabaja,
        "horas_estudio_semanal": horas_estudio.round(1),
        "apoyo_financiero": apoyo_financiero,
        "distancia_km": distancia_km.round(1),
        "participacion_extracurricular": participacion_extra,
        "promedio_anterior": promedio_anterior.round(2),
        "asistencia_pct": asistencia_pct.round(1),
        "materias_reprobadas": materias_reprobadas,
        "creditos_matriculados": creditos_matriculados,
        "desercion": desercion,
    })
    return df


def grupo_estrato(e):
    if e <= 2:
        return "Bajo (1-2)"
    elif e <= 4:
        return "Medio (3-4)"
    else:
        return "Alto (5-6)"


# ---------------------------------------------------------------------------
# 2. PREPROCESAMIENTO
# ---------------------------------------------------------------------------
def construir_pipeline_preprocesamiento():
    num_cols = [
        "edad", "semestre", "trabaja", "horas_estudio_semanal",
        "apoyo_financiero", "distancia_km", "participacion_extracurricular",
        "promedio_anterior", "asistencia_pct", "materias_reprobadas",
        "creditos_matriculados", "estrato",
    ]
    cat_cols = ["genero"]

    pre = ColumnTransformer([
        ("num", StandardScaler(), num_cols),
        ("cat", OneHotEncoder(drop="if_binary"), cat_cols),
    ])
    return pre, num_cols + cat_cols


# ---------------------------------------------------------------------------
# 3. MÉTRICAS Y GRÁFICAS
# ---------------------------------------------------------------------------
def evaluar_modelo(nombre, y_true, y_pred, y_prob):
    return {
        "modelo": nombre,
        "exactitud": accuracy_score(y_true, y_pred),
        "precision": precision_score(y_true, y_pred),
        "sensibilidad_recall": recall_score(y_true, y_pred),
        "f1": f1_score(y_true, y_pred),
        "auc_roc": roc_auc_score(y_true, y_prob),
    }


def graficar_comparacion_modelos(tabla_metricas):
    df = pd.DataFrame(tabla_metricas).set_index("modelo")
    metrics = ["exactitud", "precision", "sensibilidad_recall", "f1", "auc_roc"]
    labels = ["Exactitud", "Precisión", "Sensibilidad", "F1", "AUC-ROC"]

    fig, ax = plt.subplots(figsize=(8, 4.5))
    x = np.arange(len(metrics))
    width = 0.25
    colors = [COLOR_NEUTRAL, COLOR_SECONDARY, COLOR_PRIMARY]
    for i, modelo in enumerate(df.index):
        vals = [df.loc[modelo, m] for m in metrics]
        ax.bar(x + (i - 1) * width, vals, width, label=modelo, color=colors[i % len(colors)])
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_ylim(0, 1.05)
    ax.set_ylabel("Valor de la métrica")
    ax.set_title("Comparación de desempeño entre modelos (conjunto de prueba)")
    ax.legend(loc="lower right", frameon=False)
    ax.yaxis.grid(True, alpha=0.3)
    ax.set_axisbelow(True)
    fig.tight_layout()
    fig.savefig(f"{FIG_DIR}/comparacion_modelos.png")
    plt.close(fig)


def graficar_matriz_confusion(y_true, y_pred, nombre_archivo, titulo):
    cm = confusion_matrix(y_true, y_pred)
    fig, ax = plt.subplots(figsize=(5.4, 4.6))
    disp = ConfusionMatrixDisplay(confusion_matrix=cm,
                                   display_labels=["No deserta", "Deserta"])
    disp.plot(ax=ax, cmap="Blues", colorbar=False, values_format="d")
    ax.set_xlabel("Clase predicha")
    ax.set_ylabel("Clase real")
    ax.set_title(titulo, fontsize=12)
    fig.tight_layout()
    fig.savefig(f"{FIG_DIR}/{nombre_archivo}.png")
    plt.close(fig)


def graficar_roc(modelos_probs, y_true):
    fig, ax = plt.subplots(figsize=(5.2, 4.6))
    colors = {"Regresión Logística": COLOR_NEUTRAL,
              "Random Forest": COLOR_SECONDARY,
              "Red Neuronal (MLP)": COLOR_PRIMARY}
    for nombre, prob in modelos_probs.items():
        fpr, tpr, _ = roc_curve(y_true, prob)
        auc = roc_auc_score(y_true, prob)
        ax.plot(fpr, tpr, label=f"{nombre} (AUC={auc:.3f})", color=colors.get(nombre))
    ax.plot([0, 1], [0, 1], linestyle="--", color="lightgray")
    ax.set_xlabel("Tasa de falsos positivos")
    ax.set_ylabel("Tasa de verdaderos positivos")
    ax.set_title("Curvas ROC — comparación de modelos")
    ax.legend(loc="lower right", frameon=False, fontsize=9)
    fig.tight_layout()
    fig.savefig(f"{FIG_DIR}/curvas_roc.png")
    plt.close(fig)


def graficar_curva_perdida(mlp):
    fig, ax = plt.subplots(figsize=(6.5, 4.3))
    ax.plot(mlp.loss_curve_, color=COLOR_PRIMARY, linewidth=2)
    ax.set_xlabel("Iteración de entrenamiento")
    ax.set_ylabel("Pérdida (log-loss)")
    ax.set_title("Curva de pérdida durante el entrenamiento\nRed Neuronal (MLP)", fontsize=12)
    ax.yaxis.grid(True, alpha=0.3)
    ax.set_axisbelow(True)
    fig.tight_layout()
    fig.savefig(f"{FIG_DIR}/curva_perdida_mlp.png")
    plt.close(fig)


def graficar_importancia_variables(rf, feature_names):
    importancias = pd.Series(rf.feature_importances_, index=feature_names)
    importancias = importancias.sort_values(ascending=True).tail(10)
    fig, ax = plt.subplots(figsize=(6.5, 5))
    ax.barh(importancias.index, importancias.values, color=COLOR_PRIMARY)
    ax.set_xlabel("Importancia relativa")
    ax.set_title("Variables más influyentes en la predicción\n(Random Forest)")
    fig.tight_layout()
    fig.savefig(f"{FIG_DIR}/importancia_variables.png")
    plt.close(fig)


# ---------------------------------------------------------------------------
# 4. ANÁLISIS DE EQUIDAD (FAIRNESS)
# ---------------------------------------------------------------------------
def metricas_equidad(y_true, y_pred, grupo):
    """Calcula, por cada valor del atributo sensible `grupo`:
       - tasa de selección (P(pred=1))
       - TPR (sensibilidad / equal opportunity)
       - FPR
    Y a partir de ellas: Disparate Impact Ratio, Equal Opportunity
    Difference y Equalized Odds Difference (máx. de |ΔTPR|, |ΔFPR|)."""
    df = pd.DataFrame({"y": y_true, "pred": y_pred, "grupo": grupo})
    filas = []
    for g, sub in df.groupby("grupo"):
        tp = ((sub.pred == 1) & (sub.y == 1)).sum()
        fp = ((sub.pred == 1) & (sub.y == 0)).sum()
        fn = ((sub.pred == 0) & (sub.y == 1)).sum()
        tn = ((sub.pred == 0) & (sub.y == 0)).sum()
        tasa_seleccion = (tp + fp) / len(sub)
        tpr = tp / (tp + fn) if (tp + fn) > 0 else np.nan
        fpr = fp / (fp + tn) if (fp + tn) > 0 else np.nan
        filas.append({"grupo": g, "n": len(sub), "tasa_seleccion": tasa_seleccion,
                       "tpr": tpr, "fpr": fpr})
    tabla = pd.DataFrame(filas)

    di_ratio = tabla["tasa_seleccion"].min() / tabla["tasa_seleccion"].max()
    eod_opportunity = tabla["tpr"].max() - tabla["tpr"].min()
    eod_odds = max(tabla["tpr"].max() - tabla["tpr"].min(),
                    tabla["fpr"].max() - tabla["fpr"].min())

    resumen = {
        "disparate_impact_ratio": di_ratio,
        "equal_opportunity_difference": eod_opportunity,
        "equalized_odds_difference": eod_odds,
    }
    return tabla, resumen


def graficar_equidad(tabla, titulo, nombre_archivo):
    fig, axes = plt.subplots(1, 2, figsize=(9.5, 4))
    axes[0].bar(tabla["grupo"], tabla["tasa_seleccion"], color=COLOR_PRIMARY)
    axes[0].set_title("Tasa de selección\n(% clasificado como riesgo alto)")
    axes[0].set_ylim(0, 1)
    axes[0].tick_params(axis="x", rotation=20)

    x = np.arange(len(tabla))
    w = 0.35
    axes[1].bar(x - w / 2, tabla["tpr"], width=w, color=COLOR_SECONDARY, label="TPR (Sensibilidad)")
    axes[1].bar(x + w / 2, tabla["fpr"], width=w, color=COLOR_BAD, label="FPR")
    axes[1].set_xticks(x)
    axes[1].set_xticklabels(tabla["grupo"])
    axes[1].set_title("Sensibilidad (TPR) y\nFalsos positivos (FPR)")
    axes[1].set_ylim(0, 1)
    axes[1].legend(fontsize=8, frameon=False)
    axes[1].tick_params(axis="x", rotation=20)

    fig.suptitle(titulo)
    fig.tight_layout()
    fig.savefig(f"{FIG_DIR}/{nombre_archivo}.png")
    plt.close(fig)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    import os
    os.makedirs(FIG_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)

    print("1) Generando conjunto de datos sintético...")
    df = generar_datos(2400)
    df["grupo_estrato"] = df["estrato"].apply(grupo_estrato)
    df.to_csv(f"{DATA_DIR}/dataset_desercion.csv", index=False)
    print(f"   -> {len(df)} registros. Tasa de deserción global: {df['desercion'].mean():.2%}")

    # Partición 60 / 20 / 20 (train / val / test), estratificada por el target
    df_train, df_temp = train_test_split(df, test_size=0.4, random_state=42,
                                          stratify=df["desercion"])
    df_val, df_test = train_test_split(df_temp, test_size=0.5, random_state=42,
                                        stratify=df_temp["desercion"])
    print(f"   -> Train: {len(df_train)} | Val: {len(df_val)} | Test: {len(df_test)}")

    pre, feature_cols = construir_pipeline_preprocesamiento()
    y_train, y_val, y_test = df_train["desercion"], df_val["desercion"], df_test["desercion"]

    X_train = pre.fit_transform(df_train[feature_cols])
    X_val = pre.transform(df_val[feature_cols])
    X_test = pre.transform(df_test[feature_cols])

    # -------------------- Modelos --------------------
    print("2) Entrenando modelos...")
    logreg = LogisticRegression(max_iter=1000, random_state=42)
    logreg.fit(X_train, y_train)

    rf = RandomForestClassifier(n_estimators=300, max_depth=8, random_state=42,
                                 class_weight="balanced")
    rf.fit(X_train, y_train)

    mlp = MLPClassifier(
        hidden_layer_sizes=(64, 32),
        activation="relu",
        solver="adam",
        alpha=1e-3,
        learning_rate_init=1e-3,
        max_iter=500,
        early_stopping=True,
        n_iter_no_change=20,
        validation_fraction=0.15,
        random_state=42,
    )
    mlp.fit(X_train, y_train)

    # Selección de umbral usando el conjunto de validación (optimiza F1)
    from sklearn.metrics import f1_score as _f1
    val_prob = mlp.predict_proba(X_val)[:, 1]
    umbrales = np.linspace(0.2, 0.8, 61)
    f1s = [_f1(y_val, (val_prob >= t).astype(int)) for t in umbrales]
    mejor_umbral = umbrales[int(np.argmax(f1s))]
    print(f"   -> Umbral de decisión óptimo (validación, max F1): {mejor_umbral:.2f}")

    # -------------------- Evaluación en test --------------------
    print("3) Evaluando en conjunto de prueba...")
    resultados = []
    probs = {}

    for nombre, modelo, usa_umbral in [
        ("Regresión Logística", logreg, False),
        ("Random Forest", rf, False),
        ("Red Neuronal (MLP)", mlp, True),
    ]:
        prob = modelo.predict_proba(X_test)[:, 1]
        pred = (prob >= mejor_umbral).astype(int) if usa_umbral else modelo.predict(X_test)
        resultados.append(evaluar_modelo(nombre, y_test, pred, prob))
        probs[nombre] = prob
        if nombre == "Red Neuronal (MLP)":
            mlp_pred_test = pred

    tabla_metricas = pd.DataFrame(resultados)
    tabla_metricas.to_csv(f"{DATA_DIR}/tabla_metricas.csv", index=False)
    print(tabla_metricas.round(3).to_string(index=False))

    # -------------------- Gráficas de desempeño --------------------
    print("4) Generando gráficas...")
    graficar_comparacion_modelos(resultados)
    graficar_roc(probs, y_test)
    graficar_matriz_confusion(y_test, mlp_pred_test, "matriz_confusion_mlp",
                               "Matriz de confusión — Red Neuronal (MLP)")
    graficar_curva_perdida(mlp)
    graficar_importancia_variables(rf, feature_cols)

    # -------------------- Análisis de equidad --------------------
    print("5) Analizando equidad del modelo (MLP) por género y estrato...")
    tabla_genero, resumen_genero = metricas_equidad(
        y_test.values, mlp_pred_test, df_test["genero"].values)
    tabla_estrato, resumen_estrato = metricas_equidad(
        y_test.values, mlp_pred_test, df_test["grupo_estrato"].values)

    graficar_equidad(tabla_genero, "Equidad del modelo por género", "equidad_genero")
    graficar_equidad(tabla_estrato, "Equidad del modelo por estrato socioeconómico",
                      "equidad_estrato")

    tabla_genero.to_csv(f"{DATA_DIR}/equidad_genero.csv", index=False)
    tabla_estrato.to_csv(f"{DATA_DIR}/equidad_estrato.csv", index=False)

    print("   Género:", {k: round(v, 3) for k, v in resumen_genero.items()})
    print("   Estrato:", {k: round(v, 3) for k, v in resumen_estrato.items()})

    # -------------------- Exportar resumen para el informe --------------------
    resumen_json = {
        "n_total": int(len(df)),
        "n_train": int(len(df_train)),
        "n_val": int(len(df_val)),
        "n_test": int(len(df_test)),
        "tasa_desercion_global": float(df["desercion"].mean()),
        "umbral_decision": float(mejor_umbral),
        "metricas_modelos": resultados,
        "equidad_genero": {
            "tabla": tabla_genero.to_dict(orient="records"),
            "resumen": resumen_genero,
        },
        "equidad_estrato": {
            "tabla": tabla_estrato.to_dict(orient="records"),
            "resumen": resumen_estrato,
        },
        "arquitectura_mlp": {
            "capas_ocultas": [64, 32],
            "activacion": "ReLU",
            "optimizador": "Adam",
            "n_iteraciones_reales": int(mlp.n_iter_),
            "alpha_regularizacion": 1e-3,
        },
    }
    with open(f"{DATA_DIR}/resumen_resultados.json", "w", encoding="utf-8") as f:
        json.dump(resumen_json, f, indent=2, ensure_ascii=False)

    print("\nProceso completo. Figuras en 'figures/', datos y métricas en 'data/'.")


if __name__ == "__main__":
    main()
