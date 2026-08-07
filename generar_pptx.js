const pptxgen = require("pptxgenjs");
const fs = require("fs");

const R = JSON.parse(fs.readFileSync("data/resumen_resultados.json", "utf-8"));

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5 in

const AZUL = "1F3B73";
const AZUL_CLARO = "2E5EAA";
const NARANJA = "D9822B";
const GRIS = "555555";
const GRIS_CLARO = "F2F3F5";
const BLANCO = "FFFFFF";
const ROJO = "C0392B";
const VERDE = "2E9E6B";

const FONT = "Calibri";

function footer(slide, num) {
  slide.addText(`${num} / 10`, {
    x: 12.55, y: 7.15, w: 0.7, h: 0.3, fontFace: FONT, fontSize: 9, color: "AAAAAA", align: "right",
  });
  slide.addText("Deserción Estudiantil · IA", {
    x: 0.4, y: 7.15, w: 4, h: 0.3, fontFace: FONT, fontSize: 9, color: "AAAAAA", align: "left",
  });
}

function tituloSlide(slide, kicker, titulo) {
  slide.addText(kicker.toUpperCase(), {
    x: 0.6, y: 0.35, w: 10, h: 0.35, fontFace: FONT, fontSize: 13, color: NARANJA, bold: true, charSpacing: 1,
  });
  slide.addText(titulo, {
    x: 0.6, y: 0.68, w: 12.1, h: 0.7, fontFace: FONT, fontSize: 26, color: AZUL, bold: true,
  });
}

// ---------------------------------------------------------------------
// SLIDE 1 — Portada
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: AZUL };
  s.addShape(pres.ShapeType.rect, { x: 0, y: 4.85, w: 13.33, h: 2.65, fill: { color: "17305C" }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { x: 10.6, y: -1.6, w: 5, h: 5, fill: { color: "2E5EAA", transparency: 55 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { x: -1.4, y: 5.4, w: 4, h: 4, fill: { color: "D9822B", transparency: 70 }, line: { type: "none" } });

  s.addText("CREACIÓN DE UN MODELO DE INTELIGENCIA ARTIFICIAL AVANZADA", {
    x: 0.8, y: 1.55, w: 11.7, h: 0.5, fontFace: FONT, fontSize: 15, color: NARANJA, bold: true, charSpacing: 1.5,
  });
  s.addText("Predicción temprana del riesgo de\ndeserción estudiantil universitaria", {
    x: 0.8, y: 2.1, w: 11.7, h: 1.8, fontFace: FONT, fontSize: 38, color: BLANCO, bold: true, lineSpacing: 44,
  });
  s.addText("Una red neuronal (Perceptrón Multicapa) para la detección temprana de riesgo,\nevaluada también en términos de equidad algorítmica", {
    x: 0.8, y: 3.85, w: 10.5, h: 0.7, fontFace: FONT, fontSize: 14, color: "C9D3E8", italic: true,
  });

  s.addText("Luis Javier Salgado Guzmán", { x: 0.8, y: 5.3, w: 6, h: 0.4, fontFace: FONT, fontSize: 16, color: BLANCO, bold: true });
  s.addText("Ingeniería de Software · Séptimo semestre · Universidad Tecnológica del Oriente", {
    x: 0.8, y: 5.75, w: 8, h: 0.4, fontFace: FONT, fontSize: 12, color: "C9D3E8",
  });
  s.addText("Estrategia: Aprendizaje Basado en Retos (ABR)", {
    x: 0.8, y: 6.15, w: 8, h: 0.4, fontFace: FONT, fontSize: 12, color: "C9D3E8",
  });
}

// ---------------------------------------------------------------------
// SLIDE 2 — Problema y justificación
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: BLANCO };
  tituloSlide(s, "El reto", "¿Por qué predecir la deserción estudiantil?");

  const stats = [
    { n: "1 de cada 3", d: "estudiantes universitarios abandona sus\nestudios en algún momento de la carrera" },
    { n: "↑ costo", d: "académico, económico y emocional para el\nestudiante, su familia y la institución" },
    { n: "Detección\ntemprana", d: "permite activar tutorías, apoyo financiero\ny acompañamiento antes del abandono" },
  ];
  stats.forEach((st, i) => {
    const x = 0.6 + i * 4.15;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.75, w: 3.85, h: 2.15, rectRadius: 0.08, fill: { color: GRIS_CLARO }, line: { type: "none" } });
    s.addText(st.n, { x: x + 0.25, y: 1.95, w: 3.4, h: 0.75, fontFace: FONT, fontSize: 22, bold: true, color: AZUL_CLARO });
    s.addText(st.d, { x: x + 0.25, y: 2.7, w: 3.4, h: 1.05, fontFace: FONT, fontSize: 12.5, color: GRIS, lineSpacing: 16 });
  });

  s.addText("Pregunta de investigación", { x: 0.6, y: 4.25, w: 6, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: NARANJA });
  s.addText(
    "¿Es posible, a partir de variables académicas, de asistencia y socioeconómicas disponibles durante el semestre, predecir con antelación qué estudiantes tienen mayor riesgo de desertar — priorizando los recursos limitados de acompañamiento hacia los casos de mayor riesgo, sin perjudicar de forma desproporcionada a ningún grupo?",
    { x: 0.6, y: 4.68, w: 12.1, h: 1.7, fontFace: FONT, fontSize: 16, color: "222222", italic: true, lineSpacing: 24 }
  );
  footer(s, 2);
}

// ---------------------------------------------------------------------
// SLIDE 3 — Revisión de literatura
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: BLANCO };
  tituloSlide(s, "Estado del arte", "Lo que dice la literatura");

  const cols = [
    { t: "Predicción de deserción", pts: [
      "Los indicadores académicos y de asistencia son los predictores más usados",
      "Los factores psicosociales siguen subrepresentados en los datasets",
      "Random Forest y Gradient Boosting dominan por su balance desempeño/interpretabilidad",
    ]},
    { t: "Equidad algorítmica (fairness)", pts: [
      "La paridad demográfica no siempre es el criterio adecuado si hay diferencias reales de prevalencia",
      "La igualdad de oportunidad (TPR) y la igualdad de momios son más robustas para APP*",
      "*APP: Academic Performance Prediction — predicción de desempeño académico",
    ]},
  ];
  cols.forEach((c, i) => {
    const x = 0.6 + i * 6.2;
    s.addShape(pres.ShapeType.rect, { x, y: 1.7, w: 5.9, h: 0.06, fill: { color: NARANJA }, line: { type: "none" } });
    s.addText(c.t, { x, y: 1.85, w: 5.9, h: 0.5, fontFace: FONT, fontSize: 17, bold: true, color: AZUL });
    let y = 2.45;
    c.pts.forEach((pt) => {
      s.addText([{ text: pt }], {
        x, y, w: 5.9, h: 0.9, fontFace: FONT, fontSize: 12.5, color: "333333", bullet: { code: "2022", color: NARANJA }, lineSpacing: 16,
      });
      y += 0.95;
    });
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 6.05, w: 12.1, h: 0.85, rectRadius: 0.06, fill: { color: "EAF0FA" }, line: { type: "none" } });
  s.addText("Implicación de diseño: se compara el modelo avanzado (red neuronal) contra líneas base interpretables, y se incorpora una fase explícita de auditoría de equidad.", {
    x: 0.9, y: 6.18, w: 11.5, h: 0.6, fontFace: FONT, fontSize: 12.5, italic: true, color: AZUL, valign: "middle",
  });
  footer(s, 3);
}

// ---------------------------------------------------------------------
// SLIDE 4 — Dataset y metodología
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: BLANCO };
  tituloSlide(s, "Metodología", "Datos y partición del experimento");

  s.addText("13 variables predictoras", { x: 0.6, y: 1.75, w: 6, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: AZUL_CLARO });
  const vars = [
    "Académicas: promedio anterior, asistencia, materias reprobadas, créditos, semestre",
    "Socioeconómicas*: estrato, apoyo financiero, distancia al campus, trabaja",
    "Otras: edad, género*, horas de estudio, participación extracurricular",
    "* Atributos sensibles — usados también para el análisis de equidad (sección 5)",
  ];
  s.addText(vars.map((v) => ({ text: v, options: { breakLine: true, bullet: v.startsWith("*") ? false : { code: "2022", color: NARANJA } } })), {
    x: 0.6, y: 2.25, w: 6.1, h: 2.6, fontFace: FONT, fontSize: 13, color: "333333", lineSpacing: 22, paraSpaceAfter: 8,
  });

  // Partición como barras horizontales simples
  s.addText("Partición del dataset (n = " + R.n_total + ")", { x: 7.1, y: 1.75, w: 5.6, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: AZUL_CLARO });
  const partes = [
    { label: `Entrenamiento — ${R.n_train}`, pct: 60, color: AZUL_CLARO },
    { label: `Validación — ${R.n_val}`, pct: 20, color: NARANJA },
    { label: `Prueba — ${R.n_test}`, pct: 20, color: VERDE },
  ];
  let y = 2.35;
  partes.forEach((p) => {
    s.addText(p.label, { x: 7.1, y, w: 5.6, h: 0.3, fontFace: FONT, fontSize: 11.5, color: "333333" });
    s.addShape(pres.ShapeType.rect, { x: 7.1, y: y + 0.32, w: 5.6, h: 0.28, fill: { color: "E7E7EA" }, line: { type: "none" } });
    s.addShape(pres.ShapeType.rect, { x: 7.1, y: y + 0.32, w: 5.6 * (p.pct / 100), h: 0.28, fill: { color: p.color }, line: { type: "none" } });
    y += 0.85;
  });
  s.addText(`Tasa de deserción global en los datos: ${(R.tasa_desercion_global * 100).toFixed(1)} %`, {
    x: 7.1, y: y + 0.05, w: 5.6, h: 0.4, fontFace: FONT, fontSize: 12, italic: true, color: GRIS,
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.55, w: 12.1, h: 1.3, rectRadius: 0.06, fill: { color: GRIS_CLARO }, line: { type: "none" } });
  s.addText("Nota metodológica", { x: 0.9, y: 5.68, w: 11, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: NARANJA });
  s.addText("Al no contar con acceso a datasets institucionales reales en este entorno, se generó un conjunto de datos sintético que reproduce, de forma controlada, relaciones de riesgo reportadas en la literatura — incluyendo un sesgo estructural deliberado para ilustrar el análisis de equidad.", {
    x: 0.9, y: 5.98, w: 11.5, h: 0.8, fontFace: FONT, fontSize: 11.5, color: "333333", lineSpacing: 15,
  });
  footer(s, 4);
}

// ---------------------------------------------------------------------
// SLIDE 5 — Arquitectura del modelo
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: BLANCO };
  tituloSlide(s, "Diseño del modelo", "Red neuronal: Perceptrón Multicapa (MLP)");

  // Diagrama simple de capas
  const capas = [
    { t: "Entrada", n: "13 variables", w: 1.9, color: "8FA6D6" },
    { t: "Oculta 1", n: "64 neuronas · ReLU", w: 2.5, color: AZUL_CLARO },
    { t: "Oculta 2", n: "32 neuronas · ReLU", w: 2.1, color: AZUL_CLARO },
    { t: "Salida", n: "1 neurona · Sigmoide", w: 1.6, color: NARANJA },
  ];
  let x = 0.8;
  const yMid = 2.7;
  capas.forEach((c, i) => {
    s.addShape(pres.ShapeType.roundRect, { x, y: yMid - c.w / 2 * 0.42, w: 1.9, h: c.w, rectRadius: 0.1, fill: { color: c.color }, line: { type: "none" } });
    s.addText(c.t, { x, y: yMid - c.w / 2 * 0.42 + 0.08, w: 1.9, h: 0.35, align: "center", fontFace: FONT, fontSize: 13, bold: true, color: BLANCO });
    s.addText(c.n, { x, y: yMid - c.w / 2 * 0.42 + c.w - 0.4, w: 1.9, h: 0.35, align: "center", fontFace: FONT, fontSize: 10, color: BLANCO });
    if (i < capas.length - 1) {
      s.addShape(pres.ShapeType.rightArrow, { x: x + 1.95, y: yMid - 0.1, w: 0.55, h: 0.2, fill: { color: "BBBBBB" }, line: { type: "none" } });
    }
    x += 2.5;
  });

  s.addText("Hiperparámetros clave", { x: 0.6, y: 4.55, w: 6, h: 0.35, fontFace: FONT, fontSize: 14, bold: true, color: AZUL_CLARO });
  const hp = [
    `Optimizador: ${R.arquitectura_mlp.optimizador}  ·  Regularización L2 (α): ${R.arquitectura_mlp.alpha_regularizacion}`,
    `Parada anticipada (early stopping) tras ${R.arquitectura_mlp.n_iteraciones_reales} iteraciones`,
    `Umbral de decisión ajustado en validación: ${R.umbral_decision.toFixed(2)} (maximiza F1)`,
  ];
  s.addText(hp.map((v) => ({ text: v, options: { breakLine: true, bullet: { code: "2022", color: NARANJA } } })), {
    x: 0.6, y: 5.0, w: 8.5, h: 1.4, fontFace: FONT, fontSize: 13, color: "333333", lineSpacing: 22,
  });

  s.addShape(pres.ShapeType.roundRect, { x: 9.3, y: 4.55, w: 3.4, h: 1.85, rectRadius: 0.08, fill: { color: "FBEFE0" }, line: { type: "none" } });
  s.addText("¿Por qué priorizar sensibilidad?", { x: 9.5, y: 4.68, w: 3, h: 0.5, fontFace: FONT, fontSize: 12, bold: true, color: NARANJA });
  s.addText("Omitir a un estudiante en riesgo real cuesta más que una falsa alarma de bajo costo.", {
    x: 9.5, y: 5.15, w: 3, h: 1.1, fontFace: FONT, fontSize: 11.5, color: "5A4222", lineSpacing: 15,
  });
  footer(s, 5);
}

// ---------------------------------------------------------------------
// SLIDE 6 — Resultados: comparación de modelos
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: BLANCO };
  tituloSlide(s, "Resultados", "Comparación de desempeño entre modelos");

  s.addImage({ path: "figures/comparacion_modelos.png", x: 0.6, y: 1.7, w: 7.6, h: 4.28 });

  const m = R.metricas_modelos;
  const filas = [
    [{ text: "Modelo", options: { bold: true, fill: { color: AZUL }, color: BLANCO } },
     { text: "F1", options: { bold: true, fill: { color: AZUL }, color: BLANCO } },
     { text: "AUC", options: { bold: true, fill: { color: AZUL }, color: BLANCO } }],
    ...m.map((mm) => [mm.modelo, mm.f1.toFixed(3), mm.auc_roc.toFixed(3)]),
  ];
  s.addTable(filas, {
    x: 8.55, y: 1.85, w: 4.15, h: 1.6, fontFace: FONT, fontSize: 11.5, border: { type: "solid", color: "DDDDDD", pt: 0.75 },
    colW: [2.2, 0.95, 1.0], autoPage: false,
  });

  s.addText("Random Forest logra el mejor balance precisión-sensibilidad; el MLP maximiza la sensibilidad (88.7%), clave para no dejar pasar estudiantes en riesgo real.", {
    x: 8.55, y: 3.7, w: 4.15, h: 2.1, fontFace: FONT, fontSize: 12.5, color: "333333", lineSpacing: 18,
  });
  footer(s, 6);
}

// ---------------------------------------------------------------------
// SLIDE 7 — Resultados: ROC y matriz de confusión
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: BLANCO };
  tituloSlide(s, "Resultados", "Curvas ROC y matriz de confusión (MLP)");

  s.addImage({ path: "figures/curvas_roc.png", x: 0.5, y: 1.75, w: 6.0, h: 5.2 });
  s.addImage({ path: "figures/matriz_confusion_mlp.png", x: 6.9, y: 1.9, w: 5.7, h: 4.9 });
  footer(s, 7);
}

// ---------------------------------------------------------------------
// SLIDE 8 — Equidad
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: BLANCO };
  tituloSlide(s, "Equidad algorítmica", "¿El modelo trata igual a todos los grupos?");

  s.addImage({ path: "figures/equidad_estrato.png", x: 0.5, y: 1.7, w: 7.5, h: 3.15 });

  const rg = R.equidad_genero.resumen;
  const re = R.equidad_estrato.resumen;
  const filas = [
    [{ text: "Métrica", options: { bold: true, fill: { color: AZUL }, color: BLANCO } },
     { text: "Género", options: { bold: true, fill: { color: AZUL }, color: BLANCO } },
     { text: "Estrato", options: { bold: true, fill: { color: AZUL }, color: BLANCO } }],
    ["Disparate Impact Ratio", rg.disparate_impact_ratio.toFixed(2), re.disparate_impact_ratio.toFixed(2)],
    ["Equal Opportunity Diff.", rg.equal_opportunity_difference.toFixed(2), re.equal_opportunity_difference.toFixed(2)],
    ["Equalized Odds Diff.", rg.equalized_odds_difference.toFixed(2), re.equalized_odds_difference.toFixed(2)],
  ];
  s.addTable(filas, {
    x: 8.35, y: 1.75, w: 4.4, h: 1.9, fontFace: FONT, fontSize: 11, border: { type: "solid", color: "DDDDDD", pt: 0.75 },
    colW: [2.1, 1.15, 1.15], autoPage: false,
  });

  s.addText("Lectura", { x: 8.35, y: 3.85, w: 4.4, h: 0.3, fontFace: FONT, fontSize: 12.5, bold: true, color: NARANJA });
  s.addText("La igualdad de oportunidad es buena en ambos atributos: el modelo detecta al estudiante en riesgo real de forma similar entre grupos. La brecha aparece en falsos positivos, más altos en estrato bajo y género femenino.", {
    x: 8.35, y: 4.2, w: 4.4, h: 1.9, fontFace: FONT, fontSize: 12, color: "333333", lineSpacing: 17,
  });
  footer(s, 8);
}

// ---------------------------------------------------------------------
// SLIDE 9 — Implicaciones éticas y recomendaciones
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: BLANCO };
  tituloSlide(s, "Ética y mitigación", "Recomendaciones para un uso responsable");

  const recs = [
    ["Apoyo, no automatización", "El modelo es una señal de apoyo a la decisión humana, nunca un criterio automático de sanción o exclusión."],
    ["Intervención de bajo costo", "La alerta debe traducirse en un beneficio neutral (tutorías), nunca en una etiqueta pública o estigmatizante."],
    ["Auditoría periódica", "Reentrenar y volver a medir la equidad con datos reales conforme cambia la población estudiantil."],
    ["Mitigación de sesgo", "Explorar reponderación o ajuste de umbrales por grupo si la brecha de falsos positivos persiste con datos reales."],
  ];
  let y = 1.75;
  recs.forEach((r, i) => {
    const x = i % 2 === 0 ? 0.6 : 6.85;
    if (i % 2 === 0 && i > 0) y += 2.05;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.9, h: 1.85, rectRadius: 0.08, fill: { color: i % 2 === 0 ? "EAF0FA" : "FBEFE0" }, line: { type: "none" } });
    s.addText(r[0], { x: x + 0.25, y: y + 0.15, w: 5.4, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: AZUL });
    s.addText(r[1], { x: x + 0.25, y: y + 0.58, w: 5.4, h: 1.15, fontFace: FONT, fontSize: 12, color: "333333", lineSpacing: 16 });
  });
  footer(s, 9);
}

// ---------------------------------------------------------------------
// SLIDE 10 — Conclusiones
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: AZUL };
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 1.35, fill: { color: "17305C" }, line: { type: "none" } });
  s.addText("CONCLUSIONES", { x: 0.6, y: 0.42, w: 8, h: 0.5, fontFace: FONT, fontSize: 15, bold: true, color: NARANJA, charSpacing: 1.5 });

  const concl = [
    `El MLP alcanza AUC-ROC de ${R.metricas_modelos[2].auc_roc.toFixed(2)} y sensibilidad de ${(R.metricas_modelos[2].sensibilidad_recall * 100).toFixed(0)}%, priorizando detectar a los estudiantes en riesgo real.`,
    "El modelo mantiene igualdad de oportunidad aceptable entre géneros y estratos, pero genera más falsas alarmas en los grupos ya vulnerables — un punto a vigilar antes de cualquier despliegue.",
    "Un sistema real requeriría datos institucionales históricos, variables psicosociales y una validación piloto antes de su uso en producción.",
  ];
  let y = 1.75;
  concl.forEach((c) => {
    s.addShape(pres.ShapeType.roundRect, { x: 0.6, y, w: 0.35, h: 0.35, rectRadius: 0.17, fill: { color: NARANJA }, line: { type: "none" } });
    s.addText(String(concl.indexOf(c) + 1), { x: 0.6, y, w: 0.35, h: 0.35, align: "center", valign: "middle", fontFace: FONT, fontSize: 14, bold: true, color: BLANCO });
    s.addText(c, { x: 1.2, y: y - 0.08, w: 11.5, h: 1.05, fontFace: FONT, fontSize: 15, color: BLANCO, lineSpacing: 20, valign: "middle" });
    y += 1.55;
  });

  s.addText("Gracias — Luis Javier Salgado Guzmán", {
    x: 0.6, y: 6.85, w: 8, h: 0.4, fontFace: FONT, fontSize: 12, color: "C9D3E8", italic: true,
  });
}

pres.writeFile({ fileName: "outputs/Presentacion_Desercion_Estudiantil_IA.pptx" }).then(() => {
  console.log("Presentación generada correctamente.");
});
