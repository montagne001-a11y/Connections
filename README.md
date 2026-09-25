# ONB Survey 1: Encuesta de Satisfacción (Selección e Inducción)

This is a Power BI Project (PBIP) that replaces the empty `ONB Survey 1.pbix`. It reads the Microsoft Forms export
**GCG — Encuesta de Satisfacción 1_ Selección e Inducción.xlsx** (table `OfficeForms.Table`).

## Open it

1. Open `ONB Survey 1.pbip` in Power BI Desktop. If Desktop asks, turn on *File > Options > Preview features > Power BI Project (.pbip) save option*.
2. If the Excel file isn't at the default path, go to *Transform data > Edit parameters > `RutaArchivo`* and enter the correct path.
3. Click **Refresh**.
4. To overwrite the original file, use *File > Save as > `ONB Survey 1.pbix`*.

## Semantic model (`ONB Survey 1.SemanticModel`)

| Table | Grain | Source |
|---|---|---|
| `Encuesta` | 1 row per response | Id, dates, duration, email, name, country, company, role |
| `Respuestas Likert` | response × question | The 10 Likert columns unpivoted, with `Puntaje` 1–5 |
| `Pregunta` | question | Dimension (Selección / Inducción), full text, short label, order |
| `Escala Likert` | scale level | Response → score → sentiment (Desfavorable / Neutral / Favorable) |
| `Opciones Seleccionadas` | response × option | The 3 multi-select questions split on `;` |
| `Comentarios` | comment | Open comments, excluding blank or "Sin comentarios" |
| `Calendario` | day | DAX date table; replaces the auto date/time tables |
| `_Medidas` | — | All measures |

New Likert questions whose header starts with `Proceso de Selección.` or `Proceso de Inducción.` are picked up
automatically. To give one a short chart label, add it to the `Etiquetas` list in the `Pregunta` query.

### Key KPIs

| KPI | Definition |
|---|---|
| Puntaje Promedio | Mean score on the 1–5 scale (target 4.0) |
| % Favorable (CSAT) | Top-2 Box: share of "De acuerdo" + "Muy de acuerdo" (target 80 %) |
| % Desfavorable | Bottom-2 Box |
| NSS (Net Satisfaction) | % Favorable − % Desfavorable, from −100 to +100 |
| Puntaje / CSAT Selección · Inducción | The same KPIs for each dimension |
| Brecha Inducción vs Selección | Inducción score − Selección score |
| % Listos para Entrenamiento | Top-2 Box on "Me siento listo/a para comenzar mi plan de entrenamiento", the main outcome of onboarding |
| Fortaleza / Oportunidad Principal | The question with the highest / lowest score in the current filter context |
| % Encuestados · Lo que más gustó / Mejoras / Faltantes | Share of respondents who selected each option |
| % Sin Mejoras Necesarias, % Inducción Completa | Share of respondents who answered "no improvements needed" / "nothing was missing" |

Traffic-light colors: green at 4 or above (CSAT 80 % or above), amber at 3 or above (60 % or above), red below that.

## Report pages (`ONB Survey 1.Report`)

Every page has a header with a dynamic participation summary. The País, Compañía, Rol and Fecha slicers are synced across pages.

1. **Resumen Ejecutivo**: 6 KPI cards (Respuestas, Puntaje, CSAT, NSS, % Listos, Brecha), a CSAT gauge against the target, score by dimension, sentiment by dimension, score by question (ranked), top strength and top opportunity, unique respondents, and average duration.
2. **Análisis por Pregunta**: Likert 5-point distribution per question (100 % stacked), a scorecard table with a heat-map background, score by country, CSAT by company, and responses by date.
3. **Mapa de Calor por Compañía**: a Question × Company heat-map matrix, Selección vs Inducción by company, and participation by country.
4. **Voz del Colaborador**: KPI cards (no improvements needed, nothing missing, comments), the three multi-select questions ranked by % of respondents, and a table of open comments.

The custom theme is `StaticResources/RegisteredResources/GCG_Onboarding.json`.
