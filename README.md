# Con esto pagamos el viaje a Bogotá ⚽️🇨🇴

Sistema de apuestas para el **Mundial 2026**: combinadas con cobertura, gestión de banca con Kelly, combinadas de alto pago y la promo "Combinada Segura" de Betsson. Todo el análisis está reunido en un solo Excel maestro.

## Archivo principal

**`Sistema_Mundial_2026_COMPLETO.xlsx`** — 16 hojas en 4 bloques:

1. **Combinadas iniciales** — 12 combinadas de S/20 con cobertura.
2. **Kelly** — combinada por día con fracción ½, ⅔ y 100%, comparación, percentiles P10–P90 y trayectorias.
3. **Combinadas de +S/5,000** — 18 combinadas de S/30 que pagan +S/5,000, con análisis de riesgo y mapa de diversificación.
4. **Combinada Segura (Betsson)** — 12 combinadas de S/60 que aprovechan el reembolso si falla 1 pata, diversificadas (min hitting set 4) y repartidas en ventanas de reembolso.

La hoja **Portada** resume inversión, ganancia y recomendación de cada estrategia.

## Carpetas

- **`excel-intermedios/`** — los Excel por etapa que luego se consolidaron en el maestro.
- **`scripts/`** — scripts de Python (openpyxl + numpy) que generan y verifican cada hoja.
- **`datos/`** — los `.json` con los datos intermedios (picks, probabilidades, resultados de simulación).

## Notas

- Las probabilidades "Mi prob." son estimaciones propias y se pueden editar en el Excel; las P() y los EV dependen de esos valores.
- Las combinadas de +S/5,000 son billetes de lotería con ventaja: lo más probable es perder la apuesta. Apostar solo lo que se pueda permitir perder.
