# Monitor de licitaciones — e-Oficialía Hidalgo (versión web)

Cada mañana revisa el portal de licitaciones de Hidalgo y actualiza **tu propio tablero web**, donde consultas cuando quieras: número de licitación, objeto, junta de aclaraciones, apertura y fallo, con buscador y filtros. También genera un Excel acumulado. No requiere WhatsApp ni deja nada corriendo en tu computadora.

## Paso 1 — Crear el repositorio

1. Crea cuenta en https://github.com si no tienes.
2. Botón **+** → **New repository**. Nombre: `monitor-licitaciones`. Márcalo como **Public** (necesario para que GitHub Pages sea gratis; los datos son públicos de todas formas). Crea el repositorio.

## Paso 2 — Subir los archivos

- Con **Add file → Upload files** sube: `monitor_licitaciones.py`, `plantilla.html`, `requirements.txt` y este `README.md`. Commit.
- Con **Add file → Create new file**, escribe como nombre `.github/workflows/monitor.yml` (GitHub crea las carpetas solo), pega el contenido de `monitor.yml` y haz commit. **La ruta debe quedar exactamente así.**

## Paso 3 — Primera corrida

Pestaña **Actions** → "Monitor de licitaciones Hidalgo" → **Run workflow**. Espera a que salga en verde ✅ (2-3 min). Esto crea las carpetas `data/` (Excel y estado) y `docs/` (el tablero).

## Paso 4 — Activar tu página web

1. En el repositorio: **Settings → Pages**.
2. En "Build and deployment": Source = **Deploy from a branch**; Branch = **main**, carpeta **/docs**. Guarda.
3. En 1-2 minutos tu tablero queda en:
   `https://TU-USUARIO.github.io/monitor-licitaciones/`
4. Guarda esa dirección en favoritos del celular. Se actualiza sola de lunes a viernes a las 8:00 AM.

## El tablero incluye

- Buscador por número, objeto o convocatoria.
- Filtros: **Solo nuevas** (detectadas en los últimos 3 días, marcadas con borde dorado), **Fechas próximas** y **Sin fallo aún**.
- Por cada licitación, el riel de etapas Junta → Apertura → Fallo: punto guinda = acta ya publicada, punto verde = fecha futura, vacío = pendiente.

## Opcional

- **WhatsApp**: si algún día quieres además un aviso breve cuando haya nuevas, configura los secretos `WHATSAPP_PHONE` y `CALLMEBOT_APIKEY` (Settings → Secrets → Actions). Si no existen, simplemente no se envía nada.
- **Excel**: siempre disponible en `data/licitaciones.xlsx` dentro del repositorio.
- **Cambiar horario**: edita la línea `cron` en `.github/workflows/monitor.yml` (la hora está en UTC; 14:00 UTC = 8:00 AM centro de México).

## Notas

- Las fechas del portal corresponden a la **publicación de cada acta**, por eso una licitación recién publicada muestra "pendiente" en junta/apertura/fallo hasta que el gobierno sube los documentos. Las fechas programadas con hora exacta vienen en el PDF de la convocatoria.
- La extracción del "objeto" desde el PDF es best-effort; cuando no se logra, la tarjeta dice "Objeto en el PDF de la convocatoria".
- Si el portal cambia de estructura, el workflow fallará en rojo y lo verás en Actions.
