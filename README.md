# PlasHub / Factory "Amaru I" — App del proyecto

Suite interactiva (100% estática, sin servidor) del hub de consolidación de plásticos
**Etapa 1** de **Neo Amaru Global S.A.C.**, Lima Norte, Perú.

## Punto de entrada

Abre **`index.html`** (la landing) o directamente **`plashub-app.html`** (app unificada con 4 pestañas).

| Página | Qué es |
|---|---|
| `index.html` | Landing / centro de documentos. **Empieza aquí.** |
| `manual.html` | **Manual de uso** de cada parámetro editable + guion para la reunión. |
| `plashub-app.html` | App unificada: Distribución E1 · Distribución E2 · Plan de Negocio · Tareas. |
| `plashub-plan-negocio.html` | Modelo financiero interactivo (E1 + E2). |
| `plashub-distribucion.html` / `-e2.html` | Layout + capacidad de almacén (E1) y planta (E2). |
| `plashub-tareas.html` | Lista maestra de tareas y permisos. |

## Notas

- **Los datos se guardan en el navegador (localStorage), no en la nube.** Cada visitante ve los
  valores por defecto; usa **Exportar config** / **Importar** para compartir un escenario exacto.
- Cifras **estimadas** (≈) — no usar para levantar capital hasta verificar precios con cotizaciones reales.

## Ver localmente

Abre `index.html` en el navegador, o sirve la carpeta:

```bash
npx http-server . -p 3017
# → http://localhost:3017/
```
