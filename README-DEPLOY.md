# Merch Hub — Guía de deploy (click by click)

App de inventario de merchandising de SGC Marketing. Reemplaza el sheet "Inventario 2025 - 2026".

## Arquitectura

```
Merch Hub DB (Google Sheet nuevo, lo crea la migración)
  → Apps Script (doGet lectura JSONP + doPost escritura)
  → GitHub Pages (index.html)
  → Browser (login Google, mismo Client ID del Admissions Dashboard)
```

El sheet viejo **no se toca**: queda como archivo histórico.

---

## Paso 1 — Crear el Apps Script y migrar

1. Entrá a [script.google.com](https://script.google.com) con tu cuenta.
2. **Nuevo proyecto** → nombralo `Merch Hub`.
3. Borrá todo el contenido de `Código.gs`, pegá el contenido completo de `apps-script-merch.gs` y guardá (Ctrl+S).
4. En la barra de arriba, elegí la función **`SETUP_MIGRATE`** en el desplegable y tocá **Ejecutar**.
5. Google te va a pedir autorización: **Revisar permisos** → tu cuenta → *Avanzado* → *Ir a Merch Hub (no seguro)* → **Permitir**.
6. Esperá que termine (puede tardar 1-2 minutos). Abrí **Ver → Registros** (o el panel de ejecución): vas a ver el resumen de la migración y la **URL de la planilla "Merch Hub DB"** creada en tu Drive.
7. **Verificá en la planilla nueva, hoja `MIGRACION_LOG`:** la línea "Asignación de shops" dice qué hoja de import tomó como Quilmes y cuál como North. Si está al revés, avisame y lo corregimos.

## Paso 2 — Publicar la API

1. En el mismo proyecto: **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web**.
3. *Ejecutar como:* **Yo** · *Quién tiene acceso:* **Cualquier persona**.
4. **Implementar** → copiá la **URL de la aplicación web** (termina en `/exec`).

## Paso 3 — Conectar el frontend

1. Abrí `index.html` y buscá la línea:
   ```js
   const API_URL = 'PEGAR_ACA_LA_URL_DEL_APPS_SCRIPT';
   ```
2. Pegá ahí la URL del paso 2 (o pasámela y lo hago yo).

## Paso 4 — Subir a GitHub Pages

1. En GitHub: **New repository** → nombre `sgc-merch-hub` → **Public** → Create.
2. **Add file → Upload files** → arrastrá estos 5 archivos:
   - `index.html`
   - `logo-h.png`, `logo-h-white.png`, `logo-v.png`, `favicon.png`
3. Commit.
4. **Settings → Pages** → Source: *Deploy from a branch* → Branch `main` / `(root)` → Save.
5. En unos minutos la app queda viva en:
   `https://diego-medina-sgc.github.io/sgc-merch-hub/`

> **OAuth: no hay que tocar nada.** El Client ID del Admissions Dashboard ya autoriza el origen `https://diego-medina-sgc.github.io` (el origen es por dominio, no por repo).

## Paso 5 — Dar acceso al equipo

En la planilla **Merch Hub DB**, hoja **Acceso**: un email por fila, columna A. Ya están `diego.medina@` y `marketing@`.

---

## Operación diaria

| Tarea | Dónde |
|---|---|
| Registrar un regalo/venta interna | **Pedidos** → formulario |
| Importar ventas de un shop | **Importar ventas** → elegir shop → pegar filas del export → Analizar → Confirmar |
| Código nuevo del shop | La importación lo detecta y pide asignarlo **una sola vez** |
| Mover merch entre ubicaciones | **Stock** → botón `⇄ mover` |
| Conteo físico / corregir diferencias | **Stock** → botón `± ajustar` |
| Llegó reposición | **Stock** → botón `➕ reponer` |
| Alta de producto nuevo | **Catálogo** → "Nuevo producto" (SKU automático) |
| Cambiar precios | **Catálogo** → editar |

**Cache:** el doGet cachea 10 minutos, pero **cada escritura desde la app lo invalida sola**. Si editás la planilla DB **a mano**, forzá refresco abriendo `URL_DEL_EXEC?nocache=1` en el navegador, o esperá 10 min.

## Si algo se rompe

- La verdad siempre está en la planilla **Merch Hub DB** (hojas PRODUCTOS / MOVIMIENTOS / PEDIDOS / ALIAS). Son tablas planas sin fórmulas: nada se "rompe en cascada".
- El stock **nunca se edita a mano**: si un número está mal, se corrige con un **Ajuste** desde la app (queda auditado quién y cuándo).
- Cambios al backend: pegá el `.gs` actualizado completo → **Implementar → Administrar implementaciones → editar (lápiz) → Nueva versión → Implementar**. La URL no cambia.
