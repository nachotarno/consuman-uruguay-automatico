# Consuman Uruguay → Google Sheets

Automatización desatendida para GitHub Actions.

Flujo:
1. Playwright inicia sesión en Consuman.
2. Consulta la API interna de órdenes y tareas.
3. Genera `ordenes_trabajo.json`.
4. Actualiza la misma pestaña `Ordenes de trabajo` del Google Sheet configurado.
5. Se ejecuta cada hora (en punto, UTC) y también puede lanzarse manualmente desde Actions.

## Secrets requeridos en GitHub

En `Settings → Secrets and variables → Actions → New repository secret` crear:

- `CONSUMAN_USUARIO`
- `CONSUMAN_PASSWORD`
- `GOOGLE_SERVICE_ACCOUNT_JSON` (JSON completo de una cuenta de servicio de Google)

Luego compartir el Google Sheet con el email `client_email` de esa cuenta de servicio como **Editor**.

## Seguridad

No subir credenciales, `.env` ni archivos JSON de claves al repositorio.

## Primera prueba

En GitHub: `Actions → Consuman Uruguay a Google Sheets → Run workflow`.
Revisar el log. Si termina correctamente, el cron continuará cada hora.
