# Automatizacion de Finca — que falta para que se actualice sola

Este documento es la lista exacta de lo que falta para que `finca.html` se actualice
sola desde `DS_Finca_Sistema.xlsx`, sin que nadie tenga que abrir la pagina ni iniciar
sesion. El codigo ya esta escrito (`scripts/actualizar_finca.py`,
`.github/workflows/actualizar-finca.yml`) pero no puede correr todavia porque le faltan
credenciales que solo puede crear alguien con rol de administrador de Microsoft 365 /
Azure AD en Don Salazar. Segun la guia interna del equipo, esto se coordina con
**Jonel Aguado**.

## Lo que falta (en orden)

1. **Registrar una aplicacion en Azure AD / Entra** (portal.azure.com → Azure Active
   Directory → Registros de aplicaciones → Nuevo registro). Puede ser una app nueva
   dedicada a esto — no reutilizar la de "Academia DS" que ya existe, porque esta
   automatizacion necesita un tipo de permiso mas sensible (ver el punto 2).
2. **Darle permiso de tipo "Aplicacion" (no "Delegado")** para `Sites.Read.All` (o
   `Files.Read.All`) de Microsoft Graph, y que un administrador presione
   **"Conceder consentimiento de administrador"**. Este es el paso que solo puede
   hacer un admin — sin este clic, nada de lo demas funciona.
3. **Crear un "Client secret"** para esa aplicacion (Certificados y secretos → Nuevo
   secreto de cliente) y copiar su valor una sola vez (no se puede volver a ver
   despues).
4. **Guardar 3 valores como GitHub Secrets** en este repositorio (Settings → Secrets
   and variables → Actions → New repository secret):
   - `GRAPH_TENANT_ID` (Directory (tenant) ID de la app registrada)
   - `GRAPH_CLIENT_ID` (Application (client) ID)
   - `GRAPH_CLIENT_SECRET` (el secreto del paso 3)
5. **Correr `scripts/descubrir_recursos.py` una sola vez** (con esos 3 valores como
   variables de entorno, desde cualquier computadora con Python) para encontrar los
   2 valores que faltan, y guardarlos tambien como GitHub Secrets:
   - `GRAPH_DRIVE_ID`
   - `GRAPH_ITEM_ID`
6. **Probar el workflow a mano**: pestaña "Actions" del repo → "Actualizar Finca (Fase
   2b)" → botón "Run workflow". Si corre sin errores y actualiza `finca.html`, avisar
   para activar el `schedule:` (cron) en
   `.github/workflows/actualizar-finca.yml` y que corra solo todos los dias.

## Por que no lo hizo Claude solo

Claude no tiene ni puede crear credenciales de Azure AD, ni puede otorgarse a si mismo
consentimiento de administrador — esa aprobacion es intencionalmente algo que solo un
humano con ese rol puede hacer, por seguridad. Este documento existe para que ese
paso humano sea lo mas rapido y claro posible: son los pasos 1-2-3-4-5 de arriba, en
ese orden, y despues el sistema queda funcionando solo.

## Alcance actual del script (una vez activo)

Actualiza solo las 7 hojas transaccionales: Cosecha, Lotes, Fermentacion,
Transformaciones, Gastos, Ventas, Compras. Parcelas y Supuestos (hoja Maestros) siguen
viniendo del `data.json` fijo, porque cambian con mucha menor frecuencia y su
mapeo de columnas para lectura en vivo todavia no esta escrito. Si mas adelante hace
falta automatizar tambien esas dos, se puede extender `scripts/actualizar_finca.py`
siguiendo el mismo patron que las demas hojas.

## Alternativa mas simple a considerar antes: Power Automate

Antes de invertir tiempo en los pasos 1-5, vale la pena preguntarle a Jonel Aguado si
Power Automate ya esta disponible en la licencia de Microsoft 365 de Don Salazar.
Un flujo de Power Automate programado, usando la conexion de una cuenta ya autorizada
(sin registrar ninguna aplicacion nueva en Azure AD), podria resolver la lectura del
Excel con mucho menos trabajo tecnico — el limite es que no sabe generar el HTML de la
pagina directamente, asi que igual haria falta un paso que tome esos datos y regenere
`finca.html` (podria ser este mismo repositorio via GitHub Actions, disparado por
Power Automate en vez de por un cron).
