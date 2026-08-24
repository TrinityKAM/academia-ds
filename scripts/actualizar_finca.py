#!/usr/bin/env python3
"""
Actualiza data.json leyendo en vivo las 7 hojas transaccionales de DS_Finca_Sistema.xlsx
desde SharePoint via Microsoft Graph (permiso de APLICACION, sin que ninguna persona
tenga que iniciar sesion), y despues regenera finca.html con build.py.

ESTADO: escrito y listo, pero SIN PROBAR TODAVIA — no existe todavia el registro de
aplicacion en Azure AD con permiso de aplicacion (Files.Read.All o Sites.Read.All) ni
el consentimiento de administrador que este script necesita para funcionar. Ver
AUTOMATIZACION.md en la raiz del repo para la lista exacta de lo que falta y quien
tiene que hacerlo (Jonel Aguado / quien administre Microsoft 365 de Don Salazar).

Variables de entorno requeridas (se configuran como GitHub Secrets, nunca en el codigo):
  GRAPH_TENANT_ID      - Tenant ID (o dominio) de Azure AD de Don Salazar.
  GRAPH_CLIENT_ID      - Client ID de la aplicacion registrada.
  GRAPH_CLIENT_SECRET  - Client secret de esa aplicacion.
  GRAPH_DRIVE_ID       - driveId de la biblioteca de SharePoint donde vive el Excel.
  GRAPH_ITEM_ID        - itemId del archivo DS_Finca_Sistema.xlsx dentro de ese drive.

GRAPH_DRIVE_ID y GRAPH_ITEM_ID no se saben todavia — se obtienen una sola vez, ya con
las credenciales de la app funcionando, corriendo scripts/descubrir_recursos.py.

Alcance (igual al de la prueba en finca-beta.html): solo se refrescan las 7 hojas
transaccionales (Cosecha, Lotes, Fermentacion, Transformaciones, Gastos, Ventas,
Compras). Parcelas y Supuestos (hoja Maestros) siguen viniendo del data.json existente,
porque cambian con mucha menor frecuencia y su mapeo de columnas no esta verificado
todavia contra el Excel real via Graph (a diferencia de las 7 hojas de arriba, cuyo
mapeo de encabezados SI esta verificado — es el mismo que ya se probo en el boton
"Conectar en vivo" de finca-beta.html).
"""

import json
import os
import subprocess
import sys
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime, timezone

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

REQUIRED_ENV = [
    "GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET",
    "GRAPH_DRIVE_ID", "GRAPH_ITEM_ID",
]

SHEET_SPECS = [
    {"name": "1. Cosecha", "key": "cosecha"},
    {"name": "2. Lotes", "key": "lotes"},
    {"name": "3. Fermentacion", "key": "fermentacion"},
    {"name": "4. Transformaciones", "key": "transformaciones"},
    {"name": "5. Gastos", "key": "gastos"},
    {"name": "6. Ventas", "key": "ventas"},
    {"name": "7. Compras", "key": "compras"},
]


def fail(msg):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(1)


def check_env():
    missing = [k for k in REQUIRED_ENV if not os.environ.get(k)]
    if missing:
        fail(
            "Faltan variables de entorno / GitHub Secrets: " + ", ".join(missing) +
            ". Este script todavia no puede correr sin ellas — ver AUTOMATIZACION.md."
        )


def get_app_token():
    tenant = os.environ["GRAPH_TENANT_ID"]
    client_id = os.environ["GRAPH_CLIENT_ID"]
    client_secret = os.environ["GRAPH_CLIENT_SECRET"]
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            payload = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        fail(f"No se pudo obtener token de Azure AD ({e.code}): {body}")
    token = payload.get("access_token")
    if not token:
        fail(f"Respuesta de token sin access_token: {payload}")
    return token


def graph_get(token, path):
    req = urllib.request.Request(f"{GRAPH_BASE}{path}", headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        fail(f"Graph API fallo en {path} ({e.code}): {body}")


def fetch_sheet_values(token, drive_id, item_id, sheet_name):
    encoded = urllib.parse.quote(f"'{sheet_name}'")
    path = f"/drives/{drive_id}/items/{item_id}/workbook/worksheets({encoded})/usedRange(valuesOnly=true)"
    data = graph_get(token, path)
    return data.get("values", [])


# ---- Helpers identicos en criterio a los usados en finca-beta.html (Fase 2 / JS) ----

def num_or_null(v):
    return None if v in (None, "") else v if isinstance(v, (int, float)) else _try_float(v)


def _try_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def str_or_null(v):
    return None if v in (None, "") else str(v)


def is_ejemplo(v):
    return isinstance(v, str) and "EJEMPLO" in v.upper()


def xl_date_to_iso(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        # Excel guarda fechas como numero de serie (dias desde 1899-12-30) — mismo criterio
        # que build.py / data.json y que xlDateToISO() en finca-beta.html.
        epoch = datetime(1899, 12, 30, tzinfo=timezone.utc)
        from datetime import timedelta
        return (epoch + timedelta(days=v)).date().isoformat()
    try:
        return datetime.fromisoformat(str(v)).date().isoformat()
    except ValueError:
        return None


def rows_to_objects(values, mapper):
    if not values or len(values) < 3:
        return []
    headers = values[1] or []

    def get(row, name):
        try:
            idx = headers.index(name)
        except ValueError:
            return None
        return row[idx] if idx < len(row) else None

    out = []
    for row in values[2:]:
        if not row or row[0] in (None, ""):
            continue
        out.append(mapper(lambda name, row=row: get(row, name)))
    return out


# Mapeos verificados contra los encabezados reales del Excel (mismo criterio que
# mapCosecha/mapLotes/... en finca-beta.html).

def map_cosecha(get):
    return {
        "fecha": xl_date_to_iso(get("Fecha")), "cosechero": str_or_null(get("Cosechero")),
        "parcela": str_or_null(get("Parcela")), "variedad": str_or_null(get("Variedad")),
        "proceso": str_or_null(get("Proceso")), "kg_cerezo": num_or_null(get("Kg cerezo")),
        "precio_kg": num_or_null(get("Precio/kg (S/)")), "pago": num_or_null(get("Pago (S/)")),
        "lote": str_or_null(get("Lote")), "obs": str_or_null(get("Observaciones")),
        "ejemplo": is_ejemplo(get("Observaciones")),
    }


def map_lotes(get):
    return {
        "id": str_or_null(get("ID Lote")), "campania": str_or_null(get("Campania")),
        "parcela": str_or_null(get("Parcela")), "variedad": str_or_null(get("Variedad")),
        "proceso": str_or_null(get("Proceso")), "fecha_apertura": xl_date_to_iso(get("Fecha apertura")),
        "fecha_cierre": xl_date_to_iso(get("Fecha cierre")),
        "kg_cerezo_auto": num_or_null(get("Kg cerezo (auto)")), "estado": str_or_null(get("Estado")),
        "obs": str_or_null(get("Observaciones")), "ejemplo": is_ejemplo(get("Observaciones")),
    }


def map_fermentacion(get):
    return {
        "fecha": xl_date_to_iso(get("Fecha")), "hora": get("Hora"), "lote": str_or_null(get("Lote")),
        "proceso": str_or_null(get("Proceso")), "fase": str_or_null(get("Fase")),
        "temp": num_or_null(get("Temp (C)")), "ph": num_or_null(get("pH")),
        "horas_ferm": num_or_null(get("Horas ferment.")), "humedad": num_or_null(get("Humedad grano (%)")),
        "dias_secado": num_or_null(get("Dias secado")), "responsable": str_or_null(get("Responsable")),
        "obs": str_or_null(get("Observaciones")), "ejemplo": is_ejemplo(get("Observaciones")),
    }


def map_transformaciones(get):
    return {
        "fecha": xl_date_to_iso(get("Fecha")), "lote": str_or_null(get("Lote")),
        "variedad": str_or_null(get("Variedad")), "proceso": str_or_null(get("Proceso")),
        "etapa_origen": str_or_null(get("Etapa origen")), "kg_entra": num_or_null(get("Kg que entra")),
        "etapa_destino": str_or_null(get("Etapa destino")), "kg_producido": num_or_null(get("Kg producido")),
        "rendimiento": num_or_null(get("Rendimiento %")), "merma": num_or_null(get("Merma (kg)")),
        "ubicacion": str_or_null(get("Ubicacion")), "responsable": str_or_null(get("Responsable")),
        "detalle": str_or_null(get("Detalle")), "ejemplo": is_ejemplo(get("Detalle")),
    }


def map_gastos(get):
    return {
        "fecha": xl_date_to_iso(get("Fecha")), "cuenta": str_or_null(get("Cuenta")),
        "subcuenta": str_or_null(get("Sub cuenta")), "parcela": str_or_null(get("Parcela")),
        "lote": str_or_null(get("Lote (opc.)")), "detalle": str_or_null(get("Detalle")),
        "monto": num_or_null(get("Monto (S/)")), "status": str_or_null(get("Status")),
        "fecha_pago_py": xl_date_to_iso(get("Fecha pago PY")),
        "fecha_pago_real": xl_date_to_iso(get("Fecha pago Real")),
    }


def map_ventas(get):
    return {
        "fecha": xl_date_to_iso(get("Fecha")), "lote": str_or_null(get("Lote")),
        "variedad": str_or_null(get("Variedad")), "proceso": str_or_null(get("Proceso")),
        "kg_oro_verde": num_or_null(get("Cantidad oro verde (kg)")),
        "precio_kg": num_or_null(get("Precio/kg (S/)")), "total": num_or_null(get("Total (S/)")),
        "n_factura": str_or_null(get("N factura")), "fecha_factura": xl_date_to_iso(get("Fecha factura")),
        "fecha_pago": xl_date_to_iso(get("Fecha pago")),
    }


def map_compras(get):
    return {
        "fecha": xl_date_to_iso(get("Fecha")), "lote": str_or_null(get("Lote")),
        "variedad": str_or_null(get("Variedad")), "proceso": str_or_null(get("Proceso")),
        "etapa": str_or_null(get("Etapa")), "kg": num_or_null(get("Kg")),
        "proveedor": str_or_null(get("Proveedor")), "costo": num_or_null(get("Costo (S/)")),
        "detalle": str_or_null(get("Detalle")),
    }


MAPPERS = {
    "cosecha": map_cosecha, "lotes": map_lotes, "fermentacion": map_fermentacion,
    "transformaciones": map_transformaciones, "gastos": map_gastos, "ventas": map_ventas,
    "compras": map_compras,
}


def main():
    check_env()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.dirname(script_dir)
    data_json_path = os.path.join(repo_root, "data.json")

    with open(data_json_path, encoding="utf-8") as f:
        current = json.load(f)

    token = get_app_token()
    drive_id = os.environ["GRAPH_DRIVE_ID"]
    item_id = os.environ["GRAPH_ITEM_ID"]

    nuevo = dict(current)  # conserva parcelas, supuestos, d1/d2/d3/d5 tal como estan
    nuevo["meta"] = {**current.get("meta", {}), "generado": datetime.now(timezone.utc).date().isoformat()}

    for spec in SHEET_SPECS:
        print(f"Leyendo hoja '{spec['name']}'...")
        values = fetch_sheet_values(token, drive_id, item_id, spec["name"])
        rows = rows_to_objects(values, MAPPERS[spec["key"]])
        if spec["key"] == "compras" and not rows:
            # Compras suele estar vacia legitimamente (la finca aun no compro a terceros) —
            # no se trata como error, a diferencia de las demas hojas.
            pass
        elif not rows:
            fail(
                f"La hoja '{spec['name']}' devolvio 0 filas. Esto casi seguro es un error de "
                f"mapeo de encabezados o de permisos, no que la finca no tenga actividad — "
                f"se aborta sin sobreescribir data.json para no publicar datos vacios por error."
            )
        nuevo[spec["key"]] = rows
        print(f"  {len(rows)} filas")

    with open(data_json_path, "w", encoding="utf-8") as f:
        json.dump(nuevo, f, ensure_ascii=False)
    print(f"data.json actualizado ({data_json_path})")

    subprocess.run([sys.executable, os.path.join(repo_root, "build.py")], check=True, cwd=repo_root)
    print("finca.html regenerado a partir del Excel en vivo.")


if __name__ == "__main__":
    main()
