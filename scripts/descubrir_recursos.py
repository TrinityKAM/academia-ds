#!/usr/bin/env python3
"""
Herramienta de UNA SOLA VEZ para encontrar GRAPH_DRIVE_ID y GRAPH_ITEM_ID (los dos
secrets que scripts/actualizar_finca.py necesita ademas de las credenciales de la
aplicacion). No se usa en la automatizacion normal — se corre a mano, una vez,
despues de que exista el registro de aplicacion en Azure AD con permiso de
aplicacion (Sites.Read.All o Files.Read.All) y consentimiento de administrador.

Uso:
  export GRAPH_TENANT_ID=...
  export GRAPH_CLIENT_ID=...
  export GRAPH_CLIENT_SECRET=...
  python3 scripts/descubrir_recursos.py

Que hace:
  1. Pide un token de aplicacion (client credentials).
  2. Lista los sitios de SharePoint visibles para la app.
  3. Para el sitio que elijas, lista sus drives (bibliotecas de documentos).
  4. Busca DS_Finca_Sistema.xlsx dentro del drive que elijas y muestra su driveId
     e itemId exactos — esos dos valores van directo a GitHub Secrets como
     GRAPH_DRIVE_ID y GRAPH_ITEM_ID.

Si el paso 2 devuelve una lista vacia o un error 403, significa que el permiso de
aplicacion todavia no tiene consentimiento de administrador — hay que resolver eso
primero (ver AUTOMATIZACION.md).
"""

import json
import os
import sys
import urllib.request
import urllib.parse
import urllib.error

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def fail(msg):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(1)


def get_app_token():
    for k in ("GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET"):
        if not os.environ.get(k):
            fail(f"Falta la variable de entorno {k}. Exportala antes de correr este script.")
    tenant = os.environ["GRAPH_TENANT_ID"]
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = urllib.parse.urlencode({
        "client_id": os.environ["GRAPH_CLIENT_ID"],
        "client_secret": os.environ["GRAPH_CLIENT_SECRET"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())["access_token"]
    except urllib.error.HTTPError as e:
        fail(f"No se pudo obtener token ({e.code}): {e.read().decode(errors='replace')}")


def graph_get(token, path):
    req = urllib.request.Request(f"{GRAPH_BASE}{path}", headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        fail(f"Graph fallo en {path} ({e.code}): {e.read().decode(errors='replace')}")


def prompt_choice(items, label_fn, prompt):
    for i, it in enumerate(items):
        print(f"  [{i}] {label_fn(it)}")
    while True:
        raw = input(prompt).strip()
        if raw.isdigit() and 0 <= int(raw) < len(items):
            return items[int(raw)]
        print("Opcion invalida, intenta de nuevo.")


def main():
    token = get_app_token()
    print("Token de aplicacion obtenido OK.\n")

    sites = graph_get(token, "/sites?search=*").get("value", [])
    if not sites:
        fail(
            "La busqueda de sitios devolvio 0 resultados. Lo mas probable es que el "
            "permiso de aplicacion todavia no tenga consentimiento de administrador "
            "otorgado en Azure AD — revisa eso antes de seguir."
        )
    print(f"Sitios de SharePoint visibles ({len(sites)}):")
    site = prompt_choice(sites, lambda s: f"{s.get('displayName','(sin nombre)')} — {s.get('webUrl','')}", "\nElige el numero del sitio donde vive el Excel de la finca: ")

    drives = graph_get(token, f"/sites/{site['id']}/drives").get("value", [])
    if not drives:
        fail("Ese sitio no tiene bibliotecas de documentos visibles para la app.")
    print(f"\nBibliotecas (drives) en '{site.get('displayName')}':")
    drive = prompt_choice(drives, lambda d: f"{d.get('name','(sin nombre)')} — driveId: {d['id']}", "\nElige el numero de la biblioteca donde esta el Excel: ")

    print("\nBuscando DS_Finca_Sistema.xlsx dentro de esa biblioteca (puede tardar unos segundos)...")
    results = graph_get(token, f"/drives/{drive['id']}/root/search(q='DS_Finca_Sistema')").get("value", [])
    match = next((r for r in results if r.get("name", "").lower().endswith(".xlsx")), None)
    if not match:
        fail(
            "No se encontro DS_Finca_Sistema.xlsx en esa biblioteca via busqueda. "
            "Puede que este en una subcarpeta que la busqueda no indexo todavia — "
            "en ese caso hay que navegar manualmente con /drives/{driveId}/root:/ruta:/  "
            "para obtener el itemId a mano."
        )

    print("\n" + "=" * 60)
    print("Encontrado. Estos son los valores para GitHub Secrets:")
    print(f"  GRAPH_DRIVE_ID = {drive['id']}")
    print(f"  GRAPH_ITEM_ID  = {match['id']}")
    print(f"  (archivo: {match.get('name')} — {match.get('webUrl','')})")
    print("=" * 60)


if __name__ == "__main__":
    main()
