#!/usr/bin/env python3
"""
Worker de la bandeja de envíos.

La base de datos no envía nada: encola en app.notification_outbox. Este proceso
lee lo pendiente, lo manda y marca el resultado. Los secretos viven aquí, nunca
en la base ni en el navegador.

    correo -> Microsoft Graph (sendMail con permiso de aplicación Mail.Send)
    Teams  -> un flujo de Power Automate con disparador "Cuando se recibe una
              solicitud HTTP" que hace "Publicar mensaje en un chat o canal"
              como Flow bot al UPN que se le pasa

Variables de entorno:
    SUPABASE_URL           https://xxxx.supabase.co
    SUPABASE_SERVICE_KEY   clave service_role (SOLO en servidor)
    GRAPH_TENANT_ID        id de directorio de Entra ID
    GRAPH_CLIENT_ID        id de la aplicación registrada
    GRAPH_CLIENT_SECRET    secreto de esa aplicación
    GRAPH_SENDER           buzón desde el que se envía (p.ej. pulso@samplia.com)
    TEAMS_FLOW_URL         URL del flujo de Power Automate

Uso:
    python3 send_outbox.py            # envía lo pendiente y sale
    python3 send_outbox.py --dry-run  # enseña qué enviaría
"""
from __future__ import annotations

import json
import os
import sys
import urllib.request
import urllib.error

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SERVICE_KEY  = os.environ.get("SUPABASE_SERVICE_KEY", "")
TENANT       = os.environ.get("GRAPH_TENANT_ID", "")
CLIENT_ID    = os.environ.get("GRAPH_CLIENT_ID", "")
CLIENT_SECRET= os.environ.get("GRAPH_CLIENT_SECRET", "")
SENDER       = os.environ.get("GRAPH_SENDER", "")
TEAMS_FLOW   = os.environ.get("TEAMS_FLOW_URL", "")
DRY = "--dry-run" in sys.argv


def post(url: str, payload: dict | None, headers: dict, form: bool = False) -> dict:
    if form:
        data = "&".join(f"{k}={urllib.parse.quote(str(v))}" for k, v in payload.items()).encode()
    else:
        data = json.dumps(payload).encode() if payload is not None else b""
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            body = r.read().decode() or "{}"
            return json.loads(body) if body.strip().startswith(("{", "[")) else {"raw": body}
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"{e.code} {e.read().decode()[:200]}") from None


def rpc(fn: str, args: dict) -> list | dict:
    return post(f"{SUPABASE_URL}/rest/v1/rpc/{fn}", args,
                {"Content-Type": "application/json",
                 "apikey": SERVICE_KEY,
                 "Authorization": f"Bearer {SERVICE_KEY}"})


_token_cache: dict[str, str] = {}

def graph_token() -> str:
    if "t" in _token_cache:
        return _token_cache["t"]
    import urllib.parse  # noqa: F401  (lo usa post con form=True)
    j = post(f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
             {"client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
              "scope": "https://graph.microsoft.com/.default",
              "grant_type": "client_credentials"},
             {"Content-Type": "application/x-www-form-urlencoded"}, form=True)
    _token_cache["t"] = j["access_token"]
    return _token_cache["t"]


def send_email(item: dict) -> None:
    cuerpo = item["body"] + (f"\n\n{item['link']}" if item.get("link") else "")
    post(f"https://graph.microsoft.com/v1.0/users/{SENDER}/sendMail",
         {"message": {
             "subject": item.get("subject") or "Pulso interno",
             "body": {"contentType": "Text", "content": cuerpo},
             "toRecipients": [{"emailAddress": {"address": item["recipient"]}}]},
          "saveToSentItems": False},
         {"Content-Type": "application/json",
          "Authorization": f"Bearer {graph_token()}"})


def send_teams(item: dict) -> None:
    post(TEAMS_FLOW,
         {"upn": item["recipient"],
          "subject": item.get("subject") or "",
          "text": item["body"],
          "link": item.get("link") or ""},
         {"Content-Type": "application/json"})


def main() -> int:
    if not SUPABASE_URL or not SERVICE_KEY:
        print("Faltan SUPABASE_URL o SUPABASE_SERVICE_KEY", file=sys.stderr)
        return 2

    pendientes = rpc("f_outbox_pending", {"p_limit": 200})
    if not pendientes:
        print("Nada pendiente.")
        return 0

    ok = fallidos = 0
    for item in pendientes:
        etiqueta = f"{item['channel']:5} {item['kind']:13} {item['recipient']}"
        if DRY:
            print(f"[simulacro] {etiqueta}")
            continue
        try:
            (send_email if item["channel"] == "email" else send_teams)(item)
            rpc("f_outbox_mark", {"p_id": item["id"], "p_ok": True})
            ok += 1
            print(f"[enviado]   {etiqueta}")
        except Exception as e:                      # noqa: BLE001
            rpc("f_outbox_mark", {"p_id": item["id"], "p_ok": False, "p_error": str(e)[:400]})
            fallidos += 1
            print(f"[error]     {etiqueta} — {e}", file=sys.stderr)

    print(f"\n{ok} enviados, {fallidos} fallidos, {len(pendientes)} en cola.")
    return 1 if fallidos else 0


if __name__ == "__main__":
    import urllib.parse
    sys.exit(main())
