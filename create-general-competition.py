#!/usr/bin/env python3
"""Create one FantaPilot general competition using the fixed official rules."""
import argparse
import datetime as dt
import json
import os
import re
import sys
import uuid

from player_catalog import CatalogError, Client

BACKEND = "https://platform-core.lolloapps0.workers.dev"


def timestamp(value):
    if not re.fullmatch(
        r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?(?:Z|[+-]\d{2}:\d{2})",
        value,
    ):
        raise CatalogError("Date richieste in formato ISO con fuso, es. 2026-10-30T18:00:00+01:00.")
    try:
        return dt.datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(dt.timezone.utc)
    except ValueError:
        raise CatalogError("Data o fuso non validi.") from None


def build_payload(args, now=None):
    if args.confirmation != "CREA":
        raise CatalogError("Conferma richiesta: CREA.")
    try:
        catalog_id = str(uuid.UUID(args.catalog_id))
    except ValueError:
        raise CatalogError("ID catalogo non valido.") from None
    if not re.fullmatch("[0-9a-f]{64}", args.expected_hash):
        raise CatalogError("Content hash SHA-256 non valido.")
    name = args.name.strip()
    if not 3 <= len(name) <= 100 or any(ord(char) < 32 for char in name):
        raise CatalogError("Nome richiesto: da 3 a 100 caratteri, senza caratteri di controllo.")
    opening, deadline, ending = map(timestamp, (args.opens_at, args.deadline_at, args.ends_at))
    now = now or dt.datetime.now(dt.timezone.utc)
    if not now < opening < deadline <= ending:
        raise CatalogError("Serve: adesso < apertura < scadenza <= fine. Nessuna creazione eseguita.")
    rules = {
        "version": 1, "catalogId": catalog_id, "budget": 250,
        "roster": {"total": 25, "P": 3, "D": 8, "C": 8, "A": 6},
        "starters": 11,
        "bench": {"total": 8, "minimum": {"P": 2, "D": 0, "C": 0, "A": 0}},
        "modules": [dict(P=1, D=d, C=c, A=a) for d, c, a in
                    [(5, 4, 1), (5, 3, 2), (4, 5, 1), (4, 4, 2), (4, 3, 3), (3, 5, 2), (3, 4, 3)]],
        "maxSubstitutions": 5,
    }
    return {
        "appKey": "fantapilot", "name": name, "scoreOrder": "desc",
        "opensAt": opening.isoformat(timespec="milliseconds"),
        "deadlineAt": deadline.isoformat(timespec="milliseconds"),
        "endsAt": ending.isoformat(timespec="milliseconds"),
        "submissionRequired": True, "rules": {"fantapilot": rules},
    }


def create(client, payload, expected_hash):
    catalog_id = payload["rules"]["fantapilot"]["catalogId"]
    catalog = client.call("/catalogs/" + catalog_id).get("catalog", {})
    if (catalog.get("id") != catalog_id or catalog.get("app_key") != "fantapilot"
            or catalog.get("source") != "official" or not catalog.get("published_at")
            or catalog.get("content_hash") != expected_hash):
        raise CatalogError("Il catalogo deve essere ufficiale FantaPilot, pubblicato e con hash corrispondente.")
    # POST creation is not idempotent: block an existing name and never retry it.
    for offset in range(0, 10000, 100):
        page = client.call(f"/competitions?appKey=fantapilot&limit=100&offset={offset}")
        competitions = page.get("competitions")
        if not isinstance(competitions, list):
            raise CatalogError("Elenco competizioni non valido: creazione interrotta.")
        if any(c.get("group_id") is None and c.get("name") == payload["name"]
               for c in competitions):
            raise CatalogError("Esiste gia' una competizione generale con questo nome: controllarla prima di ripetere.")
        if len(competitions) < 100:
            break
    else:
        raise CatalogError("Troppe competizioni per verificare duplicati: creazione interrotta.")
    timestamp_now = dt.datetime.now(dt.timezone.utc)
    if timestamp(payload["opensAt"]) <= timestamp_now:
        raise CatalogError("L'apertura e' gia' trascorsa durante i controlli.")
    return client.call("/competitions", payload)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    for field in ("catalog-id", "expected-hash", "name", "opens-at", "deadline-at", "ends-at", "confirmation"):
        parser.add_argument("--" + field, required=True)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)
    payload = build_payload(args)
    print(json.dumps({"validatedPayload": payload}, ensure_ascii=False, indent=2))
    if args.dry_run:
        print("Solo validazione locale: nessuna autenticazione o creazione.")
        return
    email = os.environ.get("PLATFORM_ADMIN_EMAIL")
    password = os.environ.get("PLATFORM_ADMIN_PASSWORD")
    if not email or not password:
        raise CatalogError("Configurare i Secrets PLATFORM_ADMIN_EMAIL e PLATFORM_ADMIN_PASSWORD.")
    client = Client(BACKEND)
    session = client.call("/auth/login", {"email": email, "password": password})
    token = session.get("session", {}).get("accessToken")
    if not isinstance(token, str) or not token:
        raise CatalogError("Access token assente nella risposta login.")
    client.token = token
    print("Invio singolo, senza retry automatico. Se la risposta e' incerta, verificare l'elenco nell'app prima di rilanciare.")
    print(json.dumps(create(client, payload, args.expected_hash), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except CatalogError as error:
        print(str(error), file=sys.stderr)
        sys.exit(1)
