#!/usr/bin/env python3
"""External official catalog administration. Python 3 standard library only."""
import argparse
from collections import Counter
from decimal import Decimal, InvalidOperation
import getpass
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
HEADERS = {"Id", "ruolo", "pgv", "mv", "fm", "nome", "squadra", "quotazione_base", "fv"}
KEY = re.compile(r"[a-z0-9][a-z0-9_-]{0,62}[a-z0-9]\Z")
LABEL = re.compile(r"[a-zA-Z0-9_-]{1,32}\Z")


class CatalogError(Exception):
    pass


def read_players(path):
    """Read the first worksheet by workbook relationship, not ZIP filename order."""
    try:
        if Path(path).stat().st_size > 10 * 1024 * 1024:
            raise CatalogError("Excel exceeds the 10 MiB file limit.")
        with ZipFile(path) as archive:
            if sum(i.file_size for i in archive.infolist()) > 32 * 1024 * 1024:
                raise CatalogError("Expanded Excel exceeds the 32 MiB limit.")
            if len(archive.namelist()) != len(set(archive.namelist())):
                raise CatalogError("Excel contains duplicate ZIP entries.")
            strings = []
            if "xl/sharedStrings.xml" in archive.namelist():
                root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                strings = ["".join(t.text or "" for t in s.findall(".//m:t", NS))
                           for s in root.findall("m:si", NS)]
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            sheets = workbook.findall("m:sheets/m:sheet", NS)
            if not sheets:
                raise CatalogError("Excel has no worksheet.")
            sheet_id = sheets[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            target = next((r.get("Target") for r in relations if r.get("Id") == sheet_id), None)
            if not target:
                raise CatalogError("First worksheet relationship is missing.")
            sheet_path = urllib.parse.urljoin("xl/workbook.xml", target).lstrip("/")
            root = ET.fromstring(archive.read(sheet_path))
            rows = []
            for row in root.findall("m:sheetData/m:row", NS):
                cells = {}
                for cell in row.findall("m:c", NS):
                    if cell.find("m:f", NS) is not None:
                        raise CatalogError(f"Formula at {cell.get('r')}: supply literal values, not formulas.")
                    column = re.sub(r"\d", "", cell.get("r", ""))
                    if column in cells or not re.fullmatch(r"[A-Z]{1,3}", column):
                        raise CatalogError("Duplicate or invalid cell reference.")
                    if cell.get("t") == "e":
                        raise CatalogError(f"Excel error at {cell.get('r')}.")
                    raw = cell.findtext("m:v", "", NS)
                    if cell.get("t") == "s":
                        value = strings[int(raw)]
                    elif cell.get("t") == "inlineStr":
                        value = "".join(t.text or "" for t in cell.findall(".//m:t", NS))
                    else:
                        value = raw
                    cells[column] = value.strip()
                if any(cells.values()):
                    rows.append((row.get("r"), cells))
                if len(rows) > 2001:
                    raise CatalogError("Maximum 2000 players per catalog.")
    except (OSError, BadZipFile, ET.ParseError, KeyError, ValueError, IndexError) as error:
        raise CatalogError(f"Cannot read XLSX ({type(error).__name__}).") from None
    if len(rows) < 2:
        raise CatalogError("Excel must contain headers and at least one player.")
    headers = rows[0][1]
    names = [name for name in headers.values() if name]
    if len(names) != len(set(names)):
        raise CatalogError("Duplicate column headers.")
    missing = HEADERS - set(names)
    if missing:
        raise CatalogError("Missing columns: " + ", ".join(sorted(missing)))
    players = []
    seen = set()
    for row_number, cells in rows[1:]:
        values = {name: cells.get(col, "") for col, name in headers.items() if name}
        if any(not values.get(key) for key in HEADERS):
            raise CatalogError(f"Row {row_number}: required value missing.")
        try:
            external_id = Decimal(values["Id"])
            if not external_id.is_finite() or external_id <= 0 or external_id != external_id.to_integral_value():
                raise InvalidOperation
            source_id = str(int(external_id))
            if len(source_id) > 15:
                raise CatalogError(f"Row {row_number}: ID exceeds Excel's reliable numeric precision.")
            if source_id in seen:
                raise CatalogError(f"Row {row_number}: duplicate ID {source_id}.")
            seen.add(source_id)
            numeric = {}
            for field in ("pgv", "mv", "fm", "quotazione_base", "fv"):
                number = Decimal(values[field].replace(",", "."))
                if not number.is_finite() or number < 0 or number > 1000000000:
                    raise InvalidOperation
                if field == "pgv" and number != number.to_integral_value():
                    raise InvalidOperation
                rounded = number.quantize(Decimal("0.0001"))
                if abs(number - rounded) > Decimal("0.0000000001"):
                    raise CatalogError(f"Row {row_number}: {field} exceeds four decimal places.")
                number = rounded
                numeric[field] = int(number) if number == number.to_integral_value() else float(number)
        except (InvalidOperation, ValueError, OverflowError):
            raise CatalogError(f"Row {row_number}: invalid ID or numeric field.") from None
        if values["ruolo"] not in ("P", "D", "C", "A"):
            raise CatalogError(f"Row {row_number}: unsupported role.")
        if len(values["nome"]) > 120 or len(values["squadra"]) > 120:
            raise CatalogError(f"Row {row_number}: name/team exceeds 120 characters.")
        players.append({"sourcePlayerId": source_id, "data": {
            "name": values["nome"], "team": values["squadra"], "role": values["ruolo"],
            "quotation": numeric["quotazione_base"], "matchesPlayedWithRating": numeric["pgv"],
            "averageRating": numeric["mv"], "fantasyAverage": numeric["fm"], "fantasyValue": numeric["fv"],
        }})
    return sorted(players, key=lambda p: p["sourcePlayerId"])


class Client:
    def __init__(self, base, token=None):
        parsed = urllib.parse.urlsplit(base)
        if (parsed.scheme != "https" and not (parsed.scheme == "http" and parsed.hostname in ("localhost", "127.0.0.1"))):
            raise CatalogError("Use HTTPS for the backend (HTTP is allowed only on localhost).")
        if not parsed.hostname or parsed.username or parsed.password or parsed.path not in ("", "/") or parsed.query or parsed.fragment:
            raise CatalogError("Backend must be an origin URL without credentials, path or query.")
        self.base = base
        self.token = token

    def call(self, path, body=None):
        headers = {"Accept": "application/json", "Content-Type": "application/json",
                   "User-Agent": "platform-core-catalog-admin"}
        if self.token:
            headers["Authorization"] = "Bearer " + self.token
        data = json.dumps(body, allow_nan=False).encode() if body is not None else None
        if data and len(data) > 1048576:
            raise CatalogError("Request exceeds 1 MiB.")
        req = urllib.request.Request(urllib.parse.urljoin(self.base, path), data=data, headers=headers,
                                     method="POST" if body is not None else "GET")
        # Never forward a bearer token or password through a redirect.
        class NoRedirect(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, req, fp, code, msg, headers, newurl):
                return None
        try:
            with urllib.request.build_opener(NoRedirect).open(req, timeout=60) as response:
                return json.load(response)
        except urllib.error.HTTPError as error:
            try:
                code = json.load(error).get("code", "unknown")
            except (ValueError, AttributeError):
                code = "non_json_response"
            raise CatalogError(f"Backend returned HTTP {error.code} ({code}).") from None
        except (urllib.error.URLError, TimeoutError, ValueError):
            raise CatalogError("Backend unavailable or invalid response; no automatic publication attempted.") from None


def manifest_import(manifest_url, trusted_base, app_key, source):
    """Fetch only literal HTTPS paths beneath the explicitly trusted Pages directory."""
    base = urllib.parse.urlsplit(trusted_base)
    if base.scheme != "https" or not base.hostname or base.username or base.password or base.query or base.fragment:
        raise CatalogError("Trusted download base must be an HTTPS directory.")
    if not base.path.endswith("/"):
        raise CatalogError("Trusted download base must end with '/'.")

    def download(url, maximum):
        parsed = urllib.parse.urlsplit(url)
        if (parsed.scheme != "https" or parsed.netloc != base.netloc or parsed.username or parsed.password
                or parsed.query or parsed.fragment or not parsed.path.startswith(base.path)
                or "%" in parsed.path or "\\" in parsed.path
                or any(segment in (".", "..") for segment in parsed.path.split("/"))):
            raise CatalogError("Download URL is outside the trusted Pages directory.")

        class NoRedirect(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, req, fp, code, msg, headers, newurl):
                return None

        # Pages may cache the previous manifest for minutes; request a fresh URL.
        req = urllib.request.Request(url + "?catalog_check=" + str(time.time_ns()),
                                     headers={"User-Agent": "platform-core-catalog-import",
                                              "Cache-Control": "no-cache"})
        try:
            with urllib.request.build_opener(NoRedirect).open(req, timeout=60) as response:
                data = response.read(maximum + 1)
        except urllib.error.HTTPError as error:
            raise CatalogError(f"Pages download failed (HTTP {error.code}); retry after deployment completes.") from None
        except (urllib.error.URLError, TimeoutError, ValueError):
            raise CatalogError("Pages download failed; no import attempted.") from None
        if len(data) > maximum:
            raise CatalogError("Pages download exceeds the size limit.")
        return data

    def unique_keys(pairs):
        values = {}
        for key, value in pairs:
            if key in values:
                raise CatalogError("Manifest contains duplicate JSON keys.")
            values[key] = value
        return values

    try:
        manifest = json.loads(download(manifest_url, 16384), object_pairs_hook=unique_keys)
    except (ValueError, UnicodeError):
        raise CatalogError("Manifest must be valid UTF-8 JSON.") from None
    fields = {"schemaVersion", "appKey", "source", "season", "roundKey", "fileUrl", "sha256", "autoPublish"}
    if not isinstance(manifest, dict) or set(manifest) != fields:
        raise CatalogError("Manifest fields do not match schemaVersion 1.")
    if (type(manifest["schemaVersion"]) is not int or manifest["schemaVersion"] != 1
            or manifest["appKey"] != app_key or manifest["source"] != source
            or manifest["autoPublish"] is not False):
        raise CatalogError("Manifest must match the configured app/source and autoPublish must be false.")
    for field in ("season", "roundKey"):
        if not isinstance(manifest[field], str) or not LABEL.fullmatch(manifest[field]):
            raise CatalogError("Manifest has an invalid season or round key.")
    if not isinstance(manifest["sha256"], str) or not re.fullmatch("[0-9a-f]{64}", manifest["sha256"]):
        raise CatalogError("Manifest must contain a SHA-256 checksum.")
    if not isinstance(manifest["fileUrl"], str) or not manifest["fileUrl"].endswith(".xlsx"):
        raise CatalogError("Manifest must reference an XLSX file.")
    data = download(manifest["fileUrl"], 10 * 1024 * 1024)
    if hashlib.sha256(data).hexdigest() != manifest["sha256"]:
        raise CatalogError("Excel SHA-256 does not match the manifest; no import attempted.")
    with tempfile.TemporaryDirectory(prefix="catalog-import-") as directory:
        path = Path(directory) / "players.xlsx"
        path.write_bytes(data)
        players = read_players(path)
    return {"appKey": app_key, "source": source, "season": manifest["season"],
            "roundKey": manifest["roundKey"], "players": players}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, epilog=(
        "Workflow: validate -> preview -> stage -> publish --confirm. "
        "Set PLATFORM_ACCESS_TOKEN, or PLATFORM_ADMIN_EMAIL with PLATFORM_ADMIN_PASSWORD for automation "
        "(otherwise the password is requested interactively). "
        "No service-role key is needed. Source IDs are scoped by app, source and season."))
    parser.add_argument("--backend", default=os.environ.get("PLATFORM_BACKEND_URL", "https://platform-core.lolloapps0.workers.dev"))
    parser.add_argument("--email", default=os.environ.get("PLATFORM_ADMIN_EMAIL"))
    parser.add_argument("--non-interactive", action="store_true", help="Fail instead of prompting for missing credentials.")
    commands = parser.add_subparsers(dest="command", required=True)
    for command in ("validate", "preview", "stage"):
        sub = commands.add_parser(command)
        sub.add_argument("file", type=Path)
        for field in ("app-key", "source", "season", "round-key"):
            sub.add_argument("--" + field, required=True)
    publish = commands.add_parser("publish")
    publish.add_argument("catalog_id")
    publish.add_argument("--hash", required=True, dest="expected_hash")
    publish.add_argument("--confirm", action="store_true", required=True)
    discard = commands.add_parser("discard")
    discard.add_argument("catalog_id")
    discard.add_argument("--confirm", action="store_true", required=True)
    sync = commands.add_parser("sync-manifest", help="Download, verify and stage only; never publish.")
    sync.add_argument("manifest_url")
    sync.add_argument("--trusted-base", required=True)
    sync.add_argument("--app-key", required=True)
    sync.add_argument("--source", required=True)
    sync.add_argument("--dry-run", action="store_true", help="Validate Pages data without logging in or importing.")
    args = parser.parse_args(argv)
    payload = None
    if args.command == "sync-manifest":
        if not KEY.fullmatch(args.app_key) or not KEY.fullmatch(args.source):
            raise CatalogError("Invalid app/source key.")
        payload = manifest_import(args.manifest_url, args.trusted_base, args.app_key, args.source)
        print(json.dumps({key: value for key, value in payload.items() if key != "players"} |
                         {"playerCount": len(payload["players"]), "checksumVerified": True}, indent=2))
        if args.dry_run:
            return
    if args.command in ("validate", "preview", "stage"):
        if not KEY.fullmatch(args.app_key) or not KEY.fullmatch(args.source) or not LABEL.fullmatch(args.season) or not LABEL.fullmatch(args.round_key):
            raise CatalogError("Invalid app/source/season/round key.")
        players = read_players(args.file)
        payload = {"appKey": args.app_key, "source": args.source, "season": args.season,
                   "roundKey": args.round_key, "players": players}
        if args.command == "validate":
            print(json.dumps({"valid": True, "appKey": args.app_key, "source": args.source,
                              "season": args.season, "roundKey": args.round_key,
                              "playerCount": len(players), "teams": len({p["data"]["team"] for p in players}),
                              "roles": dict(Counter(p["data"]["role"] for p in players))}, indent=2))
            return
    client = Client(args.backend, os.environ.get("PLATFORM_ACCESS_TOKEN"))
    if not client.token:
        if not args.email:
            raise CatalogError("Set PLATFORM_ACCESS_TOKEN or supply --email for an authorized catalog operator.")
        password = os.environ.get("PLATFORM_ADMIN_PASSWORD")
        if not password:
            if args.non_interactive or not sys.stdin.isatty():
                raise CatalogError("Set PLATFORM_ADMIN_PASSWORD for non-interactive login.")
            password = getpass.getpass("Password: ")
        session = client.call("/auth/login", {"email": args.email, "password": password})
        token = session.get("session", {}).get("accessToken") if isinstance(session, dict) else None
        if not isinstance(token, str) or not token:
            raise CatalogError("Login did not return an access token.")
        client.token = token
    if args.command in ("preview", "stage", "sync-manifest"):
        result = client.call("/catalogs/import-preview", payload)
        print(json.dumps(result, indent=2))
        if args.command in ("stage", "sync-manifest"):
            result = client.call("/catalogs/import-draft", {**payload, "expectedHash": result["contentHash"]})
            print(json.dumps(result, indent=2))
    else:
        import uuid
        try:
            catalog_id = str(uuid.UUID(args.catalog_id))
        except ValueError:
            raise CatalogError("Invalid catalog UUID.") from None
        body = {"expectedHash": args.expected_hash} if args.command == "publish" else {}
        if args.command == "publish" and not re.fullmatch("[0-9a-f]{64}", args.expected_hash):
            raise CatalogError("Publication requires the SHA-256 hash returned by preview/stage.")
        print(json.dumps(client.call(f"/catalogs/{catalog_id}/{args.command}", body), indent=2))


if __name__ == "__main__":
    try:
        main()
    except (CatalogError, KeyError) as error:
        print(f"Error: {error if isinstance(error, CatalogError) else 'Invalid backend response.'}", file=sys.stderr)
        sys.exit(1)
