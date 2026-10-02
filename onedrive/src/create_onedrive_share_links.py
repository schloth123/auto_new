"""Create anonymous OneDrive sharing links for files in a synced local folder."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from urllib.parse import quote

import msal
import requests
import yaml


GRAPH_API_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["Files.ReadWrite"]
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
REQUEST_TIMEOUT = (15, 60)
ONEDRIVE_ROOT_ENV_VARS = ("OneDriveCommercial", "OneDriveConsumer", "OneDrive")
DEFAULT_OUTPUT_PATH = Path(__file__).resolve().parents[1] / "workspace" / "onedrive_share_paths.yaml"


def find_onedrive_root(folder: Path, configured_root: Path | None) -> Path:
    """Find the local sync root whose relative paths match the user's OneDrive drive."""
    folder = folder.resolve()
    candidate_roots = [configured_root] if configured_root else [
        Path(os.environ[name])
        for name in ONEDRIVE_ROOT_ENV_VARS
        if os.environ.get(name)
    ]

    for candidate in candidate_roots:
        root = candidate.resolve()
        if folder == root or root in folder.parents:
            return root

    raise ValueError(
        "Der Ordner liegt nicht unter einem erkannten OneDrive-Synchronisierungsstamm. "
        "Setze --onedrive-root auf den lokalen Stamm deines OneDrive-Laufwerks."
    )


def acquire_access_token(client_id: str) -> str:
    """Authenticate interactively using Microsoft's device-code flow."""
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority="https://login.microsoftonline.com/common",
    )
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPE)
    if "user_code" not in flow:
        raise RuntimeError(f"Microsoft-Anmeldung konnte nicht gestartet werden: {flow}")

    print(flow["message"], file=sys.stderr)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        message = result.get("error_description", result.get("error", "Unbekannter Anmeldefehler"))
        raise RuntimeError(f"Microsoft-Anmeldung fehlgeschlagen: {message}")
    return result["access_token"]


def create_share_link(
    session: requests.Session,
    drive_relative_path: str,
    permission: str,
) -> str:
    encoded_path = quote(drive_relative_path, safe="/")
    url = f"{GRAPH_API_URL}/me/drive/root:/{encoded_path}:/createLink"
    response = session.post(
        url,
        json={"type": permission, "scope": "anonymous"},
        timeout=REQUEST_TIMEOUT,
    )
    response.raise_for_status()
    return response.json()["link"]["webUrl"]


def create_links(folder: Path, root: Path, permission: str, access_token: str) -> list[dict[str, str]]:
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {access_token}"})
    links = []

    for file_path in sorted(path for path in folder.rglob("*") if path.is_file()):
        relative_path = file_path.relative_to(root).as_posix()
        try:
            share_url = create_share_link(session, relative_path, permission)
        except (requests.RequestException, KeyError, ValueError) as error:
            response = getattr(error, "response", None)
            details = response.text if response is not None else str(error)
            print(f"Fehler für {file_path}: {details}", file=sys.stderr)
            continue

        links.append({"file_name": file_path.name, "one_drive_share_path": share_url})

    return links


def write_links(output_path: Path, links: list[dict[str, str]]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as output_file:
        yaml.safe_dump(links, output_file, allow_unicode=True, sort_keys=False)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--folder", type=Path, required=True, help="Lokaler Pfad zum OneDrive-Ordner")
    parser.add_argument(
        "--permission",
        choices=("view", "edit"),
        required=True,
        help="Link-Berechtigung: view (kann anzeigen) oder edit (kann bearbeiten)",
    )
    parser.add_argument(
        "--client-id",
        default=os.environ.get("ONEDRIVE_CLIENT_ID", DEFAULT_CLIENT_ID),
        help="Optionale eigene Client-ID; Standard ist die Microsoft-Graph-PowerShell-App",
    )
    parser.add_argument(
        "--onedrive-root",
        type=Path,
        help="Lokaler OneDrive-Synchronisierungsstamm; normalerweise automatisch erkannt",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help=f"Ziel-YAML (Standard: {DEFAULT_OUTPUT_PATH})",
    )
    args = parser.parse_args()

    if not args.folder.is_dir():
        parser.error(f"Ordner nicht gefunden: {args.folder}")

    try:
        root = find_onedrive_root(args.folder, args.onedrive_root)
        access_token = acquire_access_token(args.client_id)
        links = create_links(args.folder.resolve(), root, args.permission, access_token)
        write_links(args.output, links)
    except (OSError, RuntimeError, ValueError) as error:
        parser.error(str(error))

    print(f"{len(links)} Freigabelink(s) geschrieben nach {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())