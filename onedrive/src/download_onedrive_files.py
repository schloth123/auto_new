"""Download files from OneDrive share links listed in a YAML file."""

from __future__ import annotations

import argparse
import re
import shutil
from pathlib import Path
from urllib.parse import parse_qsl, unquote, urlencode, urlparse, urlunparse

import requests
import yaml


PROJECT_DIR = Path(__file__).resolve().parents[2]
WORKSPACE_DIR = Path(__file__).resolve().parents[1] / "workspace"
DOWNLOAD_TIMEOUT = (15, 120)


def add_download_parameter(url: str) -> str:
    """Ask OneDrive to return the shared file instead of a preview page."""
    parsed_url = urlparse(url)
    query = dict(parse_qsl(parsed_url.query, keep_blank_values=True))
    query["download"] = "1"
    return urlunparse(parsed_url._replace(query=urlencode(query)))


def filename_from_response(response: requests.Response, fallback: str) -> str:
    content_disposition = response.headers.get("Content-Disposition", "")
    match = re.search(r'filename\*?=(?:UTF-8\'\')?["\']?([^"\';]+)', content_disposition, re.I)
    if match:
        filename = match.group(1).strip()
    else:
        filename = Path(urlparse(response.url).path).name or fallback

    filename = unquote(Path(filename).name)
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename) or fallback


def extract_urls(raw_text: str) -> list[str]:
    urls = re.findall(r"https?://[^\s\"'<>]+", raw_text)
    cleaned_urls = []
    for url in urls:
        cleaned = url.rstrip(r".,;)]}\>")
        if cleaned:
            cleaned_urls.append(cleaned)
    return cleaned_urls


def load_links(input_path: Path) -> list[str]:
    raw_text = input_path.read_text(encoding="utf-8").strip()
    if not raw_text:
        return []

    try:
        loaded = yaml.safe_load(raw_text)
    except yaml.YAMLError:
        loaded = None

    if isinstance(loaded, list):
        if not all(isinstance(link, str) for link in loaded):
            raise ValueError("Die Eingabe-YAML muss eine Liste von OneDrive-URLs enthalten.")
        return [link.strip() for link in loaded if link.strip()]

    if isinstance(loaded, str):
        urls = extract_urls(loaded)
        if urls:
            return urls
        return [link.strip() for link in loaded.splitlines() if link.strip()]

    links: list[str] = []
    for line in raw_text.splitlines():
        cleaned = line.strip()
        if not cleaned or cleaned.startswith("#"):
            continue
        if cleaned.startswith("- "):
            cleaned = cleaned[2:].strip()
        if cleaned:
            links.append(cleaned)

    if links:
        return links

    raise ValueError("Die Eingabe muss OneDrive-URLs als YAML-Liste oder zeilenweise URLs enthalten.")


def unique_target_path(target_dir: Path, filename: str) -> Path:
    target_path = target_dir / filename
    counter = 2
    while target_path.exists():
        target_path = target_dir / f"{target_path.stem}_{counter}{target_path.suffix}"
        counter += 1
    return target_path


def download_files(links: list[str], target_dir: Path) -> list[dict[str, str]]:
    target_dir.mkdir(parents=True, exist_ok=True)
    mappings = []

    for index, original_url in enumerate(links, start=1):
        try:
            response = requests.get(
                add_download_parameter(original_url),
                allow_redirects=True,
                stream=True,
                timeout=DOWNLOAD_TIMEOUT,
                headers={"User-Agent": "onedrive-file-downloader/1.0"},
            )
            response.raise_for_status()
            filename = filename_from_response(response, f"download_{index}")
            target_path = unique_target_path(target_dir, filename)

            with target_path.open("wb") as file:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        file.write(chunk)

            mappings.append(
                {
                    "one_drive_path": original_url,
                    "target_path": str(target_path.relative_to(PROJECT_DIR)).replace("\\", "/"),
                }
            )
            print(f"Heruntergeladen: {original_url} -> {target_path}")
        except requests.RequestException as error:
            print(f"Fehler beim Herunterladen von {original_url}: {error}")

    return mappings


def write_mappings(output_path: Path, mappings: list[dict[str, str]]) -> None:
    with output_path.open("w", encoding="utf-8") as file:
        yaml.safe_dump(mappings, file, allow_unicode=False, sort_keys=False)


def reset_workspace() -> None:
    if WORKSPACE_DIR.exists():
        shutil.rmtree(WORKSPACE_DIR)
    WORKSPACE_DIR.mkdir(parents=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True, help="YAML mit OneDrive-URLs")
    parser.add_argument("--target-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True, help="Ziel-YAML für die Mappings")
    args = parser.parse_args()

    try:
        reset_workspace()
        mappings = download_files(load_links(args.input), args.target_dir)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        write_mappings(args.output, mappings)
    except (OSError, ValueError, yaml.YAMLError) as error:
        parser.error(str(error))

    print(f"{len(mappings)} Mapping(s) geschrieben nach {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())