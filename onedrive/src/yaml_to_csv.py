"""Convert OneDrive share-path entries from YAML to CSV."""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

import yaml


WORKSPACE_DIR = Path(__file__).resolve().parents[1] / "workspace"
CSV_COLUMNS = ("file_name", "one_drive_share_path")


def convert_yaml_to_csv(input_path: Path) -> Path:
    with input_path.open(encoding="utf-8") as input_file:
        entries = yaml.safe_load(input_file)

    if entries is None:
        entries = []
    if not isinstance(entries, list) or not all(isinstance(entry, dict) for entry in entries):
        raise ValueError("Die YAML-Datei muss eine Liste von Einträgen enthalten.")

    output_path = WORKSPACE_DIR / f"{input_path.stem}.csv"
    with output_path.open("w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.DictWriter(output_file, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(entries)

    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("yaml_path", type=Path, help="Pfad zur Eingabe-YAML-Datei")
    args = parser.parse_args()

    if not args.yaml_path.is_file():
        parser.error(f"Datei nicht gefunden: {args.yaml_path}")

    try:
        output_path = convert_yaml_to_csv(args.yaml_path)
    except (OSError, ValueError, yaml.YAMLError) as error:
        parser.error(str(error))

    print(f"CSV geschrieben nach {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())