#!/usr/bin/env python3
import argparse
import os
import re
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import openpyxl
import requests

DEFAULT_INPUT = "products_with_sku_and_name (2).xlsx"
USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)


def load_rows(path: Path) -> Tuple[List[str], Iterable[Tuple]]:
    workbook = openpyxl.load_workbook(path, read_only=True)
    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)
    headers = [str(value).strip() if value is not None else "" for value in next(rows)]
    return headers, rows


def find_column(headers: List[str], candidates: List[str]) -> Optional[int]:
    lowered = [header.lower() for header in headers]
    for candidate in candidates:
        if candidate in lowered:
            return lowered.index(candidate)
    for index, header in enumerate(lowered):
        for candidate in candidates:
            if candidate in header:
                return index
    return None


def get_vqd(query: str, session: requests.Session) -> str:
    response = session.get(
        "https://duckduckgo.com/",
        params={"q": query},
        headers={"User-Agent": USER_AGENT},
        timeout=30,
    )
    response.raise_for_status()
    match = re.search(r"vqd='([^']+)'", response.text)
    if match:
        return match.group(1)
    match = re.search(r"vqd=([^&]+)&", response.text)
    if match:
        return match.group(1)
    raise RuntimeError("Could not find vqd token for search query.")


def fetch_image_results(query: str, max_results: int, session: requests.Session) -> List[Dict]:
    vqd = get_vqd(query, session)
    params = {
        "l": "us-en",
        "o": "json",
        "q": query,
        "vqd": vqd,
        "f": ",,,",
        "p": "1",
    }
    response = session.get(
        "https://duckduckgo.com/i.js",
        params=params,
        headers={"User-Agent": USER_AGENT},
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    results = data.get("results", [])
    return results[:max_results]


def safe_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name.strip())
    return name.strip("_") or "image"


def guess_extension(url: str, content_type: Optional[str]) -> str:
    if content_type:
        if "jpeg" in content_type:
            return ".jpg"
        if "png" in content_type:
            return ".png"
        if "webp" in content_type:
            return ".webp"
        if "gif" in content_type:
            return ".gif"
    match = re.search(r"\.(jpg|jpeg|png|gif|webp)(?:$|\?)", url, re.IGNORECASE)
    if match:
        ext = match.group(1).lower()
        return ".jpg" if ext == "jpeg" else f".{ext}"
    return ".jpg"


def download_image(url: str, destination: Path, session: requests.Session) -> None:
    response = session.get(url, headers={"User-Agent": USER_AGENT}, timeout=60)
    response.raise_for_status()
    extension = guess_extension(url, response.headers.get("Content-Type"))
    output_path = destination.with_suffix(extension)
    output_path.write_bytes(response.content)
    print(f"Saved: {output_path}")


def prompt_yes_no(prompt: str) -> bool:
    while True:
        answer = input(prompt).strip().lower()
        if answer in {"y", "yes"}:
            return True
        if answer in {"n", "no"}:
            return False
        print("Please answer 'y' or 'n'.")


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Search for product images and download the accepted result as the SKU filename."
        )
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path(DEFAULT_INPUT),
        help="Path to the Excel spreadsheet in the repo.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("downloaded_images"),
        help="Directory to save downloaded images.",
    )
    parser.add_argument(
        "--max-results",
        type=int,
        default=5,
        help="Number of top image results to review per product.",
    )
    args = parser.parse_args()

    if not args.input.exists():
        print(f"Input file not found: {args.input}", file=sys.stderr)
        return 1

    headers, rows = load_rows(args.input)
    sku_index = find_column(headers, ["sku", "item sku", "item_sku"])
    name_index = find_column(headers, ["name", "product name", "product_name", "title"])

    if sku_index is None or name_index is None:
        print(
            "Could not find required columns. Ensure the spreadsheet has SKU and name columns.",
            file=sys.stderr,
        )
        print(f"Headers found: {headers}", file=sys.stderr)
        return 1

    args.output_dir.mkdir(parents=True, exist_ok=True)

    session = requests.Session()

    for row_number, row in enumerate(rows, start=2):
        sku_value = row[sku_index] if sku_index < len(row) else None
        name_value = row[name_index] if name_index < len(row) else None
        if not sku_value or not name_value:
            print(f"Skipping row {row_number}: missing SKU or name.")
            continue

        sku = str(sku_value).strip()
        name = str(name_value).strip()
        query = f"{name}"
        print(f"\nSearching images for SKU {sku}: {name}")

        try:
            results = fetch_image_results(query, args.max_results, session)
        except Exception as exc:
            print(f"Failed to fetch results for {sku}: {exc}")
            continue

        if not results:
            print("No image results found.")
            continue

        accepted = False
        for index, result in enumerate(results, start=1):
            url = result.get("image")
            source = result.get("title") or result.get("url") or ""
            if not url:
                continue
            print(f"[{index}/{len(results)}] {url}")
            if source:
                print(f"    Source: {source}")
            if prompt_yes_no("Download this image? (y/n): "):
                filename = safe_filename(sku)
                destination = args.output_dir / filename
                try:
                    download_image(url, destination, session)
                except Exception as exc:
                    print(f"Failed to download image: {exc}")
                accepted = True
                break

        if not accepted:
            print("No image selected for this SKU.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
