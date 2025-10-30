#!/usr/bin/env python3
"""
read_multiple_excels_from_folder.py

Reads all Excel (.xlsx) files from a given folder.
For each file:
  • Detects all sheet names dynamically
  • Reads each sheet (header on line 11)
  • Handles large files in memory-efficient chunks
  • Prints sheet summaries (first few rows)

Optionally extend this later to load data into PostgreSQL.

Usage:
    python3 read_multiple_excels_from_folder.py --folder /path/to/excel_files
"""

import argparse
import logging
import os
import sys
from typing import Iterator
import pandas as pd

# ------------------------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------------------------
HEADER_ROW = 10  # 0-based index => line 11 is header
CHUNK_SIZE = 10000  # rows per chunk for large file handling
DISPLAY_SAMPLE_ROWS = 10  # print first N rows per sheet
LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO").upper()

# ------------------------------------------------------------------------------
# LOGGING SETUP
# ------------------------------------------------------------------------------
logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

# ------------------------------------------------------------------------------
# CORE FUNCTIONS
# ------------------------------------------------------------------------------
def get_excel_sheet_names(file_path: str) -> list[str]:
    """Return all sheet names in the Excel file."""
    try:
        sheet_names = pd.ExcelFile(file_path, engine="openpyxl").sheet_names
        logger.info("Found %d sheets in %s: %s", len(sheet_names), os.path.basename(file_path), ", ".join(sheet_names))
        return sheet_names
    except Exception as e:
        logger.error("❌ Failed to get sheet names from %s: %s", file_path, e)
        return []


def read_excel_in_chunks(file_path: str, sheet_name: str, chunk_size: int = CHUNK_SIZE) -> Iterator[pd.DataFrame]:
    """Yield chunks of Excel data as DataFrames."""
    try:
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=HEADER_ROW,
            engine="openpyxl",
        )
        df.dropna(how="all", inplace=True)
        total_rows = len(df)
        logger.info("📄 %s → Sheet '%s' loaded with %d rows", os.path.basename(file_path), sheet_name, total_rows)

        for start in range(0, total_rows, chunk_size):
            yield df.iloc[start:start + chunk_size].copy()

    except Exception as e:
        logger.error("❌ Error reading sheet '%s' from %s: %s", sheet_name, os.path.basename(file_path), e)
        return


def process_excel_sheet(file_path: str, sheet_name: str):
    """Print sheet summary and sample data."""
    try:
        chunks = read_excel_in_chunks(file_path, sheet_name)
        first_chunk = None
        total_rows = 0

        for chunk in chunks:
            total_rows += len(chunk)
            if first_chunk is None:
                first_chunk = chunk.head(DISPLAY_SAMPLE_ROWS)

        print("\n" + "=" * 120)
        print(f"📘 FILE: {os.path.basename(file_path)}")
        print(f"📄 SHEET: {sheet_name}")
        print("=" * 120)
        print(f"Total rows (approx): {total_rows}")
        if first_chunk is not None:
            print(f"Showing first {len(first_chunk)} rows:\n")
            with pd.option_context("display.max_columns", None):
                print(first_chunk.to_string(index=False))
        else:
            print("⚠️  No data found in this sheet.")
        print("-" * 120)

    except Exception as e:
        logger.exception("Error processing sheet '%s' from %s: %s", sheet_name, file_path, e)


def process_excel_file(file_path: str):
    """Process all sheets from a single Excel file."""
    sheet_names = get_excel_sheet_names(file_path)
    if not sheet_names:
        return
    for sheet_name in sheet_names:
        process_excel_sheet(file_path, sheet_name)


def process_all_excels_in_folder(folder_path: str):
    """Process all .xlsx files in the given folder."""
    if not os.path.isdir(folder_path):
        logger.error("❌ Folder not found: %s", folder_path)
        sys.exit(1)

    excel_files = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(".xlsx")
    ]

    if not excel_files:
        logger.warning("⚠️  No Excel files found in folder: %s", folder_path)
        return

    logger.info("🔍 Found %d Excel files in folder %s", len(excel_files), folder_path)

    for file_path in excel_files:
        logger.info("➡️ Processing file: %s", os.path.basename(file_path))
        process_excel_file(file_path)

    logger.info("✅ Completed processing all Excel files in folder: %s", folder_path)


# ------------------------------------------------------------------------------
# MAIN ENTRY POINT
# ------------------------------------------------------------------------------
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Read multiple Excel files and print summaries of all sheets.")
    parser.add_argument("--folder", "-d", required=True, help="Path to folder containing Excel files")
    return parser.parse_args()


def main():
    args = parse_args()
    process_all_excels_in_folder(args.folder)


if __name__ == "__main__":
    main()
