#!/usr/bin/env python3
"""
Generate Record Cards from XLSX - Creates printable record cards from participant XLSX data.

Reads an XLSX file (supporting either direct headers or two-row merged headers),
normalises the data into the structure expected by the PDF generation pipeline,
and produces the printable record cards with an executive health and dietary summary.
"""

import os
import sys
import argparse
import datetime

# Allow importing shared PDF-generation functions from the sibling module
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_record_cards import (  # noqa: E402
    format_age,
    calculate_age,
    format_address,
    create_compact_record_card,
    read_and_clean_xlsx,
    find_default_input_file,
    generate_pdf,
    INPUT_DIR,
    OUTPUT_DIR,
)

_today = datetime.datetime.now().strftime("%Y-%m-%d")
DEFAULT_OUTPUT_FILE = os.path.join(OUTPUT_DIR, f'RecordCards_{_today}.pdf')


def read_xlsx(file_path):
    """Read the XLSX file and return normalised participant records."""
    return read_and_clean_xlsx(file_path)


def main():
    parser = argparse.ArgumentParser(
        description='Generate record cards from participant XLSX data'
    )
    parser.add_argument('--input', help='Path to input XLSX file', default=None)
    parser.add_argument('--output', help='Path to output PDF file', default=None)
    parser.add_argument('--records-per-page', type=int, default=8,
                        help='Number of records per page')
    args = parser.parse_args()

    input_file = args.input if args.input else find_default_input_file()
    output_file = args.output if args.output else DEFAULT_OUTPUT_FILE

    os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)

    print(f"Reading data from: {input_file}")
    data = read_and_clean_xlsx(input_file)
    print(f"Found {len(data)} participants")

    generate_pdf(data, output_file, args.records_per_page)


if __name__ == '__main__':
    main()
