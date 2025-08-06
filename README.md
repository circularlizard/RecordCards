# Record Cards Generator

This tool processes participant data from a CSV file and generates printable record cards as a PDF, with multiple participants per page, organized by age group.

## Requirements

- Python 3.6 or higher
- Dependencies listed in `requirements.txt`

## Directory Structure

```text
RecordCards/
├── input/
│   └── ParticipantDetails.csv  # Input participant data
├── output/
│   └── RecordCards_YYYY-MM-DD.pdf  # Generated output with date stamp
├── src/
│   └── generate_record_cards.py # Main script
└── requirements.txt
```

## Setup

1. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Ensure your participant data is saved as `input/ParticipantDetails.csv`
2. Run the script:

```bash
python src/generate_record_cards.py
```

The generated PDF will be available at `output/RecordCards_YYYY-MM-DD.pdf`

### Command-line Arguments

The script supports several command-line arguments:

```bash
python src/generate_record_cards.py --records-per-page=8 --input=custom_input.csv --output=custom_output.pdf
```

- `--records-per-page`: Number of records to display per page (default: 8)
- `--input`: Custom input CSV file path
- `--output`: Custom output PDF file path

## Features

- Processes and standardizes participant data
- Handles inconsistent entries (e.g., "As above" for emergency contacts)
- Formats ages from "14 / 05" notation to "14 years 5 months"
- **Separates participants into youth (under 18) and adult (18+) sections**
- Sorts participants by section and last name within each age group
- Creates compact record cards with multiple entries per page
- Uses proper spacing and borders to prevent text collisions
- **Ensures records don't split between pages**
- **Optimizes spacing to avoid large gaps at page boundaries**
- **Excludes non-essential fields like school and religion for more compact presentation**
- Includes:
  - Personal details
  - Medical information
  - Contact information for primary and emergency contacts

## Customization

You can modify the `src/generate_record_cards.py` script to:

- Change the formatting of the record cards
- Add or remove fields
- Adjust the page layout and styling
