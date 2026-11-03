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
python src/generate_record_cards.py --records-per-page=8 --input=input/SkycampExplorersDetails.xlsx --output=output/custom_output.pdf
```

- `--records-per-page`: Number of records to display per page (default: 8)
- `--input`: Custom input XLSX or CSV file path (defaults automatically to the spreadsheet in `input/`)
- `--output`: Custom output PDF file path (defaults to `output/RecordCards_YYYY-MM-DD.pdf`)

## Features

- **Automatic Input Detection**: Seamlessly processes `.xlsx` files (both single-row headers and two-row merged headers) and `.csv` files
- **Executive Health, Dietary & Medical Summary at Start**:
  - **Overview Dashboard & Statistics**: At-a-glance breakdown counts and B&W visual symbol legend
  - **Dietary Requirements Breakdown**: Grouped by diet category (Gluten-Free, Dairy-Free, Nut Restrictions, Diabetic, Vegan, Pescatarian, Vegetarian, etc.)
  - **Allergies Breakdown**: Highlights Severe/EpiPen alerts at the top, followed by Nuts, Wasp stings, Food allergies, Medication, etc.
  - **Medical Conditions Breakdown**: Highlights Critical/Chronic conditions, Asthma/Inhalers, ADHD/Autism/ASD, Physical/Mobility, Dyslexia, etc.
  - **Comprehensive Participant Signposting**: Every summary entry signposts the participant's Name, Card Number (`[#N]`), Section/Unit, and exact PDF Page Number (`p. X`)
- **Optimized for Black & White (B&W) Printing**:
  - No reliance on colour: uses high-contrast icons and badges (`[✚ EPIPEN]`, `[✚ MED]`, `[◆ DIET]`, `[★ INFO]`)
  - Dedicated grey-tinted rows (`#E0E0E0` for Medical/Allergies, `#EBEBEB` for Dietary, `#F3F3F3` for Notes) with crisp top/bottom borders that print cleanly on monochrome laser printers
  - Clean separation between Medical/Allergies and Dietary information so food restrictions are never overlooked
  - Tetanus information placed unobtrusively in personal details when no medical/allergy condition exists
- **Two-Pass Build Engine**: Automatically measures card page layout to guarantee 100% accurate page numbers in all summary tables
- **Participant Details & Contact Information**:
  - Formats ages to "X years Y months"
  - Handles inconsistent entries (e.g., "As above" emergency contacts)
  - Cleans and formats phone numbers and addresses
  - Separates participants into youth (under 18) and adult (18+) sections, sorted by section and last name
  - Keeps individual cards together to prevent splitting across pages

