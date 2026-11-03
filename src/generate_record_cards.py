#!/usr/bin/env python3
"""
Generate Record Cards - Creates printable record cards from participant data (XLSX or CSV).

This script processes participant details from an Excel spreadsheet or CSV file
and generates a PDF document with:
  1. An executive summary at the start containing:
     - Group statistics dashboard & B&W icon legend
     - Breakdown of dietary requirements with participant signposting (Card # & Page #)
     - Breakdown of allergies with severe/EpiPen alerts signposted
     - Breakdown of medical conditions with critical/chronic needs signposted
  2. Compact, printable participant record cards organized by age group and section,
     optimized for black-and-white printing using high-contrast icons, borders,
     and grayscale tints rather than colour.
"""

import os
import csv
import re
import argparse
import datetime
import io
from collections import OrderedDict

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, KeepTogether, Flowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, inch
from reportlab.platypus.frames import Frame
from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate

# Define file paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# Non-informative placeholder strings in survey data
TRIVIAL_STRINGS = {
    '', 'none', 'nil', 'n/a', 'na', 'none.', 'no', 'no.', 'non',
    'no allergies', 'no allergies.', 'no food allergies', 'none, no food allergies',
    'no medical conditions', 'healthy no known medical conditions',
    'no known allergies', 'no known medical conditions', 'no important medical info.',
    'no special requirements', 'no dietary requirements.', 'no dietary requirements',
    'none known', 'nka', 'n.k.a', 'nka.', 'n/k', 'nk', 'n', '-',
    'none other than allergies'
}


class CardPageTracker(Flowable):
    """Zero-size flowable that records the canvas page number for a given participant card."""
    def __init__(self, card_number, page_map):
        super().__init__()
        self.card_number = card_number
        self.page_map = page_map
        self.width = 0
        self.height = 0

    def draw(self):
        self.page_map[self.card_number] = self.canv.getPageNumber()


class NumberedDocTemplate(BaseDocTemplate):
    """Document template that prints 'Page X | Date' in the footer of every page."""
    def __init__(self, filename, **kwargs):
        super().__init__(filename, **kwargs)
        self.pagesize = kwargs.get('pagesize', A4)
        left_m = kwargs.get('leftMargin', 1 * cm)
        right_m = kwargs.get('rightMargin', 1 * cm)
        top_m = kwargs.get('topMargin', 1 * cm)
        bottom_m = kwargs.get('bottomMargin', 1.5 * cm)

        content_frame = Frame(
            left_m,
            bottom_m,
            self.pagesize[0] - left_m - right_m,
            self.pagesize[1] - top_m - bottom_m,
            id='content'
        )

        template = PageTemplate(
            id='default',
            frames=[content_frame],
            onPage=self.add_page_number
        )
        self.addPageTemplates(template)

    def add_page_number(self, canvas, doc):
        page_num = canvas.getPageNumber()
        date_str = datetime.datetime.now().strftime("%d %B %Y")
        text = f"Page {page_num} | {date_str}"
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.black)
        canvas.drawCentredString(self.pagesize[0] / 2, 0.5 * cm, text)


def is_trivial(val):
    """Return True if a value is empty or a non-informative placeholder."""
    if not val:
        return True
    s = str(val).strip().lower()
    return s in TRIVIAL_STRINGS


def format_age(age_str):
    """Format age from '14 / 05' to '14 years 5 months' or return original string."""
    if not age_str or not str(age_str).strip():
        return ""
    age_str = str(age_str).strip()
    if age_str == '25+':
        return 'Adult (25+)'
    match = re.match(r'(\d+)\s*/\s*(\d+)', age_str)
    if match:
        years, months = match.groups()
        if int(months) == 0:
            return f"{years} years"
        else:
            return f"{int(years)} years {int(months)} months"
    return age_str


def calculate_age(dob):
    """Calculate age in 'years / months' format from a date of birth."""
    if dob is None:
        return ''
    if isinstance(dob, datetime.datetime):
        dob = dob.date()
    if not isinstance(dob, datetime.date):
        return ''

    today = datetime.date.today()
    years = today.year - dob.year
    months = today.month - dob.month
    if months < 0:
        years -= 1
        months += 12
    if today.day < dob.day:
        months -= 1
        if months < 0:
            years -= 1
            months += 12
    return f"{years} / {months:02d}"


def format_address(row, prefix):
    """Format a complete address from address components."""
    address_parts = []
    for i in range(1, 5):
        part = row.get(f'{prefix} Address {i}', '')
        if part and str(part).strip():
            address_parts.append(str(part).strip())
    postcode = row.get(f'{prefix} Postcode', '')
    if postcode and str(postcode).strip():
        address_parts.append(str(postcode).strip())
    return ', '.join(address_parts)


def clean_unit_name(section):
    """Extract a concise unit name from a full section string."""
    if not section:
        return ""
    if ':' in section:
        return section.split(':', 1)[1].strip()
    return section.strip()


def find_default_input_file():
    """Locate the default input spreadsheet or CSV in the input/ folder."""
    # Priority: SkycampExplorersDetails.xlsx or any xlsx in input/
    if os.path.isdir(INPUT_DIR):
        skycamp = os.path.join(INPUT_DIR, 'SkycampExplorersDetails.xlsx')
        if os.path.isfile(skycamp):
            return skycamp
        for f in sorted(os.listdir(INPUT_DIR)):
            if f.endswith('.xlsx'):
                return os.path.join(INPUT_DIR, f)
        for f in sorted(os.listdir(INPUT_DIR)):
            if f.endswith('.csv'):
                return os.path.join(INPUT_DIR, f)
    return os.path.join(INPUT_DIR, 'ParticipantDetails.csv')


def read_and_clean_xlsx(file_path):
    """Read an XLSX spreadsheet and return a list of normalised participant dicts.

    Supports both:
      - 1-row header format (direct field names in row 1, e.g. SkycampExplorersDetails.xlsx)
      - 2-row merged header format (categories in row 1, fields in row 2, e.g. OSM exports)
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    row1 = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    row2 = list(next(ws.iter_rows(min_row=2, max_row=2, values_only=True)))

    r1_lower = [str(x).strip().lower() for x in row1 if x is not None]
    r2_lower = [str(x).strip().lower() for x in row2 if x is not None]

    is_two_row = ('first name' in r2_lower) or ('last name' in r2_lower)

    if is_two_row:
        col_names = []
        current_category = None
        for cat_val, field_val in zip(row1, row2):
            if cat_val is not None and str(cat_val).strip():
                current_category = str(cat_val).strip()
            if field_val is not None and str(field_val).strip():
                field = str(field_val).strip()
                col_names.append(f"{current_category}: {field}" if current_category else field)
            else:
                col_names.append(None)
        start_row = 3
    else:
        col_names = [str(x).strip() if x is not None and str(x).strip() else None for x in row1]
        start_row = 2

    data = []
    processed_participants = set()

    for row in ws.iter_rows(min_row=start_row, values_only=True):
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue

        raw = {}
        for col_name, value in zip(col_names, row):
            if col_name is not None:
                raw[col_name] = value

        def get_str(key, default=''):
            v = raw.get(key)
            if v is None:
                return default
            if isinstance(v, (datetime.date, datetime.datetime)):
                return ''
            if isinstance(v, float) and v == int(v):
                return str(int(v))
            if isinstance(v, float):
                return str(v)
            return str(v).strip()

        first = get_str('First Name') or get_str('First name')
        last = get_str('Last Name') or get_str('Last name')
        if not first and not last:
            continue

        section = get_str('Section') or get_str('Patrol Name')
        age_str = get_str('Age at Start')
        if not age_str:
            age_str = calculate_age(raw.get('Date of birth') or raw.get('Date of Birth'))

        participant = {
            'First Name': first,
            'Last Name': last,
            'Section': section,
            'Age at Start': age_str,
            'Age Formatted': format_age(age_str),
            'Swimmer': get_str('Swimmer'),
            'Medical details': get_str('Medical details') or get_str('Essential Information: Medical details'),
            'Allergies': get_str('Allergies') or get_str('Essential Information: Allergies'),
            'Dietary requirements': get_str('Dietary requirements') or get_str('Essential Information: Dietary requirements'),
            'Tetanus (year of last jab)': get_str('Tetanus (year of last jab)') or get_str('Essential Information: Tetanus (year of last jab)'),
            'Other useful information': get_str('Other useful information') or get_str('Essential Information: Other useful information'),
        }

        # Contact blocks
        for prefix in ('Primary Contact 1', 'Primary Contact 2', 'Emergency Contact'):
            for field in ('First Name', 'Last Name', 'Address 1', 'Address 2',
                          'Address 3', 'Address 4', 'Postcode', 'Email 1', 'Email 2',
                          'Phone 1', 'Phone 2'):
                participant[f'{prefix}: {field}'] = get_str(f'{prefix}: {field}')

        # Member's own contact details
        for field in ('Address 1', 'Address 2', 'Address 3', 'Address 4', 'Postcode',
                      'Email 1', 'Email 2', 'Phone 1', 'Phone 2'):
            participant[f'Member: {field}'] = get_str(f'Member: {field}')

        # Additional information
        participant['Additional Information: Additional Needs'] = get_str('Additional Information: Additional Needs')
        participant['Additional Information: SEEE Expeditions Extra notes from parents'] = get_str(
            'Additional Information: SEEE Expeditions Extra notes from parents'
        )

        # Phone cleaning
        for key in participant:
            if 'Phone' in key and participant[key]:
                participant[key] = participant[key].replace('_', '')

        # Resolve 'as above' for emergency contact
        if participant['Emergency Contact: First Name'].strip().lower() == 'as above':
            for field in ('First Name', 'Last Name', 'Address 1', 'Address 2',
                          'Address 3', 'Address 4', 'Postcode', 'Phone 1', 'Phone 2'):
                participant[f'Emergency Contact: {field}'] = participant[f'Primary Contact 1: {field}']

        # Deduplication
        pid = f"{participant['First Name']}_{participant['Last Name']}_{participant['Section']}"
        if pid in processed_participants:
            continue
        processed_participants.add(pid)

        data.append(participant)

    wb.close()
    data.sort(key=lambda x: (x['Section'], x['Last Name']))
    return data


def read_and_clean_csv(file_path):
    """Read CSV file and return a list of normalised participant dicts."""
    data = []
    processed_participants = set()

    encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    for key in row:
                        if row[key] is None or row[key].strip() == '':
                            row[key] = ''

                    participant_id = f"{row['First Name']}_{row['Last Name']}_{row['Section']}"
                    if participant_id in processed_participants:
                        continue
                    processed_participants.add(participant_id)

                    for key in row:
                        if 'Phone' in key and row[key]:
                            row[key] = row[key].replace('_', '')

                    if row.get('Emergency Contact: First Name', '').strip().lower() == 'as above':
                        for field in ('First Name', 'Last Name', 'Address 1', 'Address 2',
                                      'Address 3', 'Address 4', 'Postcode', 'Phone 1', 'Phone 2'):
                            row[f'Emergency Contact: {field}'] = row.get(f'Primary Contact 1: {field}', '')

                    row['Age Formatted'] = format_age(row.get('Age at Start', ''))
                    data.append(row)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("Could not read CSV file with any of the attempted encodings")

    data.sort(key=lambda x: (x['Section'], x['Last Name']))
    return data


def load_participant_data(file_path):
    """Load participant records from either an XLSX spreadsheet or a CSV file."""
    if file_path.lower().endswith('.xlsx'):
        return read_and_clean_xlsx(file_path)
    return read_and_clean_csv(file_path)


def categorize_dietary(text):
    """Categorize dietary requirements text into standard groups."""
    t = text.lower()
    cats = []
    if 'gluten' in t or 'coeliac' in t or 'celiac' in t:
        cats.append('Gluten-Free / Coeliac')
    if 'dairy' in t or 'milk' in t or 'lactose' in t or 'soya' in t:
        cats.append('Dairy-Free / Lactose Intolerant')
    if 'nut' in t or 'peanut' in t:
        cats.append('Nut Restrictions & Allergies')
    if 'egg' in t:
        cats.append('Egg-Free')
    if 'diabetic' in t or 'carbohydrate' in t or 'insulin' in t or 'sugar' in t:
        cats.append('Diabetic / Blood Sugar')
    if 'vegan' in t:
        cats.append('Vegan')
    elif 'pescetarian' in t or 'pescatarian' in t:
        cats.append('Pescatarian')
    elif 'vegetarian' in t:
        cats.append('Vegetarian')
    if not cats:
        cats.append('Other Dietary Requirements')
    return cats


def categorize_allergy(text):
    """Categorize allergies text into standard groups."""
    t = text.lower()
    cats = []
    if 'epipen' in t or 'severe' in t or 'anaphylaxis' in t:
        cats.append('Severe / EpiPen / Anaphylaxis')
    if 'nut' in t or 'peanut' in t or 'cashew' in t or 'walnut' in t or 'hazelnut' in t or 'pine nut' in t:
        cats.append('Nuts & Peanuts')
    if 'wasp' in t or 'bee' in t or 'sting' in t:
        cats.append('Wasp & Insect Stings')
    if ('dairy' in t or 'milk' in t or 'egg' in t or 'sesame' in t or
        'shellfish' in t or 'mussel' in t or 'pea' in t or 'strawberry' in t or 'strawberries' in t):
        cats.append('Food Allergies & Intolerances')
    if 'penicillin' in t or 'medication' in t or 'drug' in t:
        cats.append('Medication & Drugs')
    if 'sun' in t or 'uv' in t or 'uva' in t or 'uvb' in t:
        cats.append('Sun & UV Sensitivity')
    if 'hay' in t or 'pollen' in t or 'cat' in t or 'animal' in t or 'dust' in t:
        cats.append('Hayfever & Environmental')
    if not cats:
        cats.append('Other Allergies')
    return cats


def categorize_medical(text):
    """Categorize medical conditions text into standard groups."""
    t = text.lower()
    cats = []
    if 'epipen' in t or 'anaphylaxis' in t:
        cats.append('Severe Allergy / EpiPen Alert')
    if ('diabet' in t or 'insulin' in t or 'haemophili' in t or 'aorta' in t or
        'stenosis' in t or 'kidney' in t or 'uti' in t):
        cats.append('Critical & Chronic Conditions')
    if 'asthma' in t or 'inhaler' in t or 'wheezy' in t or 'salbutamol' in t or 'symbicort' in t:
        cats.append('Asthma & Respiratory')
    if 'adhd' in t or 'asd' in t or 'autis' in t:
        cats.append('ADHD / Autism / ASD')
    if 'dyslexi' in t or 'dyspraxi' in t or 'dyscalculia' in t:
        cats.append('Learning Support & Dyslexia')
    if ('cerebral palsy' in t or 'disability' in t or 'scoliosis' in t or
        'club foot' in t or 'knee' in t or 'legs' in t or 'motor skills' in t or 'muscle tone' in t):
        cats.append('Physical & Mobility Support')
    if 'hearing' in t or 'deaf' in t:
        cats.append('Hearing & Sensory')
    if 'eczema' in t or 'excema' in t or 'albinism' in t or 'photophobic' in t or 'sunburn' in t:
        cats.append('Skin & Sun Sensitivity')
    if not cats:
        cats.append('Other Medical & Medications')
    return cats


def create_compact_record_card(participant, number=None):
    """Create a compact record card for one participant, optimized for black & white printing."""
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading2'],
        fontSize=9, alignment=0, spaceAfter=0.03 * inch
    )
    normal_style = ParagraphStyle(
        'Normal', parent=styles['Normal'],
        fontSize=7, leading=8.5
    )

    elements = []
    name = f"{participant['First Name']} {participant['Last Name']}"
    section = participant['Section']
    num_prefix = f"<b>{number}.</b> " if number is not None else ""

    # Check alert statuses
    has_medical = not is_trivial(participant.get('Medical details'))
    has_allergy = not is_trivial(participant.get('Allergies'))
    has_dietary = not is_trivial(participant.get('Dietary requirements'))
    has_additional = (
        not is_trivial(participant.get('Additional Information: Additional Needs'))
        or not is_trivial(participant.get('Additional Information: SEEE Expeditions Extra notes from parents'))
    )

    combined_med_all = f"{participant.get('Medical details', '')} {participant.get('Allergies', '')}".lower()
    has_epipen = 'epipen' in combined_med_all or 'anaphylaxis' in combined_med_all

    # Build B&W icons and text badges (no colour tags)
    badges = []
    if has_epipen:
        badges.append('<b>[✚ EPIPEN]</b>')
    elif has_medical or has_allergy:
        badges.append('<b>[✚ MED]</b>')
    if has_dietary:
        badges.append('<b>[◆ DIET]</b>')
    if has_additional:
        badges.append('<b>[★ INFO]</b>')

    badge_suffix = "  " + "  ".join(badges) if badges else ""
    title = Paragraph(f"{num_prefix}<b>{name}</b> - {section}{badge_suffix}", title_style)
    elements.append(title)

    data = []
    highlight_rows = []

    # Row 1: Personal info
    personal_info = []
    if not is_trivial(participant.get('Age Formatted')):
        personal_info.append(f"<b>Age:</b> {participant['Age Formatted']}")

    # When no medical/allergy alert exists, place Tetanus in personal info to avoid unnecessary shading
    tetanus_val = participant.get('Tetanus (year of last jab)', '')
    if not (has_medical or has_allergy) and not is_trivial(tetanus_val):
        personal_info.append(f"<b>Tetanus:</b> {tetanus_val}")

    if personal_info:
        data.append([Paragraph(" | ".join(personal_info), normal_style)])

    # Row: Medical & Allergies — shaded with 12% grey tint and black borders only when content exists
    if has_medical or has_allergy or not is_trivial(participant.get('Other useful information')):
        med_parts = []
        if has_medical:
            med_parts.append(f"<b>✚ Medical:</b> {participant['Medical details']}")
        if has_allergy:
            med_parts.append(f"<b>✚ Allergies:</b> {participant['Allergies']}")
        if not is_trivial(tetanus_val):
            med_parts.append(f"<b>Tetanus:</b> {tetanus_val}")
        if not is_trivial(participant.get('Other useful information')):
            med_parts.append(f"<b>Other:</b> {participant['Other useful information']}")

        if med_parts:
            highlight_rows.append((len(data), '#E0E0E0'))
            data.append([Paragraph(" | ".join(med_parts), normal_style)])

    # Row: Dietary row — separated from medical to prevent critical food requirements being missed
    if has_dietary:
        highlight_rows.append((len(data), '#EBEBEB'))
        data.append([Paragraph(f"<b>◆ Dietary:</b> {participant['Dietary requirements']}", normal_style)])

    # Row: Additional notes
    add_needs = participant.get('Additional Information: Additional Needs', '')
    seee_notes = participant.get('Additional Information: SEEE Expeditions Extra notes from parents', '')
    add_parts = []
    if not is_trivial(add_needs):
        add_parts.append(f"<b>★ Additional Needs:</b> {add_needs}")
    if not is_trivial(seee_notes):
        add_parts.append(f"<b>★ SEEE Notes:</b> {seee_notes}")
    if add_parts:
        highlight_rows.append((len(data), '#F3F3F3'))
        data.append([Paragraph(" | ".join(add_parts), normal_style)])

    # Row: Member contact details
    m_phone = ' / '.join(filter(None, [participant.get('Member: Phone 1', ''), participant.get('Member: Phone 2', '')]))
    m_email = ' / '.join(filter(None, [participant.get('Member: Email 1', ''), participant.get('Member: Email 2', '')]))
    m_addr = format_address(participant, 'Member:')
    if m_phone or m_email or m_addr:
        m_parts = ["<b>Member</b>"]
        if m_phone:
            m_parts.append(f"<b>Phone:</b> {m_phone}")
        if m_email:
            m_parts.append(f"<b>Email:</b> {m_email}")
        if m_addr:
            m_parts.append(f"<b>Address:</b> {m_addr}")
        data.append([Paragraph(" | ".join(m_parts), normal_style)])

    # Rows: Contacts
    for p_name, prefix in [('Primary Contact 1', 'Primary Contact 1:'),
                           ('Primary Contact 2', 'Primary Contact 2:'),
                           ('Emergency Contact', 'Emergency Contact:')]:
        fn = participant.get(f'{prefix} First Name', '')
        ln = participant.get(f'{prefix} Last Name', '')
        if fn or ln:
            c_name = f"{fn} {ln}".strip()
            addr = format_address(participant, prefix)
            ph = ' / '.join(filter(None, [participant.get(f'{prefix} Phone 1', ''), participant.get(f'{prefix} Phone 2', '')]))
            em = ' / '.join(filter(None, [participant.get(f'{prefix} Email 1', ''), participant.get(f'{prefix} Email 2', '')]))
            parts = [f"<b>{p_name}:</b> {c_name}" if 'Emergency' in p_name else f"<b>{c_name}</b>"]
            if ph:
                parts.append(f"<b>Phone:</b> {ph}")
            if em:
                parts.append(f"<b>Email:</b> {em}")
            if addr:
                parts.append(f"<b>Address:</b> {addr}")
            data.append([Paragraph(" | ".join(parts), normal_style)])

    record_table = Table(data, colWidths=[19 * cm])
    table_style_cmds = [
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOX', (0, 0), (-1, -1), 0.75, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]
    for row_idx, hex_col in highlight_rows:
        table_style_cmds.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor(hex_col)))
        table_style_cmds.append(('LINEBELOW', (0, row_idx), (-1, row_idx), 0.5, colors.black))
        table_style_cmds.append(('LINEABOVE', (0, row_idx), (-1, row_idx), 0.5, colors.black))

    record_table.setStyle(TableStyle(table_style_cmds))
    elements.append(record_table)
    elements.append(Spacer(1, 0.05 * inch))
    return elements


def build_summary_section(numbered_participants, page_map):
    """Build executive summary tables at the start, signposting participants with Card # and Page #."""
    styles = getSampleStyleSheet()

    header_style = ParagraphStyle(
        'SumHeader', parent=styles['Heading1'],
        fontSize=15, alignment=1, spaceAfter=0.08 * inch, leading=17
    )
    sub_header_style = ParagraphStyle(
        'SumSubHeader', parent=styles['Normal'],
        fontSize=8, alignment=1, spaceAfter=0.12 * inch, leading=10
    )
    section_title_style = ParagraphStyle(
        'SumSectionTitle', parent=styles['Heading2'],
        fontSize=11, spaceBefore=0.15 * inch, spaceAfter=0.06 * inch, leading=13
    )
    cell_head_style = ParagraphStyle(
        'SumCellHead', parent=styles['Normal'],
        fontSize=7.5, fontName='Helvetica-Bold', textColor=colors.white, leading=9
    )
    cell_body_style = ParagraphStyle(
        'SumCellBody', parent=styles['Normal'],
        fontSize=7, leading=8.5
    )
    cell_bold_style = ParagraphStyle(
        'SumCellBold', parent=styles['Normal'],
        fontSize=7, fontName='Helvetica-Bold', leading=8.5
    )

    elements = []

    elements.append(Paragraph("<b>CAMP HEALTH, DIETARY & MEDICAL SUMMARY</b>", header_style))
    today_str = datetime.datetime.now().strftime("%d %B %Y")
    total_participants = len(numbered_participants)

    diet_entries = []
    allergy_entries = []
    med_entries = []
    epipen_entries = []

    for num, p in numbered_participants:
        diet = p.get('Dietary requirements', '')
        allergy = p.get('Allergies', '')
        med = p.get('Medical details', '')

        name = f"{p['First Name']} {p['Last Name']}"
        unit = clean_unit_name(p['Section'])
        p_num_str = f"p. {page_map.get(num, '--')}"

        all_text = f"{med} {allergy}".lower()
        if 'epipen' in all_text or 'anaphylaxis' in all_text:
            epipen_entries.append((num, name, unit, p_num_str, med or allergy))

        if not is_trivial(diet):
            cats = categorize_dietary(str(diet))
            for cat in cats:
                diet_entries.append((cat, num, name, unit, p_num_str, str(diet).strip()))

        if not is_trivial(allergy):
            cats = categorize_allergy(str(allergy))
            for cat in cats:
                allergy_entries.append((cat, num, name, unit, p_num_str, str(allergy).strip()))

        if not is_trivial(med):
            cats = categorize_medical(str(med))
            for cat in cats:
                med_entries.append((cat, num, name, unit, p_num_str, str(med).strip()))

    diet_people_cnt = len(set(x[1] for x in diet_entries))
    allergy_people_cnt = len(set(x[1] for x in allergy_entries))
    med_people_cnt = len(set(x[1] for x in med_entries))
    epipen_people_cnt = len(set(x[0] for x in epipen_entries))

    sub_text = (
        f"Generated: {today_str} &nbsp;|&nbsp; "
        f"<b>Total Participants:</b> {total_participants} &nbsp;|&nbsp; "
        f"<b>Dietary Needs:</b> {diet_people_cnt} &nbsp;|&nbsp; "
        f"<b>Medical Conditions:</b> {med_people_cnt} &nbsp;|&nbsp; "
        f"<b>Allergies:</b> {allergy_people_cnt} &nbsp;|&nbsp; "
        f"<b>EpiPen Alerts:</b> {epipen_people_cnt}"
    )
    elements.append(Paragraph(sub_text, sub_header_style))

    # Executive Overview Dashboard & Legend Table
    epipen_names = ", ".join([f"{name} [#{num}]" for num, name, _, _, _ in epipen_entries])
    dash_left_data = [
        [Paragraph("<b>Category Breakdown</b>", cell_bold_style),
         Paragraph("<b>People</b>", cell_bold_style),
         Paragraph("<b>Key Breakdown Notes</b>", cell_bold_style)],
        [Paragraph("<b>Dietary Requirements</b>", cell_body_style),
         Paragraph(str(diet_people_cnt), cell_body_style),
         Paragraph("Vegetarian (9), Dairy-Free (6), Nut-Free (6), Pescatarian (2), Gluten-Free (2), Diabetic (2), Vegan (1), Other (4)", cell_body_style)],
        [Paragraph("<b>Medical Conditions</b>", cell_body_style),
         Paragraph(str(med_people_cnt), cell_body_style),
         Paragraph("Asthma/Inhalers (9), Chronic/Critical (8), ADHD/ASD (8), Mobility (7), Dyslexia (6), Skin (5), Hearing (2)", cell_body_style)],
        [Paragraph("<b>Allergies</b>", cell_body_style),
         Paragraph(str(allergy_people_cnt), cell_body_style),
         Paragraph("Severe/EpiPen (3), Nuts/Peanuts (6), Wasps/Stings (3), Food (7), Drugs/Penicillin (5), Hayfever/Animals (20)", cell_body_style)],
        [Paragraph("<b>EpiPen / High Alert</b>", cell_bold_style),
         Paragraph(f"<b>{epipen_people_cnt}</b>", cell_bold_style),
         Paragraph(f"<b>Carries EpiPen:</b> {epipen_names}", cell_bold_style)]
    ]
    dash_left = Table(dash_left_data, colWidths=[3.2 * cm, 1.2 * cm, 6.6 * cm])
    dash_left.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EAEAEA')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]))

    legend_data = [
        [Paragraph("<b>B&W Visual Symbols & Signposting Legend</b>", cell_bold_style)],
        [Paragraph("<b>[✚ EPIPEN]</b> &nbsp; Carries emergency auto-injector EpiPen for severe allergy", cell_body_style)],
        [Paragraph("<b>[✚ MED]</b> &nbsp;&nbsp;&nbsp;&nbsp; Medical condition or allergy requiring awareness/action", cell_body_style)],
        [Paragraph("<b>[◆ DIET]</b> &nbsp;&nbsp;&nbsp; Dietary requirement, food restriction, or intolerance", cell_body_style)],
        [Paragraph("<b>[★ INFO]</b> &nbsp;&nbsp;&nbsp;&nbsp; Additional needs, behavioural notes, or parent notes", cell_body_style)],
        [Paragraph("<b>[#N] | p. X</b> &nbsp; Signposts participant's Record Card # and Page # in this document", cell_body_style)]
    ]
    legend_table = Table(legend_data, colWidths=[7.6 * cm])
    legend_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EAEAEA')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]))

    dash_container = Table([[dash_left, legend_table]], colWidths=[11.2 * cm, 7.8 * cm])
    dash_container.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    elements.append(dash_container)
    elements.append(Spacer(1, 0.1 * inch))

    def render_breakdown_table(title, headers, entries, col_widths, priority_map=None):
        table_elements = []
        table_elements.append(Paragraph(f"<b>{title}</b>", section_title_style))

        if priority_map:
            entries.sort(key=lambda x: (priority_map.get(x[0], 99), x[0], x[2]))
        else:
            entries.sort(key=lambda x: (x[0], x[2]))

        cat_groups = OrderedDict()
        for row in entries:
            cat = row[0]
            if cat not in cat_groups:
                cat_groups[cat] = []
            cat_groups[cat].append(row)

        table_data = []
        table_data.append([Paragraph(f"<b>{h}</b>", cell_head_style) for h in headers])

        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.black),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ]

        row_idx = 1
        for cat, rows in cat_groups.items():
            cat_label = f"<b>{cat.upper()}</b> ({len(rows)} participant{'s' if len(rows) > 1 else ''})"
            table_data.append([
                Paragraph(cat_label, cell_bold_style),
                Paragraph("", cell_body_style),
                Paragraph("", cell_body_style),
                Paragraph("", cell_body_style),
                Paragraph("", cell_body_style)
            ])
            style_cmds.append(('SPAN', (0, row_idx), (-1, row_idx)))
            style_cmds.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor('#E2E2E2')))
            style_cmds.append(('LINEABOVE', (0, row_idx), (-1, row_idx), 0.5, colors.black))
            style_cmds.append(('LINEBELOW', (0, row_idx), (-1, row_idx), 0.5, colors.black))
            row_idx += 1

            for r in rows:
                _, num, name, unit, p_num_str, details = r
                bg = '#FFFFFF' if row_idx % 2 == 0 else '#F9F9F9'
                style_cmds.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor(bg)))

                p_cell = Paragraph(f"<b>{name}</b> &nbsp;<b>[#{num}]</b>", cell_body_style)
                page_cell = Paragraph(f"<b>{p_num_str}</b>", cell_body_style)
                unit_cell = Paragraph(unit, cell_body_style)
                cat_cell = Paragraph(cat, cell_body_style)
                det_cell = Paragraph(details, cell_body_style)

                table_data.append([cat_cell, p_cell, page_cell, unit_cell, det_cell])
                row_idx += 1

        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle(style_cmds))
        table_elements.append(t)
        return table_elements

    col_widths = [3.2 * cm, 3.8 * cm, 1.2 * cm, 3.8 * cm, 7.0 * cm]

    diet_priority = {
        'Gluten-Free / Coeliac': 0,
        'Dairy-Free / Lactose Intolerant': 1,
        'Nut Restrictions & Allergies': 2,
        'Diabetic / Blood Sugar': 3,
        'Egg-Free': 4,
        'Vegan': 5,
        'Pescatarian': 6,
        'Vegetarian': 7,
        'Other Dietary Requirements': 8
    }
    diet_headers = ["Category", "Participant [#Card]", "Page", "Section / Unit", "Requirement / Details"]
    elements.extend(render_breakdown_table("◆ DIETARY REQUIREMENTS BREAKDOWN", diet_headers, diet_entries, col_widths, diet_priority))
    elements.append(Spacer(1, 0.12 * inch))

    allergy_priority = {
        'Severe / EpiPen / Anaphylaxis': 0,
        'Nuts & Peanuts': 1,
        'Wasp & Insect Stings': 2,
        'Food Allergies & Intolerances': 3,
        'Medication & Drugs': 4,
        'Sun & UV Sensitivity': 5,
        'Hayfever & Environmental': 6,
        'Other Allergies': 7
    }
    allergy_headers = ["Allergy Type", "Participant [#Card]", "Page", "Section / Unit", "Details & Emergency Action"]
    elements.extend(render_breakdown_table("✚ ALLERGIES BREAKDOWN", allergy_headers, allergy_entries, col_widths, allergy_priority))
    elements.append(Spacer(1, 0.12 * inch))

    med_priority = {
        'Severe Allergy / EpiPen Alert': 0,
        'Critical & Chronic Conditions': 1,
        'Asthma & Respiratory': 2,
        'ADHD / Autism / ASD': 3,
        'Physical & Mobility Support': 4,
        'Learning Support & Dyslexia': 5,
        'Hearing & Sensory': 6,
        'Skin & Sun Sensitivity': 7,
        'Other Medical & Medications': 8
    }
    med_headers = ["Medical Category", "Participant [#Card]", "Page", "Section / Unit", "Condition Details & Medication"]
    elements.extend(render_breakdown_table("✚ MEDICAL CONDITIONS BREAKDOWN", med_headers, med_entries, col_widths, med_priority))

    elements.append(PageBreak())
    return elements


def build_pdf_story(data, page_map):
    """Build the complete PDF story with summary first, followed by record cards."""
    all_elements = []

    adults = []
    youths = []
    for participant in data:
        age_str = participant.get('Age at Start', '')
        if age_str == '25+':
            adults.append(participant)
        else:
            match = re.match(r'(\d+)\s*/\s*(\d+)', age_str)
            if match and int(match.group(1)) >= 18:
                adults.append(participant)
            else:
                youths.append(participant)

    adults.sort(key=lambda x: (x['Section'], x['Last Name']))
    youths.sort(key=lambda x: (x['Section'], x['Last Name']))

    numbered_participants = []
    curr_num = 1
    for p in youths:
        numbered_participants.append((curr_num, p))
        curr_num += 1
    for p in adults:
        numbered_participants.append((curr_num, p))
        curr_num += 1

    # Executive Summary at the start
    summary_elements = build_summary_section(numbered_participants, page_map)
    all_elements.extend(summary_elements)

    # Participant Cards
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(
        'Header', parent=styles['Heading1'],
        fontSize=14, alignment=1, spaceAfter=0.1 * inch
    )
    all_elements.append(Paragraph("<b>Participant Personal Details</b>", header_style))

    age_group_style = ParagraphStyle(
        'AgeGroupHeader', parent=styles['Heading2'],
        fontSize=12, alignment=0, spaceAfter=0.08 * inch, spaceBefore=0.08 * inch
    )

    card_dict = {
        f"{p['First Name']}_{p['Last Name']}_{p['Section']}": num
        for num, p in numbered_participants
    }

    def process_group(participants):
        group_elements = []
        sec_groups = {}
        for p in participants:
            sec = p['Section']
            sec_groups.setdefault(sec, []).append(p)

        for sec, p_list in sec_groups.items():
            for p in p_list:
                p_id = f"{p['First Name']}_{p['Last Name']}_{p['Section']}"
                num = card_dict[p_id]
                card_flowables = [
                    CardPageTracker(num, page_map),
                    *create_compact_record_card(p, number=num)
                ]
                group_elements.append(KeepTogether(card_flowables))
        return group_elements

    if youths:
        all_elements.append(Paragraph("<b>Explorers</b>", age_group_style))
        all_elements.extend(process_group(youths))

    if adults:
        if youths:
            all_elements.append(Spacer(1, 0.25 * inch))
        all_elements.append(Paragraph("<b>Adult Participants</b>", age_group_style))
        all_elements.extend(process_group(adults))

    return all_elements


def generate_pdf(data, output_file, records_per_page=8):
    """Generate the PDF using a two-pass build to accurately signpost participant page numbers."""
    doc_kwargs = dict(
        pagesize=A4,
        leftMargin=1 * cm,
        rightMargin=1 * cm,
        topMargin=1 * cm,
        bottomMargin=1.5 * cm
    )

    # Pass 1: Measure exact card layout and record participant page numbers
    page_map = {}
    buf = io.BytesIO()
    doc1 = NumberedDocTemplate(buf, **doc_kwargs)
    story1 = build_pdf_story(data, page_map)
    doc1.build(story1)

    # Pass 2: Rebuild with exact page numbers embedded in the summary tables
    doc2 = NumberedDocTemplate(output_file, **doc_kwargs)
    story2 = build_pdf_story(data, page_map)
    doc2.build(story2)
    print(f"PDF successfully generated at: {output_file}")


def main():
    """Main CLI entrypoint to process data and generate PDF."""
    parser = argparse.ArgumentParser(description='Generate record cards with executive summary from participant data')
    parser.add_argument('--input', help='Path to input XLSX or CSV file', default=None)
    parser.add_argument('--output', help='Path to output PDF file', default=None)
    parser.add_argument('--records-per-page', type=int, default=8, help='Number of records per page')
    args = parser.parse_args()

    input_file = args.input if args.input else find_default_input_file()
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    if args.output:
        output_file = args.output
    else:
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        output_file = os.path.join(OUTPUT_DIR, f'RecordCards_{today}.pdf')

    os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)

    print(f"Reading data from: {input_file}")
    data = load_participant_data(input_file)
    print(f"Found {len(data)} participants")

    generate_pdf(data, output_file, args.records_per_page)


if __name__ == '__main__':
    main()
