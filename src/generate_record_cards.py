#!/usr/bin/env python3
"""
Generate Record Cards - Creates printable record cards from participant CSV data

This script processes a CSV file containing participant details and generates
a PDF document with multiple participants per page, formatted as compact record cards.
"""

import os
import csv
import re
import argparse
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus.frames import Frame
from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate

# Define file paths
INPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'input')
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'output')
INPUT_FILE = os.path.join(INPUT_DIR, 'ParticipantDetails.csv')
today = datetime.datetime.now().strftime("%Y-%m-%d")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f'RecordCards_{today}.pdf')

# Custom document template class with page numbering
class NumberedDocTemplate(BaseDocTemplate):
    def __init__(self, filename, **kwargs):
        BaseDocTemplate.__init__(self, filename, **kwargs)
        self.pagesize = kwargs.get('pagesize', A4)
        
        # Create a frame for the content
        content_frame = Frame(
            kwargs.get('leftMargin', 1*cm),
            kwargs.get('bottomMargin', 1*cm),
            self.pagesize[0] - kwargs.get('leftMargin', 1*cm) - kwargs.get('rightMargin', 1*cm),
            self.pagesize[1] - kwargs.get('topMargin', 1*cm) - kwargs.get('bottomMargin', 1*cm),
            id='content'
        )
        
        # Create a page template with the content frame and add page numbers
        template = PageTemplate(
            id='default',
            frames=[content_frame],
            onPage=self.add_page_number
        )
        
        self.addPageTemplates(template)
    
    def add_page_number(self, canvas, doc):
        # Add page number and date at the bottom of each page
        page_num = canvas.getPageNumber()
        date_str = datetime.datetime.now().strftime("%d %B %Y")
        text = f"Page {page_num} | {date_str}"
        
        # Set up text appearance
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.black)
        
        # Position the text at the bottom center
        canvas.drawCentredString(
            self.pagesize[0] / 2,
            0.5 * cm,  # 0.5 cm from bottom
            text
        )

def read_and_clean_csv(file_path):
    """Read CSV file and clean the data"""
    data = []
    processed_participants = set()  # Track processed participants to avoid duplicates
    
    # Try different encodings
    encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as csvfile:
                reader = csv.DictReader(csvfile)
                headers = reader.fieldnames
                
                for row in reader:
                    # Replace empty values with an empty string
                    for key in row:
                        if row[key] is None or row[key].strip() == '':
                            row[key] = ''
                    
                    # Create a unique identifier for the participant
                    participant_id = f"{row['First Name']}_{row['Last Name']}_{row['Section']}"
                    
                    # Skip if we've already processed this participant
                    if participant_id in processed_participants:
                        continue
                    
                    processed_participants.add(participant_id)
                    
                    # Clean phone numbers (remove underscores used for formatting)
                    for key in row:
                        if 'Phone' in key and row[key]:
                            row[key] = row[key].replace('_', '')
                    
                    # Process 'As above' entries for emergency contacts
                    if row['Emergency Contact: First Name'].strip().lower() == 'as above':
                        # Copy primary contact details to emergency contact
                        row['Emergency Contact: First Name'] = row['Primary Contact 1: First Name']
                        row['Emergency Contact: Last Name'] = row['Primary Contact 1: Last Name']
                        row['Emergency Contact: Address 1'] = row['Primary Contact 1: Address 1']
                        row['Emergency Contact: Address 2'] = row['Primary Contact 1: Address 2']
                        row['Emergency Contact: Address 3'] = row['Primary Contact 1: Address 3']
                        row['Emergency Contact: Address 4'] = row['Primary Contact 1: Address 4']
                        row['Emergency Contact: Postcode'] = row['Primary Contact 1: Postcode']
                        row['Emergency Contact: Phone 1'] = row['Primary Contact 1: Phone 1']
                        row['Emergency Contact: Phone 2'] = row['Primary Contact 1: Phone 2']
                    
                    # Format age
                    row['Age Formatted'] = format_age(row['Age at Start'])
                    
                    data.append(row)
            
            print(f"Successfully read file with encoding: {encoding}")
            break  # Break the loop if successful
        except UnicodeDecodeError:
            print(f"Failed to read with encoding: {encoding}")
            continue
    else:
        raise ValueError("Could not read CSV file with any of the attempted encodings")
    
    # Sort by section and then by last name
    data.sort(key=lambda x: (x['Section'], x['Last Name']))
    
    return data

def format_age(age_str):
    """Format age from '14 / 05' to '14 years 5 months' or return original if not in expected format"""
    if not age_str or not age_str.strip():
        return ""
    
    if age_str == '25+':
        return 'Adult (25+)'
    
    match = re.match(r'(\d+)\s*/\s*(\d+)', age_str)
    if match:
        years, months = match.groups()
        if int(months) == 0:
            return f"{years} years"
        else:
            return f"{years} years {months} months"
    return age_str

def format_address(row, prefix):
    """Format a complete address from address components"""
    address_parts = []
    for i in range(1, 5):
        part = row.get(f'{prefix} Address {i}', '')
        if part and part.strip():
            address_parts.append(part.strip())
    
    postcode = row.get(f'{prefix} Postcode', '')
    if postcode and postcode.strip():
        address_parts.append(postcode.strip())
    
    # For compact display, join with commas instead of newlines
    return ', '.join(address_parts)

def is_trivial(val):
    """Return True if a value is empty or a non-informative placeholder such as 'none', 'nil' or 'n/a'."""
    if not val:
        return True
    return str(val).strip().lower() in ('', 'none', 'nil', 'n/a', 'na')


def create_compact_record_card(participant, number=None):
    """Create a compact record card for one participant"""
    styles = getSampleStyleSheet()
    
    # Create custom styles
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading2'],
        fontSize=9,
        alignment=0,  # Left alignment
        spaceAfter=0.03*inch
    )
    
    normal_style = ParagraphStyle(
        'Normal',
        parent=styles['Normal'],
        fontSize=7,
        leading=8  # Reduce line spacing
    )
    
    # Content elements
    elements = []
    
    # Title - Participant name and section
    name = f"{participant['First Name']} {participant['Last Name']}"
    section = participant['Section']
    num_prefix = f"<b>{number}.</b> " if number is not None else ""

    # Build alert badges for the title line
    badges = []
    has_medical = (
        not is_trivial(participant.get('Medical details'))
        or not is_trivial(participant.get('Allergies'))
    )
    has_dietary = not is_trivial(participant.get('Dietary requirements'))
    has_additional = (
        not is_trivial(participant.get('Additional Information: Additional Needs'))
        or not is_trivial(participant.get('Additional Information: SEEE Expeditions Extra notes from parents'))
    )
    if has_medical:
        badges.append('<font color="#cc0000"><b>[MED]</b></font>')
    if has_dietary:
        badges.append('<font color="#cc6600"><b>[DIET]</b></font>')
    if has_additional:
        badges.append('<font color="#0055cc"><b>[INFO]</b></font>')
    badge_suffix = "  " + "  ".join(badges) if badges else ""

    title = Paragraph(f"{num_prefix}<b>{name}</b> - {section}{badge_suffix}", title_style)
    elements.append(title)
    
    # Create a table for the rest of the information
    data = []
    highlight_rows = []  # (row_index, hex_colour) — rows to be given a background colour

    # Row 1: Personal details
    personal_info = []
    if not is_trivial(participant.get('Age Formatted')):
        personal_info.append(f"<b>Age:</b> {participant['Age Formatted']}")
    if personal_info:
        data.append([Paragraph(" | ".join(personal_info), normal_style)])
    
    # Row: Medical information — highlighted when any real content is present
    medical_info = []
    if not is_trivial(participant['Medical details']):
        medical_info.append(f"<b>Medical:</b> {participant['Medical details']}")
    if not is_trivial(participant['Allergies']):
        medical_info.append(f"<b>Allergies:</b> {participant['Allergies']}")
    if not is_trivial(participant['Dietary requirements']):
        medical_info.append(f"<b>Diet:</b> {participant['Dietary requirements']}")
    if not is_trivial(participant['Tetanus (year of last jab)']):
        medical_info.append(f"<b>Tetanus:</b> {participant['Tetanus (year of last jab)']}")
    if not is_trivial(participant['Other useful information']):
        medical_info.append(f"<b>Other:</b> {participant['Other useful information']}")

    if medical_info:
        highlight_rows.append((len(data), '#FFE8E8'))
        medical_info_text = Paragraph(" | ".join(medical_info), normal_style)
        data.append([medical_info_text])

    # Row: Additional information — optional; from XLSX 'Additional Information' columns (BL–BM)
    add_needs = participant.get('Additional Information: Additional Needs', '')
    seee_notes = participant.get('Additional Information: SEEE Expeditions Extra notes from parents', '')
    additional_info = []
    if not is_trivial(add_needs):
        additional_info.append(f"<b>Additional Needs:</b> {add_needs}")
    if not is_trivial(seee_notes):
        additional_info.append(f"<b>SEEE Notes:</b> {seee_notes}")
    if additional_info:
        highlight_rows.append((len(data), '#FFF8DC'))
        data.append([Paragraph(" | ".join(additional_info), normal_style)])

    # Row: Member's own contact details — optional; from XLSX 'Member' columns (AW–BE)
    member_phone1 = participant.get('Member: Phone 1', '')
    member_phone2 = participant.get('Member: Phone 2', '')
    member_email1 = participant.get('Member: Email 1', '')
    member_email2 = participant.get('Member: Email 2', '')
    member_address = format_address(participant, 'Member:')
    if member_phone1 or member_phone2 or member_email1 or member_email2 or member_address:
        member_parts = ["<b>Member</b>"]
        member_phones = ' / '.join(filter(None, [member_phone1, member_phone2]))
        member_emails = ' / '.join(filter(None, [member_email1, member_email2]))
        if member_phones:
            member_parts.append(f"<b>Phone:</b> {member_phones}")
        if member_emails:
            member_parts.append(f"<b>Email:</b> {member_emails}")
        if member_address:
            member_parts.append(f"<b>Address:</b> {member_address}")
        data.append([Paragraph(" | ".join(member_parts), normal_style)])

    # Row: Primary Contact 1
    if participant['Primary Contact 1: First Name'] or participant['Primary Contact 1: Last Name']:
        primary1_name = f"{participant['Primary Contact 1: First Name']} {participant['Primary Contact 1: Last Name']}".strip()
        primary1_address = format_address(participant, 'Primary Contact 1:')
        primary1_phone = ' / '.join(filter(None, [
            participant['Primary Contact 1: Phone 1'],
            participant['Primary Contact 1: Phone 2']
        ]))
        primary1_email = ' / '.join(filter(None, [
            participant.get('Primary Contact 1: Email 1', ''),
            participant.get('Primary Contact 1: Email 2', '')
        ]))
        p1_parts = [f"<b>{primary1_name}</b>"]
        if primary1_phone:
            p1_parts.append(f"<b>Phone:</b> {primary1_phone}")
        if primary1_email:
            p1_parts.append(f"<b>Email:</b> {primary1_email}")
        if primary1_address:
            p1_parts.append(f"<b>Address:</b> {primary1_address}")
        data.append([Paragraph(" | ".join(p1_parts), normal_style)])

    # Row: Primary Contact 2
    if participant['Primary Contact 2: First Name'] or participant['Primary Contact 2: Last Name']:
        primary2_name = f"{participant['Primary Contact 2: First Name']} {participant['Primary Contact 2: Last Name']}".strip()
        primary2_address = format_address(participant, 'Primary Contact 2:')
        primary2_phone = ' / '.join(filter(None, [
            participant['Primary Contact 2: Phone 1'],
            participant['Primary Contact 2: Phone 2']
        ]))
        primary2_email = ' / '.join(filter(None, [
            participant.get('Primary Contact 2: Email 1', ''),
            participant.get('Primary Contact 2: Email 2', '')
        ]))
        p2_parts = [f"<b>{primary2_name}</b>"]
        if primary2_phone:
            p2_parts.append(f"<b>Phone:</b> {primary2_phone}")
        if primary2_email:
            p2_parts.append(f"<b>Email:</b> {primary2_email}")
        if primary2_address:
            p2_parts.append(f"<b>Address:</b> {primary2_address}")
        data.append([Paragraph(" | ".join(p2_parts), normal_style)])

    # Row: Emergency Contact
    if participant['Emergency Contact: First Name'] or participant['Emergency Contact: Last Name']:
        emergency_name = f"{participant['Emergency Contact: First Name']} {participant['Emergency Contact: Last Name']}".strip()
        emergency_address = format_address(participant, 'Emergency Contact:')
        emergency_phone = ' / '.join(filter(None, [
            participant['Emergency Contact: Phone 1'],
            participant['Emergency Contact: Phone 2']
        ]))
        emergency_email = ' / '.join(filter(None, [
            participant.get('Emergency Contact: Email 1', ''),
            participant.get('Emergency Contact: Email 2', '')
        ]))
        em_parts = [f"<b>Emergency:</b> {emergency_name}"]
        if emergency_phone:
            em_parts.append(f"<b>Phone:</b> {emergency_phone}")
        if emergency_email:
            em_parts.append(f"<b>Email:</b> {emergency_email}")
        if emergency_address:
            em_parts.append(f"<b>Address:</b> {emergency_address}")
        data.append([Paragraph(" | ".join(em_parts), normal_style)])
    
    # Create a table with the data; apply conditional row highlights dynamically
    record_table = Table(data, colWidths=[17*cm])
    table_style_cmds = [
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]
    for row_idx, hex_color in highlight_rows:
        table_style_cmds.append(
            ('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor(hex_color))
        )
    record_table.setStyle(TableStyle(table_style_cmds))
    
    elements.append(record_table)
    elements.append(Spacer(1, 0.05*inch))  # Smaller spacer between records
    
    return elements

def generate_pdf(data, output_file, records_per_page=8):
    """Generate the PDF with multiple record cards per page"""
    # Use the custom document template with page numbering
    doc = NumberedDocTemplate(
        output_file,
        pagesize=A4,
        leftMargin=1*cm,
        rightMargin=1*cm,
        topMargin=1*cm,
        bottomMargin=1.5*cm  # Increased bottom margin to make room for page numbers
    )
    
    all_elements = []
    
    # Page header
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(
        'Header',
        parent=styles['Heading1'],
        fontSize=14,
        alignment=1,  # Center alignment
        spaceAfter=0.1*inch
    )
    
    today = datetime.datetime.now().strftime("%d %B %Y")
    header = Paragraph(f"Participant Personal Details", header_style)
    all_elements.append(header)
    
    # First sort participants by age (adult/youth)
    adults = []
    youths = []
    
    for participant in data:
        age_str = participant['Age at Start']
        # Check if adult (25+ or age >= 18)
        if age_str == '25+':
            adults.append(participant)
        else:
            match = re.match(r'(\d+)\s*/\s*(\d+)', age_str)
            if match:
                years = int(match.group(1))
                if years >= 18:
                    adults.append(participant)
                else:
                    youths.append(participant)
            else:
                # If age format is not recognized, default to youth section
                youths.append(participant)
    
    # Sort each age group by section and last name
    adults.sort(key=lambda x: (x['Section'], x['Last Name']))
    youths.sort(key=lambda x: (x['Section'], x['Last Name']))
    
    # Helper function to create section content with KeepTogether
    def process_section_participants(section, participants, start_number):
        section_elements = []
        
        section_header_style = ParagraphStyle(
            'SectionHeader',
            parent=styles['Heading3'],
            fontSize=10,
            alignment=0,  # Left alignment
            spaceAfter=0.05*inch,
            spaceBefore=0.25*inch
        )
        
        # Process each participant individually (no chunking)
        for i, participant in enumerate(participants):
            # Create participant card elements with sequential number
            participant_elements = create_compact_record_card(participant, number=start_number + i)
            
            # Keep each participant card together
            section_elements.append(KeepTogether(participant_elements))
        
        # Return all section elements and the next available number
        return section_elements, start_number + len(participants)
    
    # Running sequential counter across all age groups and sections
    next_number = 1

    # Process youth participants first
    if youths:
        youth_header_style = ParagraphStyle(
            'AgeGroupHeader',
            parent=styles['Heading2'],
            fontSize=12,
            alignment=0,  # Left alignment
            spaceAfter=0.1*inch,
            spaceBefore=0.1*inch
        )
        
        youth_header = Paragraph("<b>Explorers</b>", youth_header_style)
        all_elements.append(youth_header)
        
        # Group by section
        section_groups = {}
        for participant in youths:
            section = participant['Section']
            if section not in section_groups:
                section_groups[section] = []
            section_groups[section].append(participant)
        
        # Process each section
        for section, section_participants in section_groups.items():
            section_elements, next_number = process_section_participants(
                section, section_participants, next_number
            )
            all_elements.extend(section_elements)
    
    # Process adult participants
    if adults:
        # Add a small spacer after youth section
        if youths:
            all_elements.append(Spacer(1, 0.3*inch))
        
        adult_header_style = ParagraphStyle(
            'AgeGroupHeader',
            parent=styles['Heading2'],
            fontSize=12,
            alignment=0,  # Left alignment
            spaceAfter=0.1*inch,
            spaceBefore=0.1*inch
        )
        adult_header = Paragraph("<b>Adult Participants</b>", adult_header_style)
        all_elements.append(adult_header)
        
        # Group by section
        section_groups = {}
        for participant in adults:
            section = participant['Section']
            if section not in section_groups:
                section_groups[section] = []
            section_groups[section].append(participant)
        
        # Process each section
        for section, section_participants in section_groups.items():
            section_elements, next_number = process_section_participants(
                section, section_participants, next_number
            )
            all_elements.extend(section_elements)
    
    # Build the PDF
    doc.build(all_elements)
    print(f"PDF generated at: {output_file}")

def main():
    """Main function to process data and generate PDF"""
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Generate record cards from participant CSV data')
    parser.add_argument('--input', help='Path to input CSV file', default=INPUT_FILE)
    parser.add_argument('--output', help='Path to output PDF file', default=OUTPUT_FILE)
    parser.add_argument('--records-per-page', type=int, help='Number of records per page', default=8)
    args = parser.parse_args()
    
    # Use provided input/output files or defaults
    input_file = args.input
    output_file = args.output
    records_per_page = args.records_per_page
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Read the CSV data
    print(f"Reading data from: {input_file}")
    data = read_and_clean_csv(input_file)
    
    # Generate the PDF
    generate_pdf(data, output_file, records_per_page)

if __name__ == "__main__":
    main()
