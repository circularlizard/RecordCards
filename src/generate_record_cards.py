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

def create_compact_record_card(participant):
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
    title = Paragraph(f"<b>{name}</b> - {section}", title_style)
    elements.append(title)
    
    # Create a table for the rest of the information
    data = []
    
    # Row 1: Personal details (removed school and religion)
    personal_info = [
        f"<b>Age:</b> {participant['Age Formatted']}",
        f"<b>Swimmer:</b> {participant['Swimmer']}"
    ]
    
    # Filter out empty values and join with pipe symbols
    personal_info_filtered = [p for p in personal_info if p.endswith('>') is False]
    if personal_info_filtered:
        personal_info_text = Paragraph(" | ".join(personal_info_filtered), normal_style)
        data.append([personal_info_text])
    
    # Row 2: Medical information
    medical_info = []
    if participant['Medical details']:
        medical_info.append(f"<b>Medical:</b> {participant['Medical details']}")
    if participant['Allergies']:
        medical_info.append(f"<b>Allergies:</b> {participant['Allergies']}")
    if participant['Dietary requirements']:
        medical_info.append(f"<b>Diet:</b> {participant['Dietary requirements']}")
    if participant['Tetanus (year of last jab)']:
        medical_info.append(f"<b>Tetanus:</b> {participant['Tetanus (year of last jab)']}")
    if participant['Other useful information']:
        medical_info.append(f"<b>Other:</b> {participant['Other useful information']}")
    
    if medical_info:
        medical_info_text = Paragraph(" | ".join(medical_info), normal_style)
        data.append([medical_info_text])
    
    # Row 3: Primary Contact 1
    if participant['Primary Contact 1: First Name'] or participant['Primary Contact 1: Last Name']:
        primary1_name = f"{participant['Primary Contact 1: First Name']} {participant['Primary Contact 1: Last Name']}".strip()
        primary1_address = format_address(participant, 'Primary Contact 1:')
        primary1_phone = ' / '.join(filter(None, [
            participant['Primary Contact 1: Phone 1'],
            participant['Primary Contact 1: Phone 2']
        ]))
        
        primary1_text = Paragraph(f"<b>Primary 1:</b> {primary1_name} | <b>Phone:</b> {primary1_phone} | <b>Address:</b> {primary1_address.replace(', ', ', ')}", normal_style)
        data.append([primary1_text])
    
    # Row 4: Primary Contact 2
    if participant['Primary Contact 2: First Name'] or participant['Primary Contact 2: Last Name']:
        primary2_name = f"{participant['Primary Contact 2: First Name']} {participant['Primary Contact 2: Last Name']}".strip()
        primary2_address = format_address(participant, 'Primary Contact 2:')
        primary2_phone = ' / '.join(filter(None, [
            participant['Primary Contact 2: Phone 1'],
            participant['Primary Contact 2: Phone 2']
        ]))
        
        primary2_text = Paragraph(f"<b>Primary 2:</b> {primary2_name} | <b>Phone:</b> {primary2_phone} | <b>Address:</b> {primary2_address.replace(', ', ', ')}", normal_style)
        data.append([primary2_text])
    
    # Row 5: Emergency Contact
    if participant['Emergency Contact: First Name'] or participant['Emergency Contact: Last Name']:
        emergency_name = f"{participant['Emergency Contact: First Name']} {participant['Emergency Contact: Last Name']}".strip()
        emergency_address = format_address(participant, 'Emergency Contact:')
        emergency_phone = ' / '.join(filter(None, [
            participant['Emergency Contact: Phone 1'],
            participant['Emergency Contact: Phone 2']
        ]))
        
        emergency_text = Paragraph(f"<b>Emergency:</b> {emergency_name} | <b>Phone:</b> {emergency_phone} | <b>Address:</b> {emergency_address.replace(', ', ', ')}", normal_style)
        data.append([emergency_text])
    
    # Create a table with the data
    record_table = Table(data, colWidths=[17*cm])
    record_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Reduced padding
        ('TOPPADDING', (0, 0), (-1, -1), 2),     # Reduced padding
        ('LEFTPADDING', (0, 0), (-1, -1), 4),    # Reduced padding
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),   # Reduced padding
    ]))
    
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
    def process_section_participants(section, participants, current_record_count):
        section_elements = []
        
        section_header_style = ParagraphStyle(
            'SectionHeader',
            parent=styles['Heading3'],
            fontSize=10,
            alignment=0,  # Left alignment
            spaceAfter=0.05*inch,
            spaceBefore=0.25*inch
        )
        
        # Add a section header if needed
        section_header = Paragraph(f"<b>{section}</b>", section_header_style)
        section_elements.append(section_header)
        
        # Process each participant individually (no chunking)
        for participant in participants:
            # Create participant card elements
            participant_elements = create_compact_record_card(participant)
            
            # Keep each participant card together
            section_elements.append(KeepTogether(participant_elements))
        
        # Return all section elements and count of participants
        return section_elements, len(participants)
    
    # Track records on current page
    records_on_page = 0
    
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
        for section_idx, (section, section_participants) in enumerate(section_groups.items()):
            # Process this section's participants
            section_elements, _ = process_section_participants(
                section, 
                section_participants, 
                0  # No longer tracking records on page
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
        for section_idx, (section, section_participants) in enumerate(section_groups.items()):
            # Process this section's participants
            section_elements, _ = process_section_participants(
                section, 
                section_participants, 
                0  # No longer tracking records on page
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
