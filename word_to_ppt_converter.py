#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word to PowerPoint Converter
Extracts data from Word document and populates PowerPoint template
"""
from docx import Document
from pptx import Presentation
import re
import sys
import os
from pathlib import Path


class WordToPPTConverter:
    def __init__(self, word_path, ppt_template_path):
        self.word_path = word_path
        self.ppt_template_path = ppt_template_path
        self.data = {}

    def extract_word_data(self):
        """Extract data from Word document"""
        print(f"Reading Word document: {self.word_path}")
        doc = Document(self.word_path)

        # Extract project number from filename or document
        # Try filename first (e.g., "02_785_692_OFFER_APPROVED.docx")
        filename = os.path.basename(self.word_path)

        # First try to find 6-digit project number pattern (XXX_YYY)
        match = re.search(r'(\d{3}_\d{3})', filename)
        if match:
            self.data['project_number'] = match.group(1)
            print(f"  Found project number: {self.data['project_number']}")
        else:
            # Fallback to any number_number pattern
            match = re.search(r'(\d+_\d+)', filename)
            if match:
                self.data['project_number'] = match.group(1)
                print(f"  Found project number: {self.data['project_number']}")

        # If not in filename, try document content
        if 'project_number' not in self.data:
            for para in doc.paragraphs[:10]:
                match = re.search(r'Project\s+Nu[:.]\s*(\d+_\d+)', para.text, re.IGNORECASE)
                if match:
                    self.data['project_number'] = match.group(1)
                    print(f"  Found project number: {self.data['project_number']}")
                    break

        # Extract from header
        for section in doc.sections:
            header = section.header
            for para in header.paragraphs:
                text = para.text
                # Project number in header
                match = re.search(r'Project\s+Nu[:.]\s*(\d+_\d+)', text, re.IGNORECASE)
                if match:
                    self.data['project_number'] = match.group(1)

                # Creation date in header
                match = re.search(r'Creation\s+Date[:.]\s*(\d{4}-\d{2}-\d{2})', text, re.IGNORECASE)
                if match:
                    self.data['creation_date'] = match.group(1)

        # Extract from tables
        for table_idx, table in enumerate(doc.tables):
            for row in table.rows:
                cells = row.cells
                if len(cells) >= 2:
                    key = cells[0].text.strip()
                    value = cells[1].text.strip()

                    # Skip empty values
                    if not value or value == '[empty]':
                        continue

                    # Map table keys to our data dictionary
                    key_lower = key.lower()

                    if 'approval date' in key_lower:
                        self.data['approval_date'] = value
                        print(f"  Found approval date: {value}")
                    elif 'finish of the project (estimated)' in key_lower:
                        self.data['finish_estimated'] = value
                        print(f"  Found finish date (estimated): {value}")
                    elif 'finish of the project (official)' in key_lower:
                        self.data['finish_official'] = value
                        print(f"  Found finish date (official): {value}")
                    elif 'project creation date' in key_lower:
                        self.data['project_creation_date'] = value
                    elif 'offer creation date' in key_lower:
                        self.data['offer_creation_date'] = value
                    elif 'po sent date' in key_lower:
                        self.data['po_sent_date'] = value
                    elif 'plant owner' in key_lower:
                        self.data['plant_owner'] = value
                    elif 'plant code' in key_lower:
                        self.data['plant_code'] = value
                    elif 'name of the part' in key_lower:
                        self.data['part_name'] = value
                    elif 'reference' in key_lower and 'tool' not in key_lower:
                        self.data['reference'] = value
                    elif 'tool number' in key_lower:
                        self.data['tool_number'] = value
                    elif 'se inventory number' in key_lower:
                        self.data['se_inventory_number'] = value
                    elif 'type of service' in key_lower:
                        self.data['type_of_service'] = value
                    elif 'general condition' in key_lower:
                        self.data['general_condition'] = value
                        print(f"  Found general condition: {value}")

            # Extract total cost from pricing table
            for row in table.rows:
                cells = row.cells
                for cell in cells:
                    if 'Total cost' in cell.text:
                        # Look for number in same row
                        for c in cells:
                            # Try to find a number (not in "Total cost" text)
                            text = c.text.strip()
                            if 'Total cost' not in text:
                                match = re.search(r'(\d+)', text)
                                if match:
                                    self.data['total_cost'] = match.group(1)
                                    print(f"  Found total cost: {match.group(1)}")
                                    break

        # Also check content controls
        try:
            from docx.oxml.text.paragraph import CT_P
            all_text = []

            def get_text_from_element(element):
                """Recursively get text from element"""
                texts = []
                if element is not None:
                    text_nodes = element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    for node in text_nodes:
                        if node.text:
                            texts.append(node.text)
                return texts

            all_texts = get_text_from_element(doc.element.body)

            # Search for patterns in all text
            full_text = ' '.join(all_texts)

            # Find dates in format YYYY-MM-DD
            dates = re.findall(r'\d{4}-\d{2}-\d{2}', full_text)
            if dates and 'approval_date' not in self.data:
                self.data['approval_date'] = dates[0]

        except Exception as e:
            print(f"  Warning: Could not extract content controls: {e}")

        print(f"\nExtracted data summary:")
        for key, value in self.data.items():
            print(f"  {key}: {value}")

        return self.data

    def update_powerpoint(self, output_path):
        """Update PowerPoint template with extracted data"""
        print(f"\nUpdating PowerPoint template...")

        prs = Presentation(self.ppt_template_path)
        slide = prs.slides[0]  # Assuming first slide

        # Update title with project number
        if 'project_number' in self.data:
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    text = shape.text_frame.text
                    if "Project 742_051" in text or re.search(r'Project\s+\d+_\d+', text):
                        # Replace project number while preserving formatting
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if re.search(r'Project\s+\d+_\d+', run.text):
                                    run.text = re.sub(r'Project\s+\d+_\d+', f'Project {self.data["project_number"]}', run.text)
                                    print(f"  ✓ Updated title to: Project {self.data['project_number']}")

        # Update Key Intent section
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                text = shape.text_frame.text

                # Update "Repair of the mold"
                if "Repair of the mold" in text and 'type_of_service' in self.data:
                    service_type = self.data['type_of_service']
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "Repair of the mold" in run.text:
                                run.text = run.text.replace("Repair of the mold", f"{service_type} of the mold")
                                print(f"  ✓ Updated service type to: {service_type} of the mold")

                # Update "Located in Gotmar - BG0200P079"
                if "Located in Gotmar" in text:
                    if 'plant_owner' in self.data and 'plant_code' in self.data:
                        new_location = f"Located in {self.data['plant_owner']} - {self.data['plant_code']}"
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "Located in Gotmar" in run.text or "BG0200P079" in run.text:
                                    # Replace the whole line
                                    run.text = re.sub(r'Located in .+ - .+', new_location, run.text)
                                    if run.text == new_location or new_location in run.text:
                                        print(f"  ✓ Updated location to: {new_location}")

                # Update "Inv Nu: SEP0431"
                if "Inv Nu:" in text and 'se_inventory_number' in self.data:
                    inv_num = self.data['se_inventory_number']
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "Inv Nu:" in run.text or "SEP" in run.text:
                                run.text = re.sub(r'Inv Nu:\s*\S+', f'Inv Nu: {inv_num}', run.text)
                                if inv_num in run.text:
                                    print(f"  ✓ Updated inventory number to: {inv_num}")

                # Update "General state of the mold: Bad"
                if "General state of the mold" in text:
                    # Try to get from extracted data, or use default
                    state = self.data.get('general_condition', 'Bad')
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "General state of the mold:" in run.text:
                                run.text = re.sub(r'General state of the mold:\s*\w+', f'General state of the mold: {state}', run.text)
                                print(f"  ✓ Updated general state to: {state}")

        # Update Finance table (preserving formatting)
        for shape in slide.shapes:
            if shape.shape_type == 19:  # Table
                try:
                    table = shape.table
                    # Check if this is the finance table
                    if len(table.rows) >= 3 and len(table.columns) >= 2:
                        first_cell = table.rows[0].cells[0].text_frame.text.strip()
                        if 'Required Capex' in first_cell or 'Capex' in first_cell:
                            # This is the finance table
                            if 'total_cost' in self.data:
                                cost = self.data['total_cost']
                                # Update values while preserving formatting
                                for paragraph in table.rows[0].cells[1].text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        run.text = f"{cost}€"
                                table.rows[0].cells[1].text_frame.paragraphs[0].runs[0].text = f"{cost}€"
                                table.rows[1].cells[1].text_frame.text = ""  # Required Opex (empty)
                                for paragraph in table.rows[2].cells[1].text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        run.text = f"{cost}€"
                                table.rows[2].cells[1].text_frame.paragraphs[0].runs[0].text = f"{cost}€"
                                print(f"  ✓ Updated Finance table with cost: {cost}€")
                except Exception as e:
                    print(f"  Warning: Error updating Finance table: {e}")

        # Update Initial Planning dates
        def update_dates_in_group(group_shape, date_mapping):
            """Recursively update dates in grouped shapes"""
            try:
                if group_shape.shape_type == 6:  # GROUP
                    for sub_shape in group_shape.shapes:
                        update_dates_in_group(sub_shape, date_mapping)
                elif hasattr(group_shape, "text_frame"):
                    text = group_shape.text_frame.text.strip()

                    # Check if this is a date text box
                    if text and re.match(r'\d{4}-\d{2}', text):
                        # Find the corresponding label to determine which date to use
                        for label, date_value in date_mapping.items():
                            if date_value:
                                # Update the date (keeping format YYYY-MM)
                                if len(text) <= 7:  # Format YYYY-MM
                                    new_date = date_value[:7]  # Take YYYY-MM part
                                    for paragraph in group_shape.text_frame.paragraphs:
                                        for run in paragraph.runs:
                                            if re.match(r'\d{4}-\d{2}', run.text):
                                                run.text = new_date
            except:
                pass

        # Prepare date mapping
        date_mapping = {}
        if 'approval_date' in self.data:
            date_mapping['BCI Validation'] = self.data['approval_date']
        if 'finish_estimated' in self.data:
            date_mapping['Finish of the project'] = self.data['finish_estimated']

        # Find and update Initial Planning group
        for shape in slide.shapes:
            if shape.shape_type == 6:  # GROUP
                try:
                    # Check if this group contains "Initial Planning"
                    for sub_shape in shape.shapes:
                        if hasattr(sub_shape, "text_frame") and "Initial Planning" in sub_shape.text_frame.text:
                            print(f"  ✓ Found Initial Planning group")
                            if date_mapping:
                                update_dates_in_group(shape, date_mapping)
                                print(f"  ✓ Updated dates in Initial Planning")
                            break
                except:
                    pass

        # Save the presentation
        prs.save(output_path)
        print(f"\nSaved updated PowerPoint to: {output_path}")

    def convert(self, output_dir=None):
        """Main conversion method"""
        # Extract data from Word
        self.extract_word_data()

        # Determine output filename
        if output_dir is None:
            output_dir = os.path.dirname(self.word_path)

        project_name = self.data.get('project_number', 'Project')
        output_filename = f"Project {project_name}.pptx"
        output_path = os.path.join(output_dir, output_filename)

        # Update PowerPoint
        self.update_powerpoint(output_path)

        return output_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python word_to_ppt_converter.py <word_file> [ppt_template] [output_dir]")
        print("\nExample:")
        print("  python word_to_ppt_converter.py document.docx")
        print("  python word_to_ppt_converter.py document.docx template.pptx")
        print("  python word_to_ppt_converter.py document.docx template.pptx D:/output")
        sys.exit(1)

    word_file = sys.argv[1]

    # Use default template or provided one
    if len(sys.argv) >= 3:
        ppt_template = sys.argv[2]
    else:
        # Look for template in same directory
        ppt_template = "Project 742_051.pptx"
        if not os.path.exists(ppt_template):
            print(f"Error: Template not found: {ppt_template}")
            print("Please provide template path as second argument")
            sys.exit(1)

    # Output directory
    output_dir = sys.argv[3] if len(sys.argv) >= 4 else None

    # Convert
    converter = WordToPPTConverter(word_file, ppt_template)
    output_path = converter.convert(output_dir)

    print(f"\n✓ Conversion complete!")
    print(f"✓ Output file: {output_path}")


if __name__ == "__main__":
    main()
