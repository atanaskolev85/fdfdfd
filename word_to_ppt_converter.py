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

        # First, try to extract from customXml (Document Properties)
        import zipfile
        import xml.etree.ElementTree as ET

        try:
            with zipfile.ZipFile(self.word_path, 'r') as docx_zip:
                try:
                    # Read customXml/item3.xml which contains all the properties
                    xml_content = docx_zip.read('customXml/item3.xml')
                    root = ET.fromstring(xml_content)

                    print("  Reading data from Document Properties (customXml)...")

                    # Extract all properties (they are children of root elements)
                    for parent in root:
                        for elem in parent:
                            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                            value = elem.text.strip() if elem.text else None

                            if value:
                                # Map XML tags to our data dictionary
                                if tag == 'Plantowner':
                                    self.data['plant_owner'] = value
                                    print(f"  ✓ Plant owner: {value}")
                                elif tag == 'Sitename':
                                    self.data['plant_code'] = value
                                    print(f"  ✓ Plant code: {value}")
                                elif tag == 'Nameofthepart':
                                    self.data['part_name'] = value
                                    print(f"  ✓ Part name: {value}")
                                elif tag == 'Ref_':
                                    self.data['reference'] = value.replace('\n', ' ')
                                    print(f"  ✓ Reference: {value[:30]}...")
                                elif 'Toolnumber' in tag:
                                    self.data['tool_number'] = value
                                    print(f"  ✓ Tool number: {value}")
                                elif tag == 'SEinventorynumber':
                                    self.data['se_inventory_number'] = value
                                    print(f"  ✓ SE inventory: {value}")
                                elif tag == 'Inventory_number_Gormar':
                                    self.data['gotmar_inventory'] = value
                                elif tag == 'GeneralToolCondition':
                                    self.data['general_condition'] = value
                                    print(f"  ✓ General condition: {value}")
                                elif tag == 'Typeofservice':
                                    self.data['type_of_service'] = value
                                    print(f"  ✓ Type of service: {value}")
                                elif tag == 'Creationdate':
                                    self.data['project_creation_date'] = value
                                elif tag == 'Offercreationdate':
                                    self.data['offer_creation_date'] = value
                                elif tag == 'Approvaldate':
                                    self.data['approval_date'] = value
                                    print(f"  ✓ Approval date: {value}")
                                elif tag == 'Finishoftheproject':
                                    self.data['finish_estimated'] = value
                                    print(f"  ✓ Finish date: {value}")
                                elif tag == 'Totalcost':
                                    self.data['total_cost'] = value
                                    print(f"  ✓ Total cost: {value}")
                                elif tag == 'ProjectStatus':
                                    self.data['project_status'] = value
                                elif tag == 'PRIORITY':
                                    self.data['priority'] = value

                except KeyError:
                    print("  Note: customXml/item3.xml not found, using legacy table extraction...")

        except Exception as e:
            print(f"  Note: Could not read customXml ({e}), using legacy table extraction...")

        # Extract project number from filename or document
        # Try filename first (e.g., "02_785_692_OFFER_APPROVED.docx")
        filename = os.path.basename(self.word_path)

        # First try to find 6-digit project number pattern (XXX_YYY)
        if 'project_number' not in self.data:
            match = re.search(r'(\d{3}_\d{3})', filename)
            if match:
                self.data['project_number'] = match.group(1)
                print(f"  ✓ Project number: {self.data['project_number']}")
            else:
                # Fallback to any number_number pattern
                match = re.search(r'(\d+_\d+)', filename)
                if match:
                    self.data['project_number'] = match.group(1)
                    print(f"  ✓ Project number: {self.data['project_number']}")

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
                    if "Project 742_051" in text or "Project" in text:
                        # Replace project number
                        new_text = re.sub(r'Project\s+\d+_\d+', f'Project {self.data["project_number"]}', text)
                        if new_text != text:
                            shape.text_frame.text = new_text
                            print(f"  Updated title to: Project {self.data['project_number']}")

        # Update Finance table
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
                                # Update values
                                table.rows[0].cells[1].text_frame.text = f"{cost}€"  # Required Capex
                                table.rows[1].cells[1].text_frame.text = ""  # Required Opex (empty)
                                table.rows[2].cells[1].text_frame.text = f"{cost}€"  # TOTAL
                                print(f"  Updated Finance table with cost: {cost}€")
                except Exception as e:
                    print(f"  Warning: Error updating table: {e}")

        # Update Initial Planning dates
        def update_dates_in_group(group_shape):
            """Recursively update dates in grouped shapes"""
            try:
                if group_shape.shape_type == 6:  # GROUP
                    for sub_shape in group_shape.shapes:
                        update_dates_in_group(sub_shape)
                elif hasattr(group_shape, "text_frame"):
                    text = group_shape.text_frame.text.strip()
                    # Update specific date fields based on your requirements
                    # This is a simplified version - you may need to adjust based on exact layout
                    if text and re.match(r'\d{4}-\d{2}', text):
                        # This is a date field - update if we have approval or finish date
                        if 'approval_date' in self.data:
                            # You can implement more specific logic here
                            pass
            except:
                pass

        # Find and update grouped shapes
        for shape in slide.shapes:
            if shape.shape_type == 6:  # GROUP
                try:
                    # Check if this group contains "Initial Planning"
                    for sub_shape in shape.shapes:
                        if hasattr(sub_shape, "text_frame") and "Initial Planning" in sub_shape.text_frame.text:
                            print(f"  Found Initial Planning group")
                            update_dates_in_group(shape)
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
