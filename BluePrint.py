import xml.etree.ElementTree as ET
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

def set_font(cell, font_name, font_size, font_color=None, bold=False):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            if font_color:
                run.font.color.rgb = font_color  
            if bold:
                run.font.bold = True  
            r = run._element
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                r.insert(0, rPr)
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)

def set_cell_background(cell, color):
    cell_properties = cell._element.get_or_add_tcPr()
    cell_shading = OxmlElement('w:shd')
    cell_shading.set(qn('w:fill'), color)  
    cell_properties.append(cell_shading)

def update_names(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    date_format = "%m/%d/%Y"
    in_date = datetime.strptime("10/10/2020", date_format)
    
    doc = Document()
    bc = root.find(".//BUSINESS_COMPONENT")
    bc_name = bc.get("NAME")
    doc.add_heading(f"{bc_name}: Fields", level=1)
    
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Field Name"
    hdr_cells[1].text = "Calculated"
    hdr_cells[2].text = "Calculated Value"
    hdr_cells[3].text = "COLUMN"
    hdr_cells[4].text = "JOIN"

    for cell in hdr_cells:
        set_font(cell, "Segoe UI Semilight", 10, font_color=RGBColor(255, 255, 255), bold=True)  
        set_cell_background(cell, "3b6982") 

    field_found = False 
    for component in root.findall(".//FIELD"):
        field_updated = component.get("UPDATED")
        if field_updated:
            date_only = field_updated.split(" ")[0]
            date_object = datetime.strptime(date_only, date_format)
            
            if date_object > in_date:
                field_found = True
                field_name = component.get("NAME")
                calculated = component.get("CALCULATED")
                calculated_value = component.get("CALCULATED_VALUE")
                field_column = component.get("COLUMN")
                field_join = component.get("JOIN")
                
                row_cells = table.add_row().cells
                row_cells[0].text = field_name if field_name else "N/A"
                row_cells[1].text = calculated if calculated else "N/A"
                row_cells[2].text = calculated_value if calculated_value else "N/A"
                row_cells[3].text = field_column if field_column else "N/A"
                row_cells[4].text = field_join if field_join else "N/A"

                for cell in row_cells:
                    set_font(cell, "Segoe UI Semilight", 10)

    if not field_found:
        doc.tables[0]._element.getparent().remove(doc.tables[0]._element)

    user_prop_found = False  
    for user_prop in root.findall(".//BUSINESS_COMPONENT_USER_PROP"):
        user_updated = user_prop.get("UPDATED")
        if user_updated:
            date_only = user_updated.split(" ")[0]
            date_object = datetime.strptime(date_only, date_format)
            
            if date_object > in_date:
                if not user_prop_found:
                    doc.add_heading(f"{bc_name}: User Properties Changes", level=1)
                    user_prop_table = doc.add_table(rows=1, cols=2)
                    user_prop_table.style = "Table Grid"
                    user_hdr_cells = user_prop_table.rows[0].cells
                    user_hdr_cells[0].text = "NAME"
                    user_hdr_cells[1].text = "VALUE"

                    for cell in user_hdr_cells:
                        set_font(cell, "Segoe UI Semilight", 10, font_color=RGBColor(255, 255, 255), bold=True)
                        set_cell_background(cell, "3b6982")  

                user_prop_found = True
                name = user_prop.get("NAME")
                value = user_prop.get("VALUE")
                
                row_cells = user_prop_table.add_row().cells
                row_cells[0].text = name if name else "N/A"
                row_cells[1].text = value if value else "N/A"

                for cell in row_cells:
                    set_font(cell, "Segoe UI Semilight", 10)
                    
    script_found = False  
    for ser_script in root.findall(".//BUSCOMP_SERVER_SCRIPT"):
        print("Script")
        user_updated = ser_script.get("UPDATED")
        if user_updated:
            date_only = user_updated.split(" ")[0]
            date_object = datetime.strptime(date_only, date_format)
            
            if date_object > in_date:
                print("Script1")
                if not script_found:
                    print("Script2")
                    doc.add_heading(f"{bc_name}: Bus Comp Server Script", level=1)
                    serv_script_table = doc.add_table(rows=1, cols=2)
                    serv_script_table.style = "Table Grid"
                    user_hdr_cells = serv_script_table.rows[0].cells
                    user_hdr_cells[0].text = "NAME"
                    user_hdr_cells[1].text = "SCRIPT"

                    for cell in user_hdr_cells:
                        set_font(cell, "Segoe UI Semilight", 10, font_color=RGBColor(255, 255, 255), bold=True)
                        set_cell_background(cell, "3b6982")  

                script_found = True
                name = ser_script.get("NAME")
                value = ser_script.get("SCRIPT")
                
                row_cells = serv_script_table.add_row().cells
                row_cells[0].text = name if name else "N/A"
                row_cells[1].text = value if value else "N/A"

                for cell in row_cells:
                    set_font(cell, "Segoe UI Semilight", 10)

    output_file = "output.docx"
    doc.save(output_file)

    if field_found or user_prop_found:
        print(f"Tables written to {output_file}")
    else:
        print("No matches found. No tables created.")

input_file = input("Enter the file name with extension: ")
update_names(input_file)
