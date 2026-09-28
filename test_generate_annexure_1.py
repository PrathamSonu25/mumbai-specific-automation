import os
import sys
# pyrefly: ignore [missing-import]
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

DOCS_DIR = os.path.join(os.path.dirname(__file__), 'Documents')
OUT_DIR = os.path.join(os.path.dirname(__file__), 'generated_docs')
os.makedirs(OUT_DIR, exist_ok=True)

def set_cell_margins(cell, top=100, bottom=100, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_borders(cell, color='CCCCCC', sz='4', val='single'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), val)
        el.set(qn('w:sz'), sz)
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

from doc_generator import format_annexure_1_model_details

def build_test_annexure_1():
    src = os.path.join(DOCS_DIR, 'ANNEXURE-1 Identification and Physical Inventory of Relevant Devices.docx')
    doc = docx.Document(src)
    
    # Set page margins to 0.6 inches for a spacious, professional table layout
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)
        
    # Team header
    for p in doc.paragraphs:
        if 'TEAM CODE:' in p.text:
            p.text = 'TEAM CODE: Team Alpha-MUM'
            p.runs[0].font.bold = True
            
    t = doc.tables[0]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    
    # Remove existing empty rows 1-6
    while len(t.rows) > 1:
        tr = t.rows[-1]._tr
        tr.getparent().remove(tr)
        
    # Column widths (Total ~ 7.05 inches, fits portrait perfectly)
    col_widths = [
        Inches(1.0),   # 0: Evidence Particulars
        Inches(1.15),  # 1: Nature of Evidence
        Inches(1.65),  # 2: Model Details
        Inches(0.95),  # 3: Found in possession
        Inches(0.65),  # 4: Is Owner
        Inches(0.85),  # 5: Relation to target
        Inches(0.80)   # 6: Remarks
    ]
    
    # Format header row
    hdr_row = t.rows[0]
    for c_i, cell in enumerate(hdr_row.cells):
        cell.width = col_widths[c_i]
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        set_cell_shading(cell, 'EAEFF5')
        set_cell_borders(cell, color='999999', sz='6')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(8.5)
                r.font.name = 'Calibri'
                
    # Sample test evidences
    test_evidences = [
        # 3 Desktops
        {
            'evidence_category': 'Desktop',
            'device_name': 'Trading Terminal 1',
            'make': 'Dell', 'model': 'OptiPlex 7090', 'serial_number': 'DL-OPT-01',
            'capacity': '512 GB NVMe', 'color': 'Black', 'os': 'Windows 11 Pro',
            'owner_user': 'Rajesh Sharma', 'is_owner': 'No', 'owner_relation': 'Senior Trader',
            'notes_by_officer': 'Found on trading desk 4'
        },
        {
            'evidence_category': 'Desktop',
            'device_name': 'Trading Terminal 2',
            'make': 'Dell', 'model': 'OptiPlex 7090', 'serial_number': 'DL-OPT-02',
            'capacity': '512 GB NVMe', 'color': 'Black', 'os': 'Windows 11 Pro',
            'owner_user': 'Rajesh Sharma', 'is_owner': 'No', 'owner_relation': 'Senior Trader',
            'notes_by_officer': 'Found on trading desk 5'
        },
        {
            'evidence_category': 'Desktop',
            'device_name': 'Accounts Terminal',
            'make': 'HP', 'model': 'ProDesk 600', 'serial_number': 'HP-PRO-88',
            'capacity': '1 TB SSD', 'color': 'Dark Grey', 'os': 'Windows 10 Pro',
            'owner_user': 'Amit Verma', 'is_owner': 'No', 'owner_relation': 'Chief Accountant',
            'notes_by_officer': 'Accounts cubicle'
        },
        # 1 Laptop
        {
            'evidence_category': 'Laptop',
            'device_name': 'MD Trading Laptop',
            'make': 'Lenovo', 'model': 'ThinkPad P1 Gen 6', 'serial_number': 'SN-LEN-99120',
            'capacity': '2 TB NVMe SSD', 'color': 'Midnight Black', 'os': 'Windows 11 Pro',
            'owner_user': 'Vikramaditya Singhania', 'is_owner': 'Yes', 'owner_relation': 'Managing Director',
            'notes_by_officer': 'Cabin of MD'
        },
        # 1 Mobile
        {
            'evidence_category': 'Mobile',
            'device_name': 'CFO Mobile Phone',
            'make': 'Apple', 'model': 'iPhone 15 Pro', 'serial_number': 'F2LM8900XYZ',
            'capacity': '256 GB', 'color': 'Natural Titanium', 'os': 'iOS 17.5.1',
            'imei1': '354890123456789', 'imei2': '354890123456780',
            'owner_user': 'Pooja Mehta', 'is_owner': 'Yes', 'owner_relation': 'CFO & Key Managerial Person',
            'notes_by_officer': 'Seized from handbag'
        },
        # 1 Cloud
        {
            'evidence_category': 'Cloud',
            'device_name': 'CEO Exchange Mailbox Backup',
            'make': 'Microsoft M365', 'model': 'Exchange Online', 'serial_number': 'M365-EXCH-01',
            'capacity': '85 GB PST', 'email_id': 'vikram.singhania@marineequities.in',
            'cloud_id': 'marineequities_tenant_guid_8892', 'domain_name': 'marineequities.in',
            'owner_user': 'Vikramaditya Singhania', 'is_owner': 'No', 'owner_relation': 'Promoter / Account Holder',
            'notes_by_officer': 'Compliance export'
        }
    ]
    
    # Categories to evaluate
    standard_categories = ['Desktop', 'Laptop', 'Mobile', 'Pendrive', 'Harddisk', 'Server']
    
    # Detect any extra categories in evidences
    extra_categories = []
    for ev in test_evidences:
        cat = ev.get('evidence_category', '')
        if cat in ['HDD', 'SSD']:
            cat_norm = 'Harddisk'
        elif cat in ['Flash Drive']:
            cat_norm = 'Pendrive'
        else:
            cat_norm = cat
            
        if cat_norm not in standard_categories and cat_norm not in extra_categories:
            extra_categories.append(cat_norm)
            
    all_categories = standard_categories + extra_categories
    
    for cat_name in all_categories:
        # Match evidences
        matching = []
        for ev in test_evidences:
            c = ev.get('evidence_category', '')
            if cat_name == 'Harddisk' and c in ['Harddisk', 'HDD', 'SSD']:
                matching.append(ev)
            elif cat_name == 'Pendrive' and c in ['Pendrive', 'Flash Drive']:
                matching.append(ev)
            elif cat_name == 'File/Folder' and c in ['File/Folder', 'FileFolder']:
                matching.append(ev)
            elif c == cat_name:
                matching.append(ev)
                
        if matching:
            start_row = len(t.rows)
            for ev in matching:
                r = t.add_row()
                r.cells[0].text = cat_name
                loc = ev.get('location_found', '').strip() or ev.get('notes_by_officer', '').strip() or 'Recovered from premise'
                r.cells[1].text = loc
                r.cells[2].text = format_annexure_1_model_details(ev)
                r.cells[3].text = ev.get('owner_user', '') or 'N/A'
                is_owner_val = ev.get('is_owner', 'Yes')
                owner_name = ev.get('owner_user', '').strip() or 'Yes'
                r.cells[4].text = owner_name if (is_owner_val == 'Yes' or is_owner_val is True) else 'No'
                r.cells[5].text = ev.get('owner_relation', '') or 'N/A'
                r.cells[6].text = ''
                
                # Format cell styles
                for c_i, cell in enumerate(r.cells):
                    cell.width = col_widths[c_i]
                    set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
                    set_cell_borders(cell, color='CCCCCC', sz='4')
                    for p in cell.paragraphs:
                        if c_i in [0, 4]:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        for run in p.runs:
                            run.font.size = Pt(8.5)
                            run.font.name = 'Calibri'
                            
            end_row = len(t.rows) - 1
            if end_row > start_row:
                merged = t.cell(start_row, 0).merge(t.cell(end_row, 0))
                merged.text = cat_name
                merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = merged.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if p.runs:
                    p.runs[0].font.bold = True
                    p.runs[0].font.size = Pt(9)
            else:
                p = t.cell(start_row, 0).paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if p.runs:
                    p.runs[0].font.bold = True
                    p.runs[0].font.size = Pt(9)
        else:
            # If standard category has no devices, add 1 row with -
            r = t.add_row()
            r.cells[0].text = cat_name
            r.cells[1].text = '-'
            r.cells[2].text = '-'
            r.cells[3].text = '-'
            r.cells[4].text = '-'
            r.cells[5].text = '-'
            r.cells[6].text = '-'
            for c_i, cell in enumerate(r.cells):
                cell.width = col_widths[c_i]
                set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
                set_cell_borders(cell, color='E0E0E0', sz='4')
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        run.font.size = Pt(8.5)
                        run.font.name = 'Calibri'
                        if c_i == 0:
                            run.font.bold = True

    # Signature paragraphs
    for p in doc.paragraphs:
        if 'SIGN OF THE AUTHORISED OFFICER:' in p.text:
            p.text = 'SIGN OF THE AUTHORISED OFFICER: __________________________\n(Shri V. K. Malhotra, General Manager)'
        elif 'DATE:' in p.text:
            p.text = 'DATE: 26/09/2026'

    out_file = os.path.join(OUT_DIR, 'test_mumbai_annexure_1_demo.docx')
    doc.save(out_file)
    print(f"Generated test Annexure 1 at: {out_file}")
    return out_file

if __name__ == '__main__':
    build_test_annexure_1()
