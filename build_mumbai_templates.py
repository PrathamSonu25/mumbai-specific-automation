import os
import sys
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

DOCS_DIR = os.path.join(os.path.dirname(__file__), 'Documents')
OUT_DIR = os.path.join(os.path.dirname(__file__), 'templates_docx')
os.makedirs(OUT_DIR, exist_ok=True)

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
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

# -------------------------------------------------------------
# 1. Mumbai Annexure 1: Base Template
# -------------------------------------------------------------
def build_mumbai_annexure_1():
    src = os.path.join(DOCS_DIR, 'ANNEXURE-1 Identification and Physical Inventory of Relevant Devices.docx')
    doc = docx.Document(src)
    
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)
        
    out_path = os.path.join(OUT_DIR, 'mumbai_annexure_1.docx')
    doc.save(out_path)
    print("Generated:", out_path)

# -------------------------------------------------------------
# 2. Mumbai Annexure 3: Chain of Custody (Mumbai COC)
# -------------------------------------------------------------
def build_mumbai_annexure_3():
    src = os.path.join(DOCS_DIR, 'Mumbai COC.docx')
    doc = docx.Document(src)
    
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Table 0: Headers & Repeating Evidences
    t0 = doc.tables[0]
    
    t0.rows[0].cells[0].text = 'CHAIN OF CUSTODY'
    t0.rows[1].cells[0].text = 'SEARCH WARRANT DETAILS: {{ search_warrant_details }}'
    t0.rows[2].cells[0].text = 'AUTHORISED OFFICER(S) NAME: {{ authorised_officer }}'
    t0.rows[3].cells[0].text = 'PARTY NAME: In the case of {{ case_name }}'
    t0.rows[4].cells[0].text = 'ADDRESS: {{ address }}'
    t0.rows[5].cells[0].text = 'TEAM : {{ team }}\tANNEXURE: {{ annexure_no }}'
    
    t0.rows[8].cells[0].text = 'DESCRIPTION OF THE DEVICE'
    t0.rows[8].cells[2].text = 'MASTER COPY DETAILS'
    t0.rows[8].cells[4].text = 'WORKING COPY DETAILS'
    
    # Remove existing rows 9, 10
    while len(t0.rows) > 9:
        tr = t0.rows[-1]._tr
        tr.getparent().remove(tr)
        
    # Add row for loop start
    r_loop_start = t0.add_row()
    r_loop_start.cells[0].text = '{%tr for ev in evidences %}'
    
    # Add data row
    r_data = t0.add_row()
    c0 = r_data.cells[0]
    c0.merge(r_data.cells[1])
    c2 = r_data.cells[2]
    c2.merge(r_data.cells[3])
    c4 = r_data.cells[4]
    c4.merge(r_data.cells[5]).merge(r_data.cells[6])
    
    c0.text = (
        'EVIDENCE DEVICE {{ loop.index }}\n\n'
        'NAME: {{ ev.device_name }}\n'
        'MAKE MODEL: {{ ev.make_model }}\tS/N: {{ ev.serial_number }}\n'
        'OS: {{ ev.os }}\n'
        'SIZE: {{ ev.capacity }}\tIMEI1: {{ ev.imei1 }}\tIMEI2: {{ ev.imei2 }}\n'
        'IP ADDRESS: {{ ev.ip_address }}\tDOMAIN NAME: {{ ev.domain_name }}\tEMAIL ID: {{ ev.email_id }}'
    )
    
    c2.text = (
        'HASH (MD5): {{ ev.md5_hash }}\n'
        'HASH (SHA1): {{ ev.sha1_hash }}\n\n'
        'MASTER COPY DEVICE DETAILS:\n'
        'S.No: {{ ev.master_copy_sn }}\n'
        'SIZE: {{ ev.master_copy_capacity }}\n'
        'MAKE MODEL: {{ ev.master_copy_make_model }}'
    )
    
    c4.text = (
        'HASH (MD5): {{ ev.md5_hash }}\n'
        'HASH (SHA1): {{ ev.sha1_hash }}\n\n'
        'WORKING COPY DEVICE DETAILS:\n'
        'S.No: {{ ev.working_copy_sn }}\n'
        'SIZE: {{ ev.working_copy_capacity }}\n'
        'MAKE MODEL: {{ ev.working_copy_make_model }}'
    )
    
    # Add row for loop end
    r_loop_end = t0.add_row()
    r_loop_end.cells[0].text = '{%tr endfor %}'
    
    # Table 1: Clean 6-Column Chain of Custody Documentation
    t1_old = doc.tables[1]
    t1_p = t1_old._tbl.getparent()
    t1_p.remove(t1_old._tbl)
    
    t1 = doc.add_table(rows=4, cols=6)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.autofit = False
    
    col_widths = [Inches(2.2), Inches(1.5), Inches(1.15), Inches(1.15), Inches(0.85), Inches(1.15)]
    
    # Row 0: Title
    r0 = t1.rows[0]
    r0.cells[0].merge(r0.cells[5])
    r0.cells[0].text = 'CHAIN OF CUSTODY DOCUMENTATION'
    p0 = r0.cells[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.runs[0].font.bold = True
    p0.runs[0].font.size = Pt(11)
    set_cell_shading(r0.cells[0], 'EAEFF5')
    set_cell_margins(r0.cells[0], top=100, bottom=100, left=100, right=100)
    
    # Row 1: Headers
    headers_t1 = [
        'EVIDENCE DEVICES',
        'PURPOSE OF TRANSFER OF CUSTODY',
        'RELEASED BY\n(Name & Sign)',
        'RECEIVED BY\n(Name & Sign)',
        'DATE & TIME',
        'REMARKS'
    ]
    r1 = t1.rows[1]
    for c_i, h in enumerate(headers_t1):
        r1.cells[c_i].text = h
        p = r1.cells[c_i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(9)
        set_cell_shading(r1.cells[c_i], 'F2F4F7')
        set_cell_margins(r1.cells[c_i], top=80, bottom=80, left=80, right=80)
        
    # Row 2: Transfer 1 (Handing over after imaging)
    r2 = t1.rows[2]
    r2.cells[0].text = '{{ transfer_1_devices }}'
    r2.cells[1].text = 'HANDING OVER OF HDD WITH DATA AFTER FORENSIC BACKUP AND HASHING'
    r2.cells[2].text = 'Authorised officer (Name & Sign)\n\n{{ authorised_officer }}'
    r2.cells[3].text = 'Forensic Expert (Name & Sign)\n\n{{ expert_name }}'
    r2.cells[4].text = '{{ collection_date }}\n{{ collection_time }}'
    r2.cells[5].text = 'HANDED OVER DEVICES/HDD AFTER IMAGING'
    
    # Row 3: Transfer 2 (Handing over after conclusion of search)
    r3 = t1.rows[3]
    r3.cells[0].text = '{{ transfer_2_devices }}'
    r3.cells[1].text = 'HANDING OVER OF HDD WITH DATA AFTER FORENSIC BACKUP AND HASHING AFTER CONCLUSION OF SEARCH'
    r3.cells[2].text = 'Forensic Expert (Name & Sign)\n\n{{ expert_name }}'
    r3.cells[3].text = 'Authorised officer (Name & Sign)\n\n{{ authorised_officer }}'
    r3.cells[4].text = ''
    r3.cells[5].text = 'HANDED OVER DEVICES/HDD AFTER CONCLUSION OF SEARCH TO CUSTODIAN'
    
    # Apply cell styling and widths across all rows of t1
    for r_i, row in enumerate(t1.rows):
        for c_i, cell in enumerate(row.cells):
            cell.width = col_widths[c_i]
            set_cell_borders(cell, color='999999', sz='4')
            if r_i >= 2:
                set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
                        r.font.name = 'Calibri'
                        
    out_path = os.path.join(OUT_DIR, 'mumbai_annexure_3_coc.docx')
    doc.save(out_path)
    print("Generated:", out_path)

# -------------------------------------------------------------
# 3. Mumbai Non-Mobile Evidence Form
# -------------------------------------------------------------
def build_mumbai_non_mobile():
    src = os.path.join(DOCS_DIR, 'DIGITAL EVIDENCE COLLECTION FORM FOR DEVICES EXCEPT MOBILE DEVICES.docx')
    doc = docx.Document(src)
    
    t0 = doc.tables[0]
    t0.rows[0].cells[0].text = 'NAME OF AUTHORISED OFFICER: {{ authorised_officer }}'
    t0.rows[1].cells[0].text = 'NAME OF CONCERNED PERSON/PARTY: {{ party_name }}'
    t0.rows[2].cells[0].text = 'OWNER/USER NAME: {{ owner_user }}'
    
    t0.rows[3].cells[0].text = 'DATE'
    t0.rows[3].cells[1].text = 'TIME'
    t0.rows[3].cells[4].text = 'ADDRESS'
    
    t0.rows[4].cells[0].text = '{{ collection_date }}'
    t0.rows[4].cells[1].text = '{{ collection_time }}'
    t0.rows[4].cells[4].text = '{{ collection_address }}'
    
    t0.rows[5].cells[0].text = "EXAMINER'S NAMES AND DETAILS"
    t0.rows[6].cells[0].text = '{{ examiner_name }} - {{ examiner_details }}'
    
    t0.rows[7].cells[0].text = 'LAPTOP'
    t0.rows[7].cells[2].text = 'DESKTOP'
    t0.rows[7].cells[5].text = 'SERVER'
    t0.rows[7].cells[7].text = 'FILE/FOLDER'
    
    t0.rows[8].cells[0].text = '{{ is_laptop_sym }}'
    t0.rows[8].cells[2].text = '{{ is_desktop_sym }}'
    t0.rows[8].cells[5].text = '{{ is_server_sym }}'
    t0.rows[8].cells[7].text = '{{ is_filefolder_sym }}'
    
    t0.rows[9].cells[0].text = 'OTHERS {{ is_other_sym }}'
    t0.rows[9].cells[2].text = 'IF OTHERS SPECIFY: {{ non_mobile_other_specify }}'
    
    t0.rows[10].cells[0].text = 'HIBERNATION/SLEEP'
    t0.rows[10].cells[3].text = 'IF SWITCHED ON WHAT IS VISIBLE ON SCREEN?'
    t0.rows[11].cells[0].text = 'ON {{ hib_on_sym }}    OFF {{ hib_off_sym }}'
    t0.rows[11].cells[3].text = '{{ visible_on_screen }}'
    
    t0.rows[12].cells[0].text = 'DEVICE INFO'
    t0.rows[12].cells[3].text = (
        'MAKE: {{ make }}\n'
        'MODEL: {{ model }}\n'
        'SR.NO: {{ serial_number }}\n'
        'STORAGE CAPACITY: {{ capacity }}\n'
        'Mail ID: {{ email_id }}'
    )
    
    t0.rows[13].cells[0].text = 'SHUT DOWN TYPE'
    t0.rows[13].cells[3].text = 'NORMAL'
    t0.rows[13].cells[6].text = 'POWER PLUG PULLED'
    t0.rows[13].cells[8].text = 'BATTERY REMOVED'
    
    t0.rows[14].cells[3].text = '{{ shut_normal_sym }}'
    t0.rows[14].cells[6].text = '{{ shut_plug_sym }}'
    t0.rows[14].cells[8].text = '{{ shut_battery_sym }}'
    
    t0.rows[15].cells[0].text = 'ENCRYPTION PRESENT?'
    t0.rows[15].cells[3].text = 'ENCRYPTION SOFTWARE USED'
    t0.rows[16].cells[0].text = 'YES {{ enc_yes_sym }}    NO {{ enc_no_sym }}'
    t0.rows[16].cells[3].text = '{{ encryption_software }}'
    
    t0.rows[17].cells[0].text = 'DEVICE HANDLING'
    t0.rows[17].cells[3].text = 'SEIZURE {{ handle_seizure_sym }}'
    t0.rows[17].cells[6].text = 'BACKUP {{ handle_backup_sym }}'
    t0.rows[17].cells[8].text = 'REPORT {{ handle_reporting_sym }}'
    
    t0.rows[18].cells[0].text = 'IMAGING TOOLS AND DETAILS: {{ imaging_tools_details }}'
    t0.rows[19].cells[0].text = 'MD5 HASH VALUE: {{ md5_hash }}'
    t0.rows[20].cells[0].text = 'SHA1 HASH VALUE: {{ sha1_hash }}'
    t0.rows[21].cells[0].text = 'SHA256 HASH VALUE1: {{ sha256_hash_1 }}'
    
    t1 = doc.tables[1]
    t1.rows[0].cells[0].text = 'MASTER COPY'
    t1.rows[0].cells[2].text = 'WORKING COPY'
    t1.rows[1].cells[0].text = 'MAKE: {{ master_copy_make }}'
    t1.rows[1].cells[1].text = 'MODEL: {{ master_copy_model }}'
    t1.rows[1].cells[2].text = 'MAKE: {{ working_copy_make }}'
    t1.rows[1].cells[3].text = 'MODEL: {{ working_copy_model }}'
    t1.rows[2].cells[0].text = 'S/N: {{ master_copy_sn }}'
    t1.rows[2].cells[1].text = 'CAPACITY: {{ master_copy_capacity }}'
    t1.rows[2].cells[2].text = 'S/N: {{ working_copy_sn }}'
    t1.rows[2].cells[3].text = 'CAPACITY: {{ working_copy_capacity }}'
    
    t2 = doc.tables[2]
    t2.rows[0].cells[0].text = 'NOTES BY AUTHORISED OFFICER'
    t2.rows[1].cells[0].text = ''
    
    out_path = os.path.join(OUT_DIR, 'mumbai_device_non_mobile.docx')
    doc.save(out_path)
    print("Generated:", out_path)

# -------------------------------------------------------------
# 4. Mumbai Mobile Evidence Form
# -------------------------------------------------------------
def build_mumbai_mobile():
    src = os.path.join(DOCS_DIR, 'DEVICE COLLECTION FORM FOR MOBILE DEVICES.docx')
    doc = docx.Document(src)
    
    # Keep Table 0, 1, 2 only
    while len(doc.tables) > 3:
        tbl = doc.tables[-1]._tbl
        tbl.getparent().remove(tbl)
        
    for p in list(doc.paragraphs):
        txt = p.text.strip()
        if txt == 'DEVICE COLLECTION FORM FOR MOBILE DEVICES':
            continue
        p._p.getparent().remove(p._p)
        
    t0 = doc.tables[0]
    t0.rows[0].cells[0].text = 'NAME OF AUTHORISED OFFICER: {{ authorised_officer }}'
    t0.rows[1].cells[0].text = 'NAME OF CONCERNED PERSON/PARTY: {{ party_name }}'
    t0.rows[2].cells[0].text = 'OWNER/USER NAME: {{ owner_user }}'
    
    t0.rows[3].cells[0].text = 'DATE'
    t0.rows[3].cells[1].text = 'TIME'
    t0.rows[3].cells[4].text = 'ADDRESS'
    
    t0.rows[4].cells[0].text = '{{ collection_date }}'
    t0.rows[4].cells[1].text = '{{ collection_time }}'
    t0.rows[4].cells[4].text = '{{ collection_address }}'
    
    t0.rows[5].cells[0].text = "EXAMINER'S NAMES AND DETAILS"
    t0.rows[6].cells[0].text = '{{ examiner_name }} - {{ examiner_details }}'
    
    t0.rows[7].cells[0].text = 'HIBERNATION/SLEEP'
    t0.rows[7].cells[3].text = 'IF SWITCHED ON WHAT IS VISIBLE ON SCREEN?'
    t0.rows[8].cells[0].text = 'ON {{ hib_on_sym }}    OFF {{ hib_off_sym }}'
    t0.rows[8].cells[3].text = '{{ visible_on_screen }}'
    
    t0.rows[9].cells[0].text = 'DEVICE INFO'
    t0.rows[9].cells[3].text = (
        'MAKE: {{ make }}\n'
        'MODEL: {{ model }}\n'
        'SR.NO: {{ serial_number }}\n'
        'STORAGE CAPACITY: {{ capacity }}\n'
        'OS: {{ os }}\n'
        'IMEI 1: {{ imei1 }}'
    )
    t0.rows[9].cells[5].text = 'IMEI 2: {{ imei2 }}'
    
    t0.rows[10].cells[0].text = 'SIM CARD PRESENT?'
    t0.rows[10].cells[3].text = 'SIM CARD DETAILS'
    t0.rows[11].cells[0].text = 'YES {{ sim_yes_sym }}    NO {{ sim_no_sym }}'
    t0.rows[11].cells[3].text = 'PROVIDER: {{ sim_provider }}\tIMSI: {{ imsi }}'
    
    t0.rows[12].cells[0].text = 'MEDIA CARD PRESENT?'
    t0.rows[12].cells[4].text = 'MEDIA CARD DETAILS'
    t0.rows[13].cells[0].text = 'YES {{ media_yes_sym }}    NO {{ media_no_sym }}'
    t0.rows[13].cells[4].text = '{{ media_card_details }}'
    
    t0.rows[14].cells[0].text = 'MOBILE PHONE STATE'
    t0.rows[14].cells[3].text = 'ON {{ mob_state_on_sym }}'
    t0.rows[14].cells[5].text = 'OFF {{ mob_state_off_sym }}'
    t0.rows[14].cells[7].text = 'OFFLINE {{ mob_state_offline_sym }}'
    
    t0.rows[15].cells[0].text = 'DEVICE HANDLING'
    t0.rows[15].cells[3].text = 'SEIZURE {{ handle_seizure_sym }}'
    t0.rows[15].cells[5].text = 'BACKUP {{ handle_backup_sym }}'
    t0.rows[15].cells[7].text = 'REPORT {{ handle_reporting_sym }}'
    
    t0.rows[16].cells[0].text = 'IMAGING TOOLS AND DETAILS: {{ imaging_tools_details }}'
    t0.rows[17].cells[0].text = 'MD5 HASH VALUE: {{ md5_hash }}'
    t0.rows[18].cells[0].text = 'SHA1 HASH VALUE: {{ sha1_hash }}'
    t0.rows[19].cells[0].text = 'SHA256 HASH VALUE1: {{ sha256_hash_1 }}'
    t0.rows[20].cells[0].text = 'SHA256 HASH VALUE2: {{ sha256_hash_2 }}'
    t0.rows[21].cells[0].text = 'SHA256 HASH VALUE3: {{ sha256_hash_3 }}'
    t0.rows[22].cells[0].text = 'SHA256 HASH VALUE4: {{ sha256_hash_4 }}'
    t0.rows[23].cells[0].text = 'SHA256 HASH VALUE5: {{ sha256_hash_5 }}'
    
    t1 = doc.tables[1]
    t1.rows[0].cells[0].text = 'MASTER COPY (COPY 1)'
    t1.rows[0].cells[2].text = 'WORKING COPY (COPY 2)'
    t1.rows[1].cells[0].text = 'MAKE: {{ master_copy_make }}'
    t1.rows[1].cells[1].text = 'MODEL: {{ master_copy_model }}'
    t1.rows[1].cells[2].text = 'MAKE: {{ working_copy_make }}'
    t1.rows[1].cells[3].text = 'MODEL: {{ working_copy_model }}'
    t1.rows[2].cells[0].text = 'S/N: {{ master_copy_sn }}'
    t1.rows[2].cells[1].text = 'CAPACITY: {{ master_copy_capacity }}'
    t1.rows[2].cells[2].text = 'S/N: {{ working_copy_sn }}'
    t1.rows[2].cells[3].text = 'CAPACITY: {{ working_copy_capacity }}'
    
    t2 = doc.tables[2]
    t2.rows[0].cells[0].text = 'NOTES BY AUTHORISED OFFICER'
    t2.rows[1].cells[0].text = ''
    
    out_path = os.path.join(OUT_DIR, 'mumbai_device_mobile.docx')
    doc.save(out_path)
    print("Generated:", out_path)

# -------------------------------------------------------------
# 5. Mumbai Per Device Evidence (BSA 2023 Sec 63(4)(c))
# -------------------------------------------------------------
def build_mumbai_per_device_bsa():
    src = os.path.join(DOCS_DIR, 'Per device Evidence.docx')
    doc = docx.Document(src)
    
    paras_to_remove = []
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        # Part A party intro
        if 'I,' in t and '(Name)' in t and 'residing/employed at' in t and i < 30:
            p.text = (
                "I, {{ party_name }} (Name), Son/daughter/spouse of {{ party_relation }} "
                "residing/employed at {{ party_residing_at }} do hereby solemnly affirm and sincerely state and submit as follows: —"
            )
        elif 'Computer / Storage Media' in t and 'Flash Drive' in t:
            p.text = "Computer / Storage Media {{ cb_computer }}   DVR {{ cb_dvr }}   Mobile {{ cb_mobile }}   Flash Drive {{ cb_flash }}"
        elif 'CD/DVD' in t and 'Server' in t and 'Cloud' in t:
            p.text = "CD/DVD {{ cb_cddvd }}\tServer {{ cb_server }}\tCloud {{ cb_cloud }}\tOther {{ cb_other }}"
        elif t.startswith('Other:') and len(t) < 15:
            p.text = "Other: {{ annexure_e_other_specify }}"
        elif t.startswith('Make & Model:'):
            p.text = "Make & Model: {{ make_model }}"
        elif t.startswith('Serial Number:'):
            p.text = "Serial Number: {{ serial_number }}"
        elif t.startswith('Color:'):
            p.text = "Color: {{ color }}"
        elif 'IMEI/UIN/UID/MAC/Cloud ID' in t:
            p.text = "IMEI/UIN/UID/MAC/Cloud ID: {{ unique_id }} (as applicable)"
        elif 'and any other relevant information' in t:
            if 'I state that the HASH value' in t:
                p.text = (
                    "and any other relevant information, if any, about the device/digital record: {{ other_specific_notes }}. "
                    "I state that the HASH value/s of the electronic/digital record/s is {{ hash_value }}, obtained through the following algorithm: —"
                )
            else:
                p.text = "and any other relevant information, if any, about the device/digital record: {{ other_specific_notes }} (specify)."
        elif 'Owned' in t and 'Maintained' in t and 'Managed' in t and 'Operated' in t:
            p.text = "Owned {{ ctrl_owned }}   Maintained {{ ctrl_maintained }}   Managed {{ ctrl_managed }}   Operated {{ ctrl_operated }}"
        elif 'I state that the HASH value/s' in t and i < 35:
            p.text = "I state that the HASH value/s of the electronic/digital record/s is {{ hash_value }}, obtained through the following algorithm: —"
        elif t == 'SHA1:':
            p.text = "SHA1: {{ hash_sha1 }}"
        elif t == 'SHA256:':
            p.text = "SHA256: {{ hash_sha256 }}"
        elif t == 'MD5:':
            p.text = "MD5: {{ hash_md5 }}"
        elif t.startswith('Other') and 'Legally acceptable standard' in t:
            p.text = "Other {{ hash_other }} (Legally acceptable standard) (Hash report to be enclosed with the certificate)"
        elif t == '(Name and signature)':
            p.text = "{{ party_name }} (Name and signature)"
        elif t == '(Name, designation and signature)':
            p.text = "{{ expert_name }}, {{ expert_designation }} (Name, designation and signature)"
        elif t.startswith('Date (DD/MM/YYYY):'):
            p.text = "Date (DD/MM/YYYY): {{ collection_date }}"
        elif 'Time (IST):' in t and 'Place:' in t:
            p.text = "Time (IST): {{ collection_time }} hours (In 24 hours format)    Place: {{ cert_place }}"
        elif 'I,' in t and '(Name)' in t and i > 35:
            # Part B expert intro line 1 - Expert Father Name and affirmation
            p.text = "I, {{ expert_name }} (Name), Son/daughter/spouse of {{ expert_father_name }} do hereby solemnly affirm and sincerely state and submit as follows: —"
        elif t.startswith('residing/employed at') and i > 35:
            # Part B expert intro line 2 - Remove company name / residing line completely
            paras_to_remove.append(p)
            
    for p in paras_to_remove:
        p._p.getparent().remove(p._p)
            
    out_path = os.path.join(OUT_DIR, 'mumbai_per_device_bsa.docx')
    doc.save(out_path)
    print("Generated:", out_path)

if __name__ == '__main__':
    print("Building Mumbai Templates...")
    build_mumbai_annexure_1()
    build_mumbai_annexure_3()
    build_mumbai_non_mobile()
    build_mumbai_mobile()
    build_mumbai_per_device_bsa()
    print("All Mumbai templates built successfully!")
