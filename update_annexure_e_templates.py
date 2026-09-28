import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), 'templates_docx')

def remove_table_borders_and_shading(table):
    """Removes all borders and background shading from a table to make it borderless."""
    tblPr = table._tbl.tblPr
    # Remove existing tblBorders if present
    for child in list(tblPr):
        if child.tag.endswith('tblBorders'):
            tblPr.remove(child)
    no_bdr = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none"/>'
        '<w:left w:val="none"/>'
        '<w:bottom w:val="none"/>'
        '<w:right w:val="none"/>'
        '<w:insideH w:val="none"/>'
        '<w:insideV w:val="none"/>'
        '</w:tblBorders>'
    )
    tblPr.append(no_bdr)
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for child in list(tcPr):
                if child.tag.endswith('tcBorders') or child.tag.endswith('shd'):
                    tcPr.remove(child)

def add_paragraph_box_border(p, color="7F7F7F", sz="4"):
    """Adds a neat rectangular border around a paragraph without using a table."""
    pPr = p._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith('pBdr'):
            pPr.remove(child)
    bdr_xml = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="6" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="6" w:color="{color}"/>'
        '</w:pBdr>'
    )
    pPr.append(bdr_xml)

def update_batch_annexure_e(filename):
    file_path = os.path.join(TEMPLATES_DIR, filename)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    
    doc = docx.Document(file_path)
    
    # 1. Update Table 1 and Table 5 (4 Things: Make & Model, Serial Number, Color, IMEI/Cloud ID) to be borderless
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        remove_table_borders_and_shading(t1)
        # Ensure row order is: Make & Model, Serial Number, Color, IMEI/Cloud ID
        if len(t1.rows) >= 4:
            t1.rows[0].cells[0].paragraphs[0].text = "Make & Model:"
            t1.rows[0].cells[1].paragraphs[0].text = "{{ ev.make_model }}"
            t1.rows[1].cells[0].paragraphs[0].text = "Serial Number:"
            t1.rows[1].cells[1].paragraphs[0].text = "{{ ev.serial_number }}"
            t1.rows[2].cells[0].paragraphs[0].text = "Color:"
            t1.rows[2].cells[1].paragraphs[0].text = "{{ ev.color }}"
            t1.rows[3].cells[0].paragraphs[0].text = "IMEI/UIN/UID/MAC/Cloud ID: (as applicable)"
            t1.rows[3].cells[1].paragraphs[0].text = "{{ ev.unique_id }}"
            
    if len(doc.tables) > 5:
        t5 = doc.tables[5]
        remove_table_borders_and_shading(t5)
        if len(t5.rows) >= 4:
            t5.rows[0].cells[0].paragraphs[0].text = "Make & Model:"
            t5.rows[0].cells[1].paragraphs[0].text = "{{ ev.make_model }}"
            t5.rows[1].cells[0].paragraphs[0].text = "Serial Number:"
            t5.rows[1].cells[1].paragraphs[0].text = "{{ ev.serial_number }}"
            t5.rows[2].cells[0].paragraphs[0].text = "Color:"
            t5.rows[2].cells[1].paragraphs[0].text = "{{ ev.color }}"
            t5.rows[3].cells[0].paragraphs[0].text = "IMEI/UIN/UID/MAC/Cloud ID: (as applicable)"
            t5.rows[3].cells[1].paragraphs[0].text = "{{ ev.unique_id }}"

    # 2. Update Paragraphs
    for p in doc.paragraphs:
        txt = p.text.strip()
        # Part A Intro
        if 'I,' in txt and 'party_name' in txt and 'residing/employed at' in txt:
            p.text = (
                "I, {{ ev.party_name }}, Son/daughter/spouse of {{ ev.party_father_name }}, "
                "residing/employed at {{ ev.party_residing_at }} do hereby solemnly affirm and sincerely state and submit as follows:"
            )
        # Part B Intro
        elif 'I,' in txt and 'expert_name' in txt and 'do hereby solemnly affirm' in txt:
            p.text = (
                "I, {{ ev.expert_name }}, Son/daughter/spouse of {{ ev.expert_father_name }}, "
                "residing/employed at R S Patel & Co. do hereby solemnly affirm and sincerely state and submit as follows:"
            )
        # Remove notes_by_officer
        elif 'notes_by_officer' in txt:
            p.text = "and any other relevant information, if any, about the device/digital record: {{ ev.other_specific_notes }}."
            
        # Hash display with checkbox BEFORE text and box border without table
        elif 'SHA1{{' in txt or ('SHA1' in txt and 'SHA256' in txt and 'MD5' in txt):
            # Checkbox before text
            p.text = (
                "{{ ev.cb_sha1 }} SHA1: {{ ev.sha1_hash }}\n"
                "{{ ev.cb_sha256 }} SHA256: {{ ev.sha256_hash_1 }}\n"
                "{{ ev.cb_md5 }} MD5: {{ ev.md5_hash }}\n"
                "{{ ev.cb_other_hash }} Other {{ ev.hash_other }} (Legally acceptable standard) (Hash report to be enclosed with the certificate)"
            )
            add_paragraph_box_border(p, color="A6A6A6", sz="6")
            
    doc.save(file_path)
    print(f"Updated {filename} successfully.")

def update_mumbai_per_device_bsa():
    file_path = os.path.join(TEMPLATES_DIR, 'mumbai_per_device_bsa.docx')
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
        
    doc = docx.Document(file_path)
    
    for p in doc.paragraphs:
        txt = p.text.strip()
        # Part A Intro
        if 'I,' in txt and 'party_name' in txt and 'residing/employed at' in txt:
            p.text = (
                "I, {{ party_name }} (Name), Son/daughter/spouse of {{ party_father_name }} "
                "residing/employed at {{ party_residing_at }} do hereby solemnly affirm and sincerely state and submit as follows: —"
            )
        # Part B Intro
        elif 'I,' in txt and 'expert_name' in txt and 'do hereby solemnly affirm' in txt:
            p.text = (
                "I, {{ expert_name }} (Name), Son/daughter/spouse of {{ expert_father_name }}, "
                "residing/employed at R S Patel & Co. do hereby solemnly affirm and sincerely state and submit as follows: —"
            )
        # Hash section with checkbox BEFORE text and border box without table
        elif txt == 'SHA1: {{ hash_sha1 }}':
            p.text = (
                "{{ cb_sha1 }} SHA1: {{ sha1_hash }}\n"
                "{{ cb_sha256 }} SHA256: {{ sha256_hash_1 }}\n"
                "{{ cb_md5 }} MD5: {{ md5_hash }}\n"
                "{{ cb_other_hash }} Other {{ hash_other }} (Legally acceptable standard) (Hash report to be enclosed with the certificate)"
            )
            add_paragraph_box_border(p, color="A6A6A6", sz="6")
        elif txt in ['SHA256: {{ hash_sha256 }}', 'MD5: {{ hash_md5 }}'] or txt.startswith('Other {{ hash_other }}'):
            p.text = ""  # blank out since combined in the first bordered paragraph

    doc.save(file_path)
    print("Updated mumbai_per_device_bsa.docx successfully.")

if __name__ == '__main__':
    update_batch_annexure_e('annexure_e_Gujarat.docx')
    update_batch_annexure_e('annexure_e_delhi.docx')
    update_batch_annexure_e('annexure_e.docx')
    update_mumbai_per_device_bsa()
    print("All Annexure E templates successfully updated!")
