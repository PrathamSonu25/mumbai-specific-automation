import os
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT

SRC_DIR = r'C:\Users\Admin\Desktop\Example\Documents'
OUT_DIR = 'templates_docx'
os.makedirs(OUT_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. Annexure B: Device Collection for Device other than mobile
# -------------------------------------------------------------
def build_annexure_b():
    src = os.path.join(SRC_DIR, '2. Annexure B- Device Collection form for Device other than mobile.docx')
    doc = docx.Document(src)
    t = doc.tables[0]
    
    t.rows[1].cells[0].text = 'NAME OF AUTHORISED OFFICER: {{ authorised_officer }}\nNAME OF ASSESSEE/PARTY: {{ assessee_name }}        NAME OF DEVICE: {{ device_name }}'
    t.rows[3].cells[0].text = '{{ collection_date }}'
    t.rows[3].cells[1].text = '{{ collection_time }}'
    t.rows[3].cells[2].text = '{{ collection_address }}'
    t.rows[6].cells[0].text = '{{ examiner_name }} - {{ examiner_details }}'
    t.rows[9].cells[0].text = '{{ is_laptop_sym }}'
    t.rows[9].cells[1].text = '{{ is_desktop_sym }}'
    t.rows[9].cells[2].text = '{{ is_server_sym }}'
    t.rows[9].cells[3].text = '{{ is_filefolder_sym }}'
    t.rows[10].cells[0].text = 'OTHERS {{ is_other_sym }}'
    t.rows[10].cells[1].text = 'IF OTHERS SPECIFY: {{ other_specify }}'
    t.rows[12].cells[0].text = 'ON {{ hib_on_sym }}    OFF {{ hib_off_sym }}'
    t.rows[12].cells[1].text = '{{ visible_on_screen }}'
    t.rows[13].cells[1].text = 'MAKE: {{ make }}    MODEL: {{ model }}\nS/N: {{ serial_number }}    CAPACITY: {{ capacity }}'
    t.rows[15].cells[1].text = '{{ shut_normal_sym }}'
    t.rows[15].cells[2].text = '{{ shut_plug_sym }}'
    t.rows[15].cells[3].text = '{{ shut_battery_sym }}'
    t.rows[17].cells[0].text = 'YES {{ enc_yes_sym }}    NO {{ enc_no_sym }}'
    t.rows[17].cells[1].text = '{{ encryption_software }}'
    t.rows[18].cells[1].text = 'SEIZURE {{ handle_seizure_sym }}'
    t.rows[18].cells[2].text = 'IMAGING {{ handle_imaging_sym }}'
    t.rows[18].cells[3].text = 'BACKUP {{ handle_backup_sym }}'
    t.rows[19].cells[0].text = 'IMAGING TOOLS AND DETAILS: {{ imaging_tools_details }}'
    t.rows[20].cells[0].text = 'MD5 HASH VALUE: {{ md5_hash }}        SHA1 HASH VALUE: {{ sha1_hash }}\nSHA 256 HASH VALUE: {{ sha256_hash_1 }}'
    t.rows[22].cells[0].text = 'MAKE: {{ master_copy_make }}    MODEL: {{ master_copy_model }}\nS/N: {{ master_copy_sn }}    CAPACITY: {{ master_copy_capacity }}'
    t.rows[22].cells[2].text = 'MAKE: {{ working_copy_make }}    MODEL: {{ working_copy_model }}\nS/N: {{ working_copy_sn }}    CAPACITY: {{ working_copy_capacity }}'
    t.rows[24].cells[0].text = 'YES {{ returned_yes_sym }}'
    t.rows[24].cells[1].text = 'NO {{ returned_no_sym }}'
    t.rows[24].cells[2].text = '{{ returned_date }}'
    t.rows[24].cells[3].text = '{{ returned_time }}'
    t.rows[26].cells[2].text = 'YES {{ witness_yes_sym }}'
    t.rows[26].cells[3].text = 'NO {{ witness_no_sym }}'
    t.rows[28].cells[0].text = ''
    
    out_path = os.path.join(OUT_DIR, 'annexure_b.docx')
    doc.save(out_path)
    print('Generated:', out_path)

# -------------------------------------------------------------
# 2. Annexure C: Chain of Custody
# -------------------------------------------------------------
def build_annexure_c():
    src = os.path.join(SRC_DIR, '3. Annexure C- Chain of custody.docx')
    doc = docx.Document(src)
    t = doc.tables[0]
    
    t.rows[1].cells[0].text = 'NAME OF ASSESSEE/PARTY: {{ assessee_name }}\nIn the case of: {{ case_name }}'
    t.rows[3].cells[0].text = '{{ collection_date }}'
    t.rows[3].cells[1].text = '{{ collection_time }}'
    t.rows[3].cells[2].text = '{{ collection_address }}'
    
    t.rows[4].cells[0].text = (
        'DEVICE NAME: {{ device_name }}\n'
        'MAIN EVIDENCE DETAILS:\n'
        'MAKE: {{ make }}    MODEL: {{ model }}\n'
        'S/N: {{ serial_number }}    CAPACITY: {{ capacity }}    OWNER/USER: {{ owner_user }}\n'
        'IMEI1: {{ imei1 }}    IMEI2: {{ imei2 }} (Applicable for Mobile Device)\n\n'
        'SUB EVIDENCE DETAILS:(If Applicable)\n'
        '{% for sub in sub_evidences %}'
        'MAKE: {{ sub.make }}    MODEL: {{ sub.model }}    S/N: {{ sub.sn }}    CAPACITY: {{ sub.capacity }}    OWNER/USER: {{ sub.owner_user }}\n'
        '{% endfor %}'
    )
    
    # Remove rows 7 to end
    for r_idx in range(len(t.rows) - 1, 6, -1):
        tr = t.rows[r_idx]._tr
        tr.getparent().remove(tr)
        
    r1 = t.add_row()
    r1.cells[0].text = '{%tr for c in custody_rows %}'
    
    r2 = t.add_row()
    r2.cells[0].text = '{{ c.reason_action }}'
    r2.cells[1].text = 'NAME: {{ c.given_by_name }}\nSIGNATURE: {{ c.given_by_signature }}'
    r2.cells[2].text = 'NAME: {{ c.received_by_name }}\nSIGNATURE: {{ c.received_by_signature }}'
    r2.cells[3].text = '{{ c.date }}'
    r2.cells[4].text = '{{ c.time }}'
    
    r3 = t.add_row()
    r3.cells[0].text = '{%tr endfor %}'
    
    out_path = os.path.join(OUT_DIR, 'annexure_c.docx')
    doc.save(out_path)
    print('Generated:', out_path)

# -------------------------------------------------------------
# 3. Annexure D: Device Collection for Mobile Phones
# -------------------------------------------------------------
def build_annexure_d():
    src = os.path.join(SRC_DIR, '4. Annexure D- Device Collection form for mobile.docx')
    doc = docx.Document(src)
    t = doc.tables[0]
    
    t.rows[1].cells[0].text = 'NAME OF AUTHORISED OFFICER: {{ authorised_officer }}'
    t.rows[2].cells[0].text = 'NAME OF ASSESSEE/PARTY: {{ assessee_name }}'
    t.rows[3].cells[0].text = 'NAME OF DEVICE: {{ device_name }}'
    t.rows[5].cells[0].text = '{{ collection_date }}'
    t.rows[5].cells[1].text = '{{ collection_time }}'
    t.rows[5].cells[2].text = '{{ collection_address }}'
    t.rows[7].cells[0].text = '{{ examiner_name }} - {{ examiner_details }}'
    t.rows[9].cells[0].text = 'ON {{ hib_on_sym }}    OFF {{ hib_off_sym }}'
    t.rows[9].cells[1].text = '{{ visible_on_screen }}'
    t.rows[10].cells[1].text = 'MAKE: {{ make }}    MODEL: {{ model }}    OS: {{ os }}\nS/N: {{ serial_number }}    STORAGE CAPACITY: {{ capacity }}'
    t.rows[11].cells[1].text = 'IMEI1: {{ imei1 }}    IMEI2: {{ imei2 }}'
    t.rows[13].cells[0].text = 'YES {{ sim_yes_sym }}    NO {{ sim_no_sym }}'
    t.rows[13].cells[1].text = 'SIM CARD PROVIDER: {{ sim_provider }}    IMSI: {{ imsi }}'
    t.rows[15].cells[0].text = 'YES {{ media_yes_sym }}    NO {{ media_no_sym }}'
    t.rows[15].cells[1].text = 'MAKE: {{ media_card_make }}    MODEL: {{ media_card_model }}    S/N: {{ media_card_sn }}'
    t.rows[16].cells[1].text = 'ON {{ state_on_sym }}'
    t.rows[16].cells[2].text = 'OFF {{ state_off_sym }}'
    t.rows[16].cells[3].text = 'OFFLINE {{ state_offline_sym }}'
    t.rows[17].cells[1].text = 'SEIZURE {{ handle_seizure_sym }}'
    t.rows[17].cells[2].text = 'BACKUP {{ handle_backup_sym }}'
    t.rows[17].cells[3].text = 'REPORTING {{ handle_reporting_sym }}'
    t.rows[18].cells[0].text = 'IMAGING TOOLS AND DETAILS: {{ imaging_tools_details }}'
    t.rows[19].cells[0].text = (
        'MD5 HASH VALUE: {{ md5_hash }}\n'
        'SHA 256 HASH VALUE1: {{ sha256_hash_1 }}    SHA 256 HASH VALUE2: {{ sha256_hash_2 }}\n'
        'SHA 256 HASH VALUE3: {{ sha256_hash_3 }}    SHA 256 HASH VALUE4: {{ sha256_hash_4 }}\n'
        'SHA 256 HASH VALUE5: {{ sha256_hash_5 }}'
    )
    t.rows[21].cells[0].text = 'MAKE: {{ master_copy_make }}    MODEL: {{ master_copy_model }}\nS/N: {{ master_copy_sn }}    CAPACITY: {{ master_copy_capacity }}'
    t.rows[21].cells[2].text = 'MAKE: {{ working_copy_make }}    MODEL: {{ working_copy_model }}\nS/N: {{ working_copy_sn }}    CAPACITY: {{ working_copy_capacity }}'
    t.rows[23].cells[0].text = ''
    
    out_path = os.path.join(OUT_DIR, 'annexure_d.docx')
    doc.save(out_path)
    print('Generated:', out_path)

# -------------------------------------------------------------
# 4. Annexure E: BSA 2023 Section 63(4)(c) Certificate
# -------------------------------------------------------------
def build_annexure_e():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    doc.add_paragraph('{% for ev in evidences %}')
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.add_run('CERTIFICATE U/S 63(4)(C) OF THE BHARATIYA SAKSHYA ADHINIYAM (BSA), 2023')
    r.bold = True
    r.font.size = Pt(13)
    
    # --- PART A ---
    p_part_a = doc.add_paragraph()
    p_part_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_part_a.add_run('PART A\n(To be filled by the Party)')
    r.bold = True
    r.font.size = Pt(11)
    
    p = doc.add_paragraph()
    p.add_run('I, ').font.size = Pt(10)
    p.add_run('{{ ev.party_name }}').bold = True
    p.add_run(' (Name), ')
    p.add_run('{{ ev.party_relation }}')
    p.add_run(' residing/employed at ')
    p.add_run('{{ ev.party_residing_at }}')
    p.add_run(' do hereby solemnly affirm and sincerely state and submit as follows: -')
    
    p = doc.add_paragraph('I have produced electronic record/output of the digital record taken from the following device/digital record source (tick Mark): -')
    p.runs[0].font.size = Pt(10)
    
    p = doc.add_paragraph()
    p.add_run('{{ ev.cb_computer }} Computer/Storage Media    {{ ev.cb_cddvd }} CD/DVD    {{ ev.cb_server }} Server')
    p = doc.add_paragraph()
    p.add_run('{{ ev.cb_dvr }} DVR    {{ ev.cb_cloud }} Cloud    {{ ev.cb_mobile }} Mobile    {{ ev.cb_flash }} Flash Drive    {{ ev.cb_other }} Other: {{ ev.annexure_e_other_specify }}')
    
    p = doc.add_paragraph()
    p.add_run('Make & Model: ').bold = True
    p.add_run('{{ ev.make_model }}\n')
    p.add_run('Serial Number: ').bold = True
    p.add_run('{{ ev.serial_number }}\n')
    p.add_run('Color: ').bold = True
    p.add_run('{{ ev.color }}\n')
    p.add_run('IMEI/UIN/UID/MAC/Cloud ID: ').bold = True
    p.add_run('{{ ev.unique_id }} (as applicable)\n')
    p.add_run('and any other relevant information, if any, about the device/digital record: ').bold = True
    p.add_run('{{ ev.other_specific_notes }} (specify).')
    
    p = doc.add_paragraph(
        'The digital device or the digital record source was under the lawful control for regularly creating, '
        'storing or processing information for the purposes of carrying out regular activities and during this period, '
        'the computer or the communication device was working properly and the relevant information was regularly fed into '
        'the computer during the ordinary course of business. If the computer/digital device at any point of time was not '
        'working properly or out of operation, then it has not affected the electronic/digital record or its accuracy. '
        'The digital device or the source of the digital record is: -'
    )
    p.runs[0].font.size = Pt(10)
    
    p = doc.add_paragraph()
    p.add_run('{{ ev.ctrl_owned }} Owned    {{ ev.ctrl_maintained }} Maintained    {{ ev.ctrl_managed }} Managed    {{ ev.ctrl_operated }} Operated    by me (select as applicable)')
    
    p = doc.add_paragraph()
    p.add_run('I state that the Hash value(s) of the electronic/digital record(s) is ').font.size = Pt(10)
    p.add_run('{{ ev.hash_value }}').bold = True
    p.add_run(', obtained through the following algorithm: -')
    
    p = doc.add_paragraph()
    p.add_run('{{ ev.hash_sha1 }} SHA1    {{ ev.hash_sha256 }} SHA256    {{ ev.hash_md5 }} MD5    {{ ev.hash_other }} Other (Legally acceptable standard)')
    doc.add_paragraph('(Hash report to be enclosed with the certificate)')
    
    p = doc.add_paragraph()
    p.add_run('Date (DD/MM/YYYY): {{ ev.collection_date }}        Time (IST): {{ ev.collection_time }} hours        Place: {{ ev.cert_place }}\n')
    p.add_run('(Name and Signature)\nName: ').bold = True
    p.add_run('{{ ev.party_name }}')
    
    # --- PART B ---
    doc.add_paragraph()
    p_part_b = doc.add_paragraph()
    p_part_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_part_b.add_run('PART B\n(To be filled by the Expert)')
    r.bold = True
    r.font.size = Pt(11)
    
    p = doc.add_paragraph()
    p.add_run('I, ').font.size = Pt(10)
    p.add_run('{{ ev.expert_name }}').bold = True
    p.add_run(' (Name), ')
    p.add_run('{{ ev.expert_designation }}')
    p.add_run(' residing/employed at ')
    p.add_run('{{ ev.collection_address }}')
    p.add_run(' do hereby solemnly affirm and sincerely state and submit as follows: -')
    
    p = doc.add_paragraph('The produced electronic record/output of the digital record taken from the following device/digital record source (tick Mark): -')
    p.runs[0].font.size = Pt(10)
    
    p = doc.add_paragraph()
    p.add_run('{{ ev.cb_computer }} Computer/Storage Media    {{ ev.cb_cddvd }} CD/DVD    {{ ev.cb_server }} Server')
    p = doc.add_paragraph()
    p.add_run('{{ ev.cb_dvr }} DVR    {{ ev.cb_cloud }} Cloud    {{ ev.cb_mobile }} Mobile    {{ ev.cb_flash }} Flash Drive    {{ ev.cb_other }} Other: {{ ev.annexure_e_other_specify }}')
    
    p = doc.add_paragraph()
    p.add_run('Make & Model: ').bold = True
    p.add_run('{{ ev.make_model }}\n')
    p.add_run('Serial Number: ').bold = True
    p.add_run('{{ ev.serial_number }}\n')
    p.add_run('Color: ').bold = True
    p.add_run('{{ ev.color }}\n')
    p.add_run('IMEI/UIN/UID/MAC/Cloud ID: ').bold = True
    p.add_run('{{ ev.unique_id }} (as applicable)\n')
    p.add_run('and any other relevant information, if any, about the device/digital record: ').bold = True
    p.add_run('{{ ev.other_specific_notes }} (specify).')
    
    p = doc.add_paragraph()
    p.add_run('I state that the Hash value(s) of the electronic/digital record(s) is ').font.size = Pt(10)
    p.add_run('{{ ev.hash_value }}').bold = True
    p.add_run(', obtained through the following algorithm: -')
    
    p = doc.add_paragraph()
    p.add_run('{{ ev.hash_sha1 }} SHA1    {{ ev.hash_sha256 }} SHA256    {{ ev.hash_md5 }} MD5    {{ ev.hash_other }} Other (Legally acceptable standard)')
    doc.add_paragraph('(Hash report to be enclosed with the certificate)')
    
    p = doc.add_paragraph()
    p.add_run('Date (DD/MM/YYYY): {{ ev.collection_date }}        Time (IST): {{ ev.collection_time }} hours        Place: {{ ev.cert_place }}\n')
    p.add_run('(Name, Designation and Signature)\nName: ').bold = True
    p.add_run('{{ ev.expert_name }} ({{ ev.expert_designation }})')
    
    # Page break condition
    doc.add_paragraph('{% if not loop.last %}')
    p_break = doc.add_paragraph()
    p_break.add_run().add_break(WD_BREAK.PAGE)
    doc.add_paragraph('{% endif %}')
    doc.add_paragraph('{% endfor %}')
    
    out_path = os.path.join(OUT_DIR, 'annexure_e.docx')
    doc.save(out_path)
    print('Generated:', out_path)

# -------------------------------------------------------------
# 5. Annexure F: Cloud / Email evidence (per premise)
# -------------------------------------------------------------
def build_annexure_f():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11.0)
        section.page_height = Inches(8.5)
        
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.add_run('Annexure F - Cloud Software / Email Data Collection Form')
    r.bold = True
    r.font.size = Pt(13)
    
    p_info = doc.add_paragraph()
    p_info.add_run('NAME OF THE CASE: ').bold = True
    p_info.add_run('{{ case_name }}\n')
    p_info.add_run('PREMISES NAME & ADDRESS: ').bold = True
    p_info.add_run('{{ premise_name }}, {{ address }}\n')
    p_info.add_run('TEAM: ').bold = True
    p_info.add_run('{{ team }}')
    
    t = doc.add_table(rows=1, cols=5)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    headers = ['SN', 'DEVICE IMAGED', 'SIZE OF DATA', 'EXPORT FORMAT', 'SOURCE']
    for c_idx, h in enumerate(headers):
        cell = t.rows[0].cells[c_idx]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    r1 = t.add_row()
    r1.cells[0].text = '{%tr for row in cloud_rows %}'
    
    r2 = t.add_row()
    r2.cells[0].text = '{{ row.sn }}'
    r2.cells[1].text = '{{ row.device_imaged }}'
    r2.cells[2].text = '{{ row.size_of_data }}'
    r2.cells[3].text = '{{ row.export_format }}'
    r2.cells[4].text = '{{ row.source }}'
    
    r3 = t.add_row()
    r3.cells[0].text = '{%tr endfor %}'
    
    doc.add_paragraph()
    p_foot = doc.add_paragraph()
    p_foot.add_run('AUTHORISED OFFICER: {{ authorised_officer }}                            FORENSIC EXAMINER: {{ examiner_name }}').bold = True
    
    out_path = os.path.join(OUT_DIR, 'annexure_f.docx')
    doc.save(out_path)
    print('Generated:', out_path)

# -------------------------------------------------------------
# 6. Annexure G: Physical Inventory of Devices
# -------------------------------------------------------------
def build_annexure_g():
    src = os.path.join(SRC_DIR, '7. Annexure G- Physical inventory of devices.docx')
    doc = docx.Document(src)
    t = doc.tables[0]
    
    t.rows[1].cells[1].text = 'NAME OF ASSESSEE: {{ assessee_name }}'
    t.rows[2].cells[1].text = 'ADDRESS: {{ address }}'
    t.rows[3].cells[0].text = 'DEVICES FOUND AT PREMISES:'
    
    for r_idx in range(len(t.rows) - 1, 4, -1):
        tr = t.rows[r_idx]._tr
        tr.getparent().remove(tr)
        
    r1 = t.add_row()
    r1.cells[0].text = '{%tr for item in inventory_items %}'
    
    r2 = t.add_row()
    r2.cells[0].text = '{{ item.sn }}'
    r2.cells[1].text = '{{ item.nature }}'
    r2.cells[2].text = '{{ item.details }}'
    r2.cells[3].text = '{{ item.possession_of }}'
    r2.cells[4].text = '{{ item.owner_relation_text }}'
    r2.cells[5].text = '{{ item.remarks }}'
    
    r3 = t.add_row()
    r3.cells[0].text = '{%tr endfor %}'
    
    out_path = os.path.join(OUT_DIR, 'annexure_g.docx')
    doc.save(out_path)
    print('Generated:', out_path)

if __name__ == '__main__':
    try:
        build_annexure_b()
        build_annexure_c()
        build_annexure_d()
        build_annexure_e()
        build_annexure_f()
        build_annexure_g()
        print('All 6 Delhi/Gujarat templates generated successfully!')
    except Exception as e:
        print(f"Skipping base Delhi templates rebuild: {e}")
        
    import build_mumbai_templates
    build_mumbai_templates.build_mumbai_annexure_1()
    build_mumbai_templates.build_mumbai_annexure_3()
    build_mumbai_templates.build_mumbai_non_mobile()
    build_mumbai_templates.build_mumbai_mobile()
    build_mumbai_templates.build_mumbai_per_device_bsa()
    print('All 5 Mumbai templates generated successfully!')
