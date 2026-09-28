import os
import zipfile
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docxtpl import DocxTemplate

import excel_db
from config import TEMPLATES_DOCX_DIR, GENERATED_DOCS_DIR

DOCS_DIR = os.path.join(os.path.dirname(__file__), 'Documents')
os.makedirs(GENERATED_DOCS_DIR, exist_ok=True)

def _get_sym(condition):
    return '☑' if condition else '☐'

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

def format_annexure_1_model_details(ev):
    cat = ev.get('evidence_category', '').strip()
    lines = []
    
    # 1. Add Evidence Name in Model Details
    dev_name = ev.get('device_name', '').strip()
    if dev_name:
        lines.append(f"Evidence Name: {dev_name}")

    # Cloud / Email Asset
    is_cloud_asset = (cat == 'Cloud' or (cat == 'Other' and (ev.get('email_id') or ev.get('cloud_id'))))
    
    if is_cloud_asset:
        # STRICT RULE: Cloud uses ONLY email id, nothing else! No tenant, no IP, no domain!
        if ev.get('email_id'):
            lines.append(f"Email ID: {ev.get('email_id')}")
    else:
        # Physical Device (Desktop, Laptop, Mobile, Server, Pendrive, HDD, etc.)
        make = ev.get('make', '').strip()
        model = ev.get('model', '').strip()
        if make or model:
            lines.append(f"Make/Model: {f'{make} {model}'.strip()}")
        sn = ev.get('serial_number', '').strip()
        if sn:
            lines.append(f"S/N: {sn}")
        cap = ev.get('capacity', '').strip()
        if cap:
            lines.append(f"Capacity: {cap}")
        color = ev.get('color', '').strip()
        if color and color.lower() not in ['n/a', 'na', 'none', '']:
            lines.append(f"Color: {color}")
        os_val = ev.get('os', '').strip()
        if os_val:
            lines.append(f"OS: {os_val}")
        # Only show IMEI if Mobile!
        if cat == 'Mobile':
            imei1 = ev.get('imei1', '').strip()
            imei2 = ev.get('imei2', '').strip()
            if imei1:
                lines.append(f"IMEI 1: {imei1}")
            if imei2:
                lines.append(f"IMEI 2: {imei2}")
            
    return '\n'.join(lines) if lines else (ev.get('serial_number') or 'N/A')

def _build_evidence_context(evidence, premise=None, search=None):
    if not premise:
        premise = excel_db.get_premise(evidence.get('premise_id', '')) or {}
    if not search:
        search = excel_db.get_search(premise.get('search_id', '')) or {}
        
    cat = evidence.get('evidence_category', '')
    auth_officer = (
        evidence.get('authorised_officer') or 
        search.get('authorised_officer') or 
        ''
    )
    col_date = evidence.get('collection_date') or premise.get('date') or search.get('date') or ''
    col_time = evidence.get('collection_time') or premise.get('time') or ''
    col_addr = evidence.get('collection_address') or premise.get('address') or ''
    
    handling = [h.strip() for h in evidence.get('device_handling', '').split(',') if h.strip()]
    control = [c.strip() for c in evidence.get('device_control', '').split(',') if c.strip()]
    
    make = evidence.get('make', '')
    model = evidence.get('model', '')
    sn = evidence.get('serial_number', '')
    imei1 = evidence.get('imei1', '')
    imei2 = evidence.get('imei2', '')
    if cat == 'Mobile' and imei1:
        unique_id = f'IMEI1: {imei1}, IMEI2: {imei2}'
    elif cat == 'Cloud' or evidence.get('email_id'):
        unique_id = evidence.get('email_id', '') or sn
    else:
        unique_id = sn
    
    hash_algo = evidence.get('hash_algorithm', 'SHA256')
    md5_val = evidence.get('md5_hash', '').strip()
    sha1_val = evidence.get('sha1_hash', '').strip()
    sha256_list = [
        evidence.get(f'sha256_hash_{i}', '').strip()
        for i in range(1, 6)
        if evidence.get(f'sha256_hash_{i}', '').strip()
    ]
    
    hash_parts = []
    if md5_val:
        hash_parts.append(f"MD5: {md5_val}")
    if sha1_val:
        hash_parts.append(f"SHA1: {sha1_val}")
    if sha256_list:
        hash_parts.append(f"SHA256: {', '.join(sha256_list)}")
        
    hash_val = ', '.join(hash_parts) if hash_parts else 'N/A'
    
    has_md5 = bool(md5_val) or (hash_algo == 'MD5')
    has_sha1 = bool(sha1_val) or (hash_algo == 'SHA1')
    has_sha256 = bool(sha256_list) or (hash_algo == 'SHA256')
    has_other = (hash_algo == 'Other') or (not has_md5 and not has_sha1 and not has_sha256)
    
    sub_evs = excel_db.get_sub_evidences(evidence.get('evidence_id', ''))
    custody_rows = excel_db.get_custody_rows(evidence.get('evidence_id', ''))
    
    # Build comprehensive model_details string for Mumbai Annexure 1
    details_parts = []
    if make or model:
        details_parts.append(f"Make: {make}  Model: {model}".strip())
    if sn:
        details_parts.append(f"S/N: {sn}")
    if evidence.get('capacity'):
        details_parts.append(f"Capacity: {evidence.get('capacity')}")
    if evidence.get('color'):
        details_parts.append(f"Color: {evidence.get('color')}")
    if evidence.get('os'):
        details_parts.append(f"OS: {evidence.get('os')}")
    if cat == 'Mobile':
        if imei1:
            details_parts.append(f"IMEI1: {imei1}")
        if imei2:
            details_parts.append(f"IMEI2: {imei2}")
    if evidence.get('email_id'):
        details_parts.append(f"Email ID: {evidence.get('email_id')}")
    if evidence.get('cloud_id'):
        details_parts.append(f"Cloud ID: {evidence.get('cloud_id')}")
    if evidence.get('ip_address'):
        details_parts.append(f"IP: {evidence.get('ip_address')}")
    if evidence.get('domain_name'):
        details_parts.append(f"Domain: {evidence.get('domain_name')}")
    model_details_str = '\n'.join(details_parts) if details_parts else (sn or "N/A")

    # Category normalization & flags
    is_cloud = (cat == 'Cloud' or (cat == 'Other' and (evidence.get('cloud_id') or evidence.get('email_id') or 'cloud' in evidence.get('device_name', '').lower() or 'mail' in evidence.get('device_name', '').lower())))
    is_file_folder = (cat in ['FileFolder', 'File/Folder'])

    # Non-mobile collection form: "IF OTHERS SPECIFY"
    if is_cloud:
        non_mobile_other_specify = evidence.get('email_id') or 'Cloud Email Account'
    elif cat in ['HDD', 'SSD', 'Pendrive']:
        non_mobile_other_specify = cat
    elif cat == 'Other':
        non_mobile_other_specify = evidence.get('device_name') or 'Other'
    else:
        non_mobile_other_specify = 'N/A'

    # Annexure E / BSA Certificate: "Other: (specify)"
    if is_file_folder:
        annexure_e_other_specify = 'File/Folder'
    elif cat == 'Other' and not is_cloud:
        annexure_e_other_specify = evidence.get('device_name') or 'Other'
    else:
        annexure_e_other_specify = ''

    # Officer sealing notes removed per user requirement
    notes_by_officer = ''
    other_specific_notes = evidence.get('other_specific_notes') or evidence.get('remarks') or 'None'

    master_make = evidence.get('master_copy_make', '')
    master_model = evidence.get('master_copy_model', '')
    master_make_model = f"{master_make} {master_model}".strip()

    working_make = evidence.get('working_copy_make', '')
    working_model = evidence.get('working_copy_model', '')
    working_make_model = f"{working_make} {working_model}".strip()

    is_owner_val = evidence.get('is_owner') or 'Yes'

    ctx = {
        'evidence_id': evidence.get('evidence_id', ''),
        'case_name': search.get('case_name', ''),
        'premise_name': premise.get('premise_name', ''),
        'search_warrant_details': search.get('search_warrant_details') or f"SW/{search.get('case_name', '')}",
        'seizure_date': col_date,
        'place': evidence.get('cert_place') or (premise.get('premise_name', '') + ', ' + col_addr),
        'team': premise.get('team', '') or 'Team 1',
        'annexure_no': 'Annexure 3',
        'authorised_officer': auth_officer,
        'assessee_name': evidence.get('assessee_name', ''),
        'device_name': evidence.get('device_name', ''),
        'collection_date': col_date,
        'collection_time': col_time,
        'collection_address': col_addr,
        'examiner_name': evidence.get('examiner_name', ''),
        'examiner_details': evidence.get('examiner_details', ''),
        
        # Non-Mobile Collection Form Checkboxes
        'is_laptop_sym': _get_sym(cat == 'Laptop'),
        'is_desktop_sym': _get_sym(cat == 'Desktop'),
        'is_server_sym': _get_sym(cat == 'Server'),
        'is_filefolder_sym': _get_sym(is_file_folder),
        'is_other_sym': _get_sym(is_cloud or (cat in ['Other', 'HDD', 'SSD', 'Pendrive'])),
        'non_mobile_other_specify': non_mobile_other_specify,
        'other_specify': non_mobile_other_specify,

        # Annexure 1 specific category checkboxes
        'cb_desktop': _get_sym(cat == 'Desktop'),
        'cb_laptop': _get_sym(cat == 'Laptop'),
        'cb_mobile': _get_sym(cat == 'Mobile'),
        'cb_pendrive': _get_sym(cat in ['Pendrive', 'Flash Drive']),
        'cb_harddisk': _get_sym(cat in ['HDD', 'SSD', 'Harddisk']),
        'cb_server': _get_sym(cat == 'Server'),

        # Annexure 1 table fields
        'model_details': format_annexure_1_model_details(evidence),
        'possession_of': evidence.get('owner_user', '') or 'N/A',
        'is_owner': is_owner_val,
        'relation': evidence.get('owner_relation', '') or 'Representative',
        'remarks': '',
        
        # Device status
        'hib_on_sym': _get_sym(evidence.get('hibernation_sleep') == 'On'),
        'hib_off_sym': _get_sym(evidence.get('hibernation_sleep') == 'Off'),
        'visible_on_screen': evidence.get('visible_on_screen', ''),
        
        # Hardware specs
        'make': make,
        'model': model,
        'make_model': f'{make} {model}'.strip(),
        'serial_number': sn,
        'capacity': evidence.get('capacity', ''),
        'color': evidence.get('color', ''),
        'os': evidence.get('os', ''),
        'owner_user': evidence.get('owner_user', ''),
        'owner_relation': evidence.get('owner_relation', ''),
        'unique_id': unique_id,
        'email_id': evidence.get('email_id', ''),
        'cloud_id': '' if is_cloud else evidence.get('cloud_id', ''),
        'ip_address': '' if is_cloud else evidence.get('ip_address', ''),
        'domain_name': '' if is_cloud else evidence.get('domain_name', ''),
        
        # Shutdown
        'shut_normal_sym': _get_sym(evidence.get('shutdown_type') == 'Normal'),
        'shut_plug_sym': _get_sym(evidence.get('shutdown_type') == 'PowerPlugPulled'),
        'shut_battery_sym': _get_sym(evidence.get('shutdown_type') == 'BatteryRemoved'),
        
        # Encryption
        'enc_yes_sym': _get_sym(evidence.get('encryption_present') == 'Y'),
        'enc_no_sym': _get_sym(evidence.get('encryption_present') == 'N'),
        'encryption_software': evidence.get('encryption_software', 'None'),
        
        # Handling
        'handle_seizure_sym': _get_sym('Seizure' in handling or not handling),
        'handle_imaging_sym': _get_sym('Imaging' in handling),
        'handle_backup_sym': _get_sym('Backup' in handling),
        'handle_reporting_sym': _get_sym('Reporting' in handling),
        'imaging_tools_details': evidence.get('imaging_tools_details', ''),
        
        # Hashes
        'md5_hash': evidence.get('md5_hash', ''),
        'sha1_hash': evidence.get('sha1_hash', ''),
        'sha256_hash_1': evidence.get('sha256_hash_1', ''),
        'sha256_hash_2': evidence.get('sha256_hash_2', ''),
        'sha256_hash_3': evidence.get('sha256_hash_3', ''),
        'sha256_hash_4': evidence.get('sha256_hash_4', ''),
        'sha256_hash_5': evidence.get('sha256_hash_5', ''),
        'hash_value': hash_val,
        'hash_values_seq': hash_val,
        'hash_sha1': _get_sym(has_sha1),
        'hash_sha256': _get_sym(has_sha256),
        'hash_md5': _get_sym(has_md5),
        'hash_other': _get_sym(has_other),
        
        # Mobile specific
        'imei1': imei1,
        'imei2': imei2,
        'imsi': evidence.get('imsi', ''),
        'sim_provider': evidence.get('sim_provider', ''),
        'sim_yes_sym': _get_sym(evidence.get('sim_card_present') == 'Y'),
        'sim_no_sym': _get_sym(evidence.get('sim_card_present') == 'N'),
        'media_card_make': evidence.get('media_card_make', ''),
        'media_card_model': evidence.get('media_card_model', ''),
        'media_card_sn': evidence.get('media_card_sn', ''),
        'media_yes_sym': _get_sym(evidence.get('media_card_present') == 'Y'),
        'media_no_sym': _get_sym(evidence.get('media_card_present') == 'N'),
        'state_on_sym': _get_sym(evidence.get('mobile_state') == 'On'),
        'state_off_sym': _get_sym(evidence.get('mobile_state') == 'Off'),
        'state_offline_sym': _get_sym(evidence.get('mobile_state') == 'Offline'),
        
        # Master and Working copies
        'master_copy_make': master_make,
        'master_copy_model': master_model,
        'master_copy_make_model': master_make_model,
        'master_copy_sn': evidence.get('master_copy_sn', ''),
        'master_copy_capacity': evidence.get('master_copy_capacity', ''),
        'working_copy_make': working_make,
        'working_copy_model': working_model,
        'working_copy_make_model': working_make_model,
        'working_copy_sn': evidence.get('working_copy_sn', ''),
        'working_copy_capacity': evidence.get('working_copy_capacity', ''),
        
        # Returned & Witness
        'returned_yes_sym': _get_sym(evidence.get('device_returned') == 'Y'),
        'returned_no_sym': _get_sym(evidence.get('device_returned') == 'N'),
        'returned_date': evidence.get('returned_date', ''),
        'returned_time': evidence.get('returned_time', ''),
        'witness_yes_sym': _get_sym(evidence.get('witness_signature_taken') == 'Y'),
        'witness_no_sym': _get_sym(evidence.get('witness_signature_taken') == 'N'),
        'notes_by_officer': notes_by_officer,
        'other_specific_notes': other_specific_notes,
        
        # Annexure E / BSA Certificate specific
        'party_name': evidence.get('party_name') or evidence.get('assessee_name', ''),
        'party_relation': evidence.get('party_relation') or 'Authorized Representative',
        'party_residing_at': evidence.get('party_residing_at') or col_addr,
        'expert_name': evidence.get('expert_name') or evidence.get('examiner_name', '') or 'Forensic Expert',
        'expert_father_name': evidence.get('expert_father_name') or 'N/A',
        'expert_designation': evidence.get('expert_designation') or evidence.get('examiner_details', 'Senior Digital Forensic Examiner'),
        'cert_place': evidence.get('cert_place') or (premise.get('premise_name', '') + ', ' + col_addr),
        
        'cb_computer': _get_sym(cat in ['Laptop', 'Desktop', 'HDD', 'SSD', 'Pendrive']),
        'cb_cddvd': _get_sym(cat == 'CD/DVD'),
        'cb_server': _get_sym(cat == 'Server'),
        'cb_dvr': _get_sym(cat == 'DVR'),
        'cb_cloud': _get_sym(is_cloud),
        'cb_mobile': _get_sym(cat == 'Mobile'),
        'cb_flash': _get_sym(cat in ['Flash Drive', 'Pendrive']),
        'cb_other': _get_sym(is_file_folder or (cat == 'Other' and not is_cloud)),
        'annexure_e_other_specify': annexure_e_other_specify,
        
        'ctrl_owned': _get_sym('Owned' in control or not control),
        'ctrl_maintained': _get_sym('Maintained' in control),
        'ctrl_managed': _get_sym('Managed' in control),
        'ctrl_operated': _get_sym('Operated' in control),
        
        # Relational collections
        'sub_evidences': sub_evs,
        'custody_rows': custody_rows
    }
    return ctx

def generate_evidence_docs(evidence_id):
    evidence = excel_db.get_evidence(evidence_id)
    if not evidence:
        raise ValueError(f'Evidence {evidence_id} not found')
        
    premise = excel_db.get_premise(evidence['premise_id']) or {}
    search = excel_db.get_search(premise.get('search_id', '')) or {}
    jurisdiction = search.get('jurisdiction', 'Delhi')
    cat = evidence.get('evidence_category', '')
    
    generated_files = []
    
    # 1. Gujarat routing
    if jurisdiction == 'Gujarat':
        ctx = _build_evidence_context(evidence, premise, search)
        tpl_path = os.path.join(TEMPLATES_DOCX_DIR, 'annexure_e.docx')
        tpl = DocxTemplate(tpl_path)
        tpl.render({'evidences': [ctx]})
        out_name = f'Annexure_E_{evidence_id}.docx'
        out_path = os.path.join(GENERATED_DOCS_DIR, out_name)
        tpl.save(out_path)
        generated_files.append(out_path)
        return {'files': generated_files, 'zip_path': None}
        
    # 2. Mumbai routing
    if jurisdiction == 'Mumbai':
        ctx = _build_evidence_context(evidence, premise, search)
        
        # 1. Device Collection Form (Mobile vs Non-Mobile)
        if cat == 'Mobile':
            mob_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'mumbai_device_mobile.docx'))
            mob_tpl.render(ctx)
            mob_path = os.path.join(GENERATED_DOCS_DIR, f'Device_Collection_Mobile_{evidence_id}.docx')
            mob_tpl.save(mob_path)
            generated_files.append(mob_path)
        else:
            nmob_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'mumbai_device_non_mobile.docx'))
            nmob_tpl.render(ctx)
            nmob_path = os.path.join(GENERATED_DOCS_DIR, f'Device_Collection_NonMobile_{evidence_id}.docx')
            nmob_tpl.save(nmob_path)
            generated_files.append(nmob_path)
            
        # 3. BSA 2023 Sec 63(4)(c) Certificate (Per device)
        bsa_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'mumbai_per_device_bsa.docx'))
        bsa_tpl.render(ctx)
        bsa_path = os.path.join(GENERATED_DOCS_DIR, f'BSA_Certificate_{evidence_id}.docx')
        bsa_tpl.save(bsa_path)
        generated_files.append(bsa_path)
        
        zip_name = f'Evidence_{evidence_id}_Mumbai_Docs.zip'
        zip_path = os.path.join(GENERATED_DOCS_DIR, zip_name)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in generated_files:
                zf.write(f, os.path.basename(f))
                
        return {'files': generated_files, 'zip_path': zip_path}

    # 3. Delhi routing
    ctx = _build_evidence_context(evidence, premise, search)
    
    if cat == 'Mobile':
        c_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'annexure_c.docx'))
        c_tpl.render(ctx)
        c_path = os.path.join(GENERATED_DOCS_DIR, f'Annexure_C_{evidence_id}.docx')
        c_tpl.save(c_path)
        generated_files.append(c_path)
        
        d_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'annexure_d.docx'))
        d_tpl.render(ctx)
        d_path = os.path.join(GENERATED_DOCS_DIR, f'Annexure_D_{evidence_id}.docx')
        d_tpl.save(d_path)
        generated_files.append(d_path)
        
        e_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'annexure_e.docx'))
        e_tpl.render({'evidences': [ctx]})
        e_path = os.path.join(GENERATED_DOCS_DIR, f'Annexure_E_{evidence_id}.docx')
        e_tpl.save(e_path)
        generated_files.append(e_path)
        
    else:
        b_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'annexure_b.docx'))
        b_tpl.render(ctx)
        b_path = os.path.join(GENERATED_DOCS_DIR, f'Annexure_B_{evidence_id}.docx')
        b_tpl.save(b_path)
        generated_files.append(b_path)
        
        c_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'annexure_c.docx'))
        c_tpl.render(ctx)
        c_path = os.path.join(GENERATED_DOCS_DIR, f'Annexure_C_{evidence_id}.docx')
        c_tpl.save(c_path)
        generated_files.append(c_path)
        
        e_tpl = DocxTemplate(os.path.join(TEMPLATES_DOCX_DIR, 'annexure_e.docx'))
        e_tpl.render({'evidences': [ctx]})
        e_path = os.path.join(GENERATED_DOCS_DIR, f'Annexure_E_{evidence_id}.docx')
        e_tpl.save(e_path)
        generated_files.append(e_path)
        
    zip_name = f'Evidence_{evidence_id}_Annexures.zip'
    zip_path = os.path.join(GENERATED_DOCS_DIR, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in generated_files:
            zf.write(f, os.path.basename(f))
            
    return {'files': generated_files, 'zip_path': zip_path}

def generate_search_annexure_e(search_id):
    search = excel_db.get_search(search_id)
    if not search:
        raise ValueError(f'Search {search_id} not found')
        
    premises = excel_db.get_premises(search_id)
    all_evidence_contexts = []
    
    for prem in premises:
        evs = excel_db.get_evidences(prem['premise_id'])
        for ev in evs:
            ctx = _build_evidence_context(ev, prem, search)
            all_evidence_contexts.append(ctx)
            
    if not all_evidence_contexts:
        raise ValueError(f'No evidence found under search {search_id}')
        
    tpl_path = os.path.join(TEMPLATES_DOCX_DIR, 'annexure_e.docx')
    tpl = DocxTemplate(tpl_path)
    tpl.render({'evidences': all_evidence_contexts})
    
    out_name = f'Annexure_E_Search_{search_id}.docx'
    out_path = os.path.join(GENERATED_DOCS_DIR, out_name)
    tpl.save(out_path)
    return out_path

def generate_premise_annexure_g(premise_id):
    premise = excel_db.get_premise(premise_id)
    if not premise:
        raise ValueError(f'Premise {premise_id} not found')
        
    search = excel_db.get_search(premise.get('search_id', '')) or {}
    evidences = excel_db.get_evidences(premise_id)
    
    inventory_items = []
    sn_counter = 1
    
    for ev in evidences:
        cat = ev.get('evidence_category', '')
        make = ev.get('make', '')
        model = ev.get('model', '')
        sn = ev.get('serial_number', '')
        cap = ev.get('capacity', '')
        imei = f" IMEI1: {ev.get('imei1', '')}" if cat == 'Mobile' and ev.get('imei1') else ''
        details = f"{make} {model}, S/N: {sn}, Cap: {cap}{imei}".strip(' ,')
        owner = ev.get('owner_user', 'N/A')
        rel = ev.get('owner_relation', '')
        owner_rel_text = f"{owner} ({rel})" if rel else owner
        
        inventory_items.append({
            'sn': str(sn_counter),
            'nature': f"{cat} ({ev.get('device_name', '')})",
            'details': details,
            'possession_of': owner,
            'owner_relation_text': owner_rel_text,
            'remarks': ev.get('notes_by_officer', '') or 'Recovered from premise'
        })
        sn_counter += 1
        
        sub_evs = excel_db.get_sub_evidences(ev.get('evidence_id', ''))
        for se in sub_evs:
            se_details = f"Internal {se.get('sub_type', 'Storage')}: {se.get('make', '')} {se.get('model', '')}, S/N: {se.get('sn', '')}, Cap: {se.get('capacity', '')}".strip(' ,')
            inventory_items.append({
                'sn': str(sn_counter),
                'nature': f"Sub-Item ({se.get('sub_type', '')})",
                'details': se_details,
                'possession_of': f"Inside {ev.get('device_name', '')} ({ev.get('evidence_id', '')})",
                'owner_relation_text': owner_rel_text,
                'remarks': 'Internal component'
            })
            sn_counter += 1
            
    assessee = evidences[0].get('assessee_name', '') if evidences else premise.get('premise_name', '')
    
    tpl_path = os.path.join(TEMPLATES_DOCX_DIR, 'annexure_g.docx')
    tpl = DocxTemplate(tpl_path)
    tpl.render({
        'assessee_name': assessee,
        'address': premise.get('address', ''),
        'inventory_items': inventory_items
    })
    
    out_name = f'Annexure_G_Premise_{premise_id}.docx'
    out_path = os.path.join(GENERATED_DOCS_DIR, out_name)
    tpl.save(out_path)
    return out_path

def generate_premise_annexure_f(premise_id):
    premise = excel_db.get_premise(premise_id)
    if not premise:
        raise ValueError(f'Premise {premise_id} not found')
        
    search = excel_db.get_search(premise.get('search_id', '')) or {}
    cloud_evs = excel_db.get_cloud_evidences(premise_id)
    
    auth_officer = search.get('authorised_officer', '')
    examiner = ''
    evs = excel_db.get_evidences(premise_id)
    if evs and evs[0].get('examiner_name'):
        examiner = evs[0]['examiner_name']
        
    rows = []
    for idx, ce in enumerate(cloud_evs, start=1):
        rows.append({
            'sn': ce.get('sn') or str(idx),
            'device_imaged': ce.get('device_imaged', ''),
            'size_of_data': ce.get('size_of_data', ''),
            'export_format': ce.get('export_format', ''),
            'source': ce.get('source', '')
        })
        
    tpl_path = os.path.join(TEMPLATES_DOCX_DIR, 'annexure_f.docx')
    tpl = DocxTemplate(tpl_path)
    tpl.render({
        'case_name': search.get('case_name', ''),
        'premise_name': premise.get('premise_name', ''),
        'address': premise.get('address', ''),
        'team': premise.get('team', ''),
        'authorised_officer': auth_officer,
        'examiner_name': examiner,
        'cloud_rows': rows
    })
    
    out_name = f'Annexure_F_Premise_{premise_id}.docx'
    out_path = os.path.join(GENERATED_DOCS_DIR, out_name)
    tpl.save(out_path)
    return out_path

def generate_mumbai_annexure_1(premise_id):
    premise = excel_db.get_premise(premise_id)
    if not premise:
        raise ValueError(f'Premise {premise_id} not found')
    search = excel_db.get_search(premise.get('search_id', '')) or {}
    evidences = excel_db.get_evidences(premise_id)
    if not evidences:
        raise ValueError(f'No evidence devices found for premise {premise_id}')
        
    src = os.path.join(DOCS_DIR, 'ANNEXURE-1 Identification and Physical Inventory of Relevant Devices.docx')
    doc = docx.Document(src)
    
    # 0.6 inch margins for spacious, crisp table layout
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)
        
    team = premise.get('team', '') or 'Team 1'
    auth_officer = search.get('authorised_officer') or (evidences[0].get('authorised_officer') if evidences else '')
    col_date = premise.get('date') or search.get('date') or ''
    
    for p in doc.paragraphs:
        if 'TEAM CODE:' in p.text:
            p.text = f'TEAM CODE: {team}'
            if p.runs:
                p.runs[0].font.bold = True
        elif 'SIGN OF THE AUTHORISED OFFICER:' in p.text:
            p.text = f'SIGN OF THE AUTHORISED OFFICER: __________________________\n({auth_officer})'
        elif 'DATE:' in p.text:
            p.text = f'DATE: {col_date}'
            
    t = doc.tables[0]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    
    # Remove existing empty rows 1-6
    while len(t.rows) > 1:
        tr = t.rows[-1]._tr
        tr.getparent().remove(tr)
        
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
                
    # Standard categories order
    standard_categories = ['Desktop', 'Laptop', 'Mobile', 'Pendrive', 'Harddisk', 'Server']
    extra_categories = []
    for ev in evidences:
        cat = ev.get('evidence_category', '').strip()
        if cat in ['HDD', 'SSD']:
            c_norm = 'Harddisk'
        elif cat in ['Flash Drive']:
            c_norm = 'Pendrive'
        else:
            c_norm = cat
        if c_norm not in standard_categories and c_norm not in extra_categories:
            extra_categories.append(c_norm)
            
    all_categories = standard_categories + extra_categories
    
    for cat_name in all_categories:
        matching = []
        for ev in evidences:
            c = ev.get('evidence_category', '').strip()
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
                
                # Column 1: Nature of Evidence / Where Found
                loc = ev.get('location_found', '').strip()
                if not loc:
                    if cat_name == 'Cloud':
                        loc = 'Cloud Storage'
                    else:
                        loc = ev.get('collection_address', '').strip() or 'Recovered from premise'
                r.cells[1].text = loc
                
                # Column 2: Model Details (Includes Evidence Name)
                r.cells[2].text = format_annexure_1_model_details(ev)
                
                # Column 3: Found in possession
                r.cells[3].text = ev.get('owner_user', '') or 'N/A'
                
                # Column 4: Is Owner: If yes Then Name, else No
                is_owner_val = ev.get('is_owner', 'Yes')
                owner_name = ev.get('owner_user', '').strip() or ev.get('party_name', '').strip() or 'Yes'
                if is_owner_val == 'Yes' or is_owner_val is True:
                    r.cells[4].text = owner_name
                else:
                    r.cells[4].text = 'No'
                    
                # Column 5: Relation to target
                r.cells[5].text = ev.get('owner_relation', '') or 'N/A'
                
                # Column 6: Remarks (blank)
                r.cells[6].text = ''
                
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
            if cat_name in standard_categories:
                r = t.add_row()
                r.cells[0].text = cat_name
                for c_i in range(1, 7):
                    r.cells[c_i].text = '-'
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
                                
    out_name = f'Annexure_1_Premise_{premise_id}.docx'
    out_path = os.path.join(GENERATED_DOCS_DIR, out_name)
    doc.save(out_path)
    return out_path

def generate_mumbai_annexure_3_coc(premise_id):
    premise = excel_db.get_premise(premise_id)
    if not premise:
        raise ValueError(f'Premise {premise_id} not found')
    search = excel_db.get_search(premise.get('search_id', '')) or {}
    evidences = excel_db.get_evidences(premise_id)
    if not evidences:
        raise ValueError(f'No evidence devices found for premise {premise_id}')
        
    all_device_contexts = []
    for ev in evidences:
        ctx = _build_evidence_context(ev, premise, search)
        all_device_contexts.append(ctx)
        
    team = premise.get('team', '') or 'Team 1'
    annexure_no = 'Annexure 3'
    
    # Unique Master and Working copy list per user requirement
    unique_pairs = []
    seen_pairs = set()
    for ev in all_device_contexts:
        m_sn = ev.get('master_copy_sn', '').strip()
        w_sn = ev.get('working_copy_sn', '').strip()
        if not m_sn and not w_sn and len(all_device_contexts) > 1:
            continue
        pair_key = (m_sn, w_sn)
        if pair_key not in seen_pairs:
            seen_pairs.add(pair_key)
            unique_pairs.append((m_sn, w_sn))

    if not unique_pairs:
        unique_pairs = [('N/A', 'N/A')]

    transfer_lines = [f"1. TEAM: {team}    ANNX: {annexure_no}\n"]
    if len(unique_pairs) == 1:
        m_sn, w_sn = unique_pairs[0]
        transfer_lines.append("Evidence:")
        transfer_lines.append(f"Master copy S/N: {m_sn or 'N/A'}")
        transfer_lines.append(f"Working copy S/N: {w_sn or 'N/A'}")
    else:
        for idx, (m_sn, w_sn) in enumerate(unique_pairs, 1):
            transfer_lines.append(f"Evidence {idx}:")
            transfer_lines.append(f"Master copy S/N: {m_sn or 'N/A'}")
            transfer_lines.append(f"Working copy S/N: {w_sn or 'N/A'}\n")
    transfer_text = "\n".join(transfer_lines).strip()

    transfer_1_text = transfer_text
    transfer_2_text = transfer_text
    
    auth_officer = (
        search.get('authorised_officer') or 
        (all_device_contexts[0].get('authorised_officer') if all_device_contexts else '')
    )
    expert_name = (
        all_device_contexts[0].get('expert_name') or 
        all_device_contexts[0].get('examiner_name') or 
        'Forensic Expert'
    )
    col_date = premise.get('date') or search.get('date') or ''
    col_time = premise.get('time') or ''
    
    coc_context = {
        'search_warrant_details': search.get('search_warrant_details') or f"SW/{search.get('case_name', '')}",
        'seizure_date': col_date,
        'authorised_officer': auth_officer,
        'place': premise.get('address') or premise.get('premise_name', ''),
        'case_name': search.get('case_name', ''),
        'address': premise.get('address', ''),
        'team': team,
        'annexure_no': annexure_no,
        'collection_date': col_date,
        'collection_time': col_time,
        'expert_name': expert_name,
        'evidences': all_device_contexts,
        'transfer_1_devices': transfer_1_text,
        'transfer_2_devices': transfer_2_text
    }
    
    tpl_path = os.path.join(TEMPLATES_DOCX_DIR, 'mumbai_annexure_3_coc.docx')
    tpl = DocxTemplate(tpl_path)
    tpl.render(coc_context)
    
    out_name = f'Annexure_3_COC_Premise_{premise_id}.docx'
    out_path = os.path.join(GENERATED_DOCS_DIR, out_name)
    tpl.save(out_path)
    return out_path
