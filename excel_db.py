import os
import threading
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import DATABASE_PATH, ID_PREFIXES, ID_LENGTHS

_db_lock = threading.RLock()

SHEET_SCHEMAS = {
    'Searches': {
        'pk': 'search_id',
        'columns': [
            'search_id', 'case_name', 'jurisdiction', 'date', 'authorised_officer', 'remarks',
            'search_warrant_details'
        ]
    },
    'Premises': {
        'pk': 'premise_id',
        'columns': [
            'premise_id', 'search_id', 'premise_name', 'address', 'team', 'date', 'time'
        ]
    },
    'Evidences': {
        'pk': 'evidence_id',
        'columns': [
            'evidence_id', 'premise_id', 'assessee_name', 'evidence_category', 'device_name',
            'make', 'model', 'serial_number', 'capacity', 'color', 'owner_user', 'owner_relation',
            'os', 'imei1', 'imei2', 'imsi', 'sim_card_present', 'sim_provider',
            'media_card_present', 'media_card_make', 'media_card_model', 'media_card_sn',
            'hibernation_sleep', 'visible_on_screen', 'shutdown_type', 'mobile_state',
            'encryption_present', 'encryption_software', 'device_handling', 'imaging_tools_details',
            'md5_hash', 'sha1_hash', 'sha256_hash_1', 'sha256_hash_2', 'sha256_hash_3',
            'sha256_hash_4', 'sha256_hash_5', 'master_copy_make', 'master_copy_model',
            'master_copy_sn', 'master_copy_capacity', 'working_copy_make', 'working_copy_model',
            'working_copy_sn', 'working_copy_capacity', 'device_returned', 'returned_date',
            'returned_time', 'witness_signature_taken', 'notes_by_officer', 'examiner_name',
            'examiner_details', 'collection_date', 'collection_time', 'collection_address',
            'authorised_officer', 'hash_algorithm', 'device_control', 'cert_place',
            'party_name', 'party_relation', 'party_residing_at', 'expert_name', 'expert_designation',
            'is_owner', 'email_id', 'cloud_id', 'ip_address', 'domain_name', 'other_specific_notes', 'expert_father_name',
            'location_found'
        ]
    },
    'SubEvidences': {
        'pk': 'sub_evidence_id',
        'columns': [
            'sub_evidence_id', 'evidence_id', 'sub_type', 'make', 'model', 'sn', 'capacity', 'owner_user'
        ]
    },
    'ChainOfCustody': {
        'pk': 'custody_id',
        'columns': [
            'custody_id', 'evidence_id', 'reason_action', 'given_by_name', 'given_by_signature',
            'received_by_name', 'received_by_signature', 'date', 'time'
        ]
    },
    'CloudEvidences': {
        'pk': 'cloud_evidence_id',
        'columns': [
            'cloud_evidence_id', 'premise_id', 'sn', 'device_imaged', 'size_of_data',
            'export_format', 'source'
        ]
    }
}

HEADER_FILL = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
HEADER_FONT = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
HEADER_ALIGNMENT = Alignment(horizontal='center', vertical='center', wrap_text=True)

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

def init_db(force=False):
    with _db_lock:
        os.makedirs(os.path.dirname(DATABASE_PATH), exist_ok=True)
        if os.path.exists(DATABASE_PATH) and not force:
            # Check for missing columns in existing database and add them
            wb = openpyxl.load_workbook(DATABASE_PATH)
            changed = False
            for sheet_name, schema in SHEET_SCHEMAS.items():
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    existing_cols = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
                    for col_name in schema['columns']:
                        if col_name not in existing_cols:
                            next_col = ws.max_column + 1
                            cell = ws.cell(row=1, column=next_col, value=col_name)
                            cell.fill = HEADER_FILL
                            cell.font = HEADER_FONT
                            cell.alignment = HEADER_ALIGNMENT
                            col_letter = get_column_letter(next_col)
                            ws.column_dimensions[col_letter].width = max(len(col_name) + 4, 15)
                            existing_cols.append(col_name)
                            changed = True
            if changed:
                wb.save(DATABASE_PATH)
            wb.close()
            return

        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # remove default sheet

        for sheet_name, schema in SHEET_SCHEMAS.items():
            ws = wb.create_sheet(title=sheet_name)
            ws.views.sheetView[0].showGridLines = True
            ws.freeze_panes = 'A2'

            for col_idx, col_name in enumerate(schema['columns'], start=1):
                cell = ws.cell(row=1, column=col_idx, value=col_name)
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = HEADER_ALIGNMENT
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = max(len(col_name) + 4, 15)

        wb.save(DATABASE_PATH)
        wb.close()

def _load_workbook():
    init_db()
    return openpyxl.load_workbook(DATABASE_PATH)

def generate_next_id(sheet_name):
    prefix = ID_PREFIXES.get(sheet_name, 'ID')
    length = ID_LENGTHS.get(sheet_name, 3)
    wb = _load_workbook()
    ws = wb[sheet_name]
    max_num = 0

    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val is not None:
            val_str = str(val).strip()
            if val_str.startswith(prefix):
                num_part = val_str[len(prefix):]
                if num_part.isdigit():
                    max_num = max(max_num, int(num_part))

    wb.close()
    return f'{prefix}{max_num + 1:0{length}d}'

def get_all(sheet_name):
    with _db_lock:
        wb = _load_workbook()
        if sheet_name not in wb.sheetnames:
            wb.close()
            return []
        ws = wb[sheet_name]
        schema = SHEET_SCHEMAS[sheet_name]
        columns = schema['columns']

        rows = []
        for r in range(2, ws.max_row + 1):
            pk_val = ws.cell(row=r, column=1).value
            if pk_val is None or str(pk_val).strip() == '':
                continue
            item = {}
            for c_idx, col_name in enumerate(columns, start=1):
                raw = ws.cell(row=r, column=c_idx).value
                item[col_name] = '' if raw is None else str(raw)
            rows.append(item)

        wb.close()
        return rows

def get_by_id(sheet_name, id_value):
    if not id_value:
        return None
    items = get_all(sheet_name)
    pk = SHEET_SCHEMAS[sheet_name]['pk']
    for item in items:
        if str(item.get(pk, '')).strip() == str(id_value).strip():
            return item
    return None

def filter_rows(sheet_name, **kwargs):
    items = get_all(sheet_name)
    results = []
    for item in items:
        match = True
        for k, v in kwargs.items():
            if str(item.get(k, '')).strip() != str(v).strip():
                match = False
                break
        if match:
            results.append(item)
    return results

def insert_row(sheet_name, data):
    schema = SHEET_SCHEMAS[sheet_name]
    pk = schema['pk']
    columns = schema['columns']

    with _db_lock:
        wb = _load_workbook()
        ws = wb[sheet_name]

        if not data.get(pk):
            prefix = ID_PREFIXES.get(sheet_name, 'ID')
            length = ID_LENGTHS.get(sheet_name, 3)
            max_num = 0
            for row in range(2, ws.max_row + 1):
                val = ws.cell(row=row, column=1).value
                if val:
                    val_str = str(val).strip()
                    if val_str.startswith(prefix):
                        num_part = val_str[len(prefix):]
                        if num_part.isdigit():
                            max_num = max(max_num, int(num_part))
            new_id = f'{prefix}{max_num + 1:0{length}d}'
            data[pk] = new_id

        next_row = ws.max_row + 1
        for c_idx, col_name in enumerate(columns, start=1):
            val = data.get(col_name, '')
            cell = ws.cell(row=next_row, column=c_idx, value=val)
            cell.border = THIN_BORDER

        wb.save(DATABASE_PATH)
        wb.close()
        return data[pk]

def update_row(sheet_name, id_value, data):
    schema = SHEET_SCHEMAS[sheet_name]
    pk = schema['pk']
    columns = schema['columns']

    with _db_lock:
        wb = _load_workbook()
        ws = wb[sheet_name]
        target_row = None

        for r in range(2, ws.max_row + 1):
            val = ws.cell(row=r, column=1).value
            if val is not None and str(val).strip() == str(id_value).strip():
                target_row = r
                break

        if not target_row:
            wb.close()
            return False

        for c_idx, col_name in enumerate(columns, start=1):
            if col_name in data:
                val = data[col_name]
                ws.cell(row=target_row, column=c_idx, value=val)

        wb.save(DATABASE_PATH)
        wb.close()
        return True

def delete_row(sheet_name, id_value):
    with _db_lock:
        wb = _load_workbook()
        ws = wb[sheet_name]
        target_row = None

        for r in range(2, ws.max_row + 1):
            val = ws.cell(row=r, column=1).value
            if val is not None and str(val).strip() == str(id_value).strip():
                target_row = r
                break

        if target_row:
            ws.delete_rows(target_row, 1)
            wb.save(DATABASE_PATH)
            wb.close()
            return True

        wb.close()
        return False

# Entity-specific helper methods
def get_searches():
    return get_all('Searches')

def get_search(search_id):
    return get_by_id('Searches', search_id)

def save_search(data):
    pk = data.get('search_id')
    if pk and get_by_id('Searches', pk):
        update_row('Searches', pk, data)
        return pk
    return insert_row('Searches', data)

def delete_search_cascade(search_id):
    premises = filter_rows('Premises', search_id=search_id)
    for p in premises:
        delete_premise_cascade(p['premise_id'])
    return delete_row('Searches', search_id)

def get_premises(search_id=None):
    if search_id:
        return filter_rows('Premises', search_id=search_id)
    return get_all('Premises')

def get_premise(premise_id):
    return get_by_id('Premises', premise_id)

def save_premise(data):
    pk = data.get('premise_id')
    if pk and get_by_id('Premises', pk):
        update_row('Premises', pk, data)
        return pk
    return insert_row('Premises', data)

def delete_premise_cascade(premise_id):
    evidences = filter_rows('Evidences', premise_id=premise_id)
    for ev in evidences:
        delete_evidence_cascade(ev['evidence_id'])
    clouds = filter_rows('CloudEvidences', premise_id=premise_id)
    for c in clouds:
        delete_row('CloudEvidences', c['cloud_evidence_id'])
    return delete_row('Premises', premise_id)

def get_evidences(premise_id=None):
    if premise_id:
        return filter_rows('Evidences', premise_id=premise_id)
    return get_all('Evidences')

def get_evidence(evidence_id):
    return get_by_id('Evidences', evidence_id)

def save_evidence(data):
    pk = data.get('evidence_id')
    if pk and get_by_id('Evidences', pk):
        update_row('Evidences', pk, data)
        return pk
    return insert_row('Evidences', data)

def delete_evidence_cascade(evidence_id):
    sub_evs = filter_rows('SubEvidences', evidence_id=evidence_id)
    for se in sub_evs:
        delete_row('SubEvidences', se['sub_evidence_id'])
    custodies = filter_rows('ChainOfCustody', evidence_id=evidence_id)
    for c in custodies:
        delete_row('ChainOfCustody', c['custody_id'])
    return delete_row('Evidences', evidence_id)

def get_sub_evidences(evidence_id):
    return filter_rows('SubEvidences', evidence_id=evidence_id)

def save_sub_evidence(data):
    pk = data.get('sub_evidence_id')
    if pk and get_by_id('SubEvidences', pk):
        update_row('SubEvidences', pk, data)
        return pk
    return insert_row('SubEvidences', data)

def delete_sub_evidence(sub_evidence_id):
    return delete_row('SubEvidences', sub_evidence_id)

def get_custody_rows(evidence_id):
    return filter_rows('ChainOfCustody', evidence_id=evidence_id)

def save_custody_row(data):
    pk = data.get('custody_id')
    if pk and get_by_id('ChainOfCustody', pk):
        update_row('ChainOfCustody', pk, data)
        return pk
    return insert_row('ChainOfCustody', data)

def delete_custody_row(custody_id):
    return delete_row('ChainOfCustody', custody_id)

def get_cloud_evidences(premise_id):
    return filter_rows('CloudEvidences', premise_id=premise_id)

def save_cloud_evidence(data):
    pk = data.get('cloud_evidence_id')
    if pk and get_by_id('CloudEvidences', pk):
        update_row('CloudEvidences', pk, data)
        return pk
    return insert_row('CloudEvidences', data)

def delete_cloud_evidence(cloud_evidence_id):
    return delete_row('CloudEvidences', cloud_evidence_id)
