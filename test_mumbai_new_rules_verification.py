import os
import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')

import config
import excel_db
import doc_generator
from app import app

def test_new_mumbai_rules():
    print("=" * 60)
    print("STARTING TEST FOR NEW MUMBAI RULES & JURISDICTION FORMS")
    print("=" * 60)

    excel_db.init_db()
    client = app.test_client()

    # 1. Test Form Route Dispatches for Mumbai, Delhi, Gujarat
    print("\n--- 1. Testing Jurisdiction Form Dispatches ---")
    
    # Create Mumbai Search & Premise
    s_mum_id = excel_db.save_search({
        'case_name': 'Operation Trident Mumbai (SEBI/MUM/2026/99)',
        'jurisdiction': 'Mumbai',
        'date': '2026-09-28',
        'authorised_officer': 'Shri R. K. Nair, General Manager',
        'remarks': 'Insider trading investigation',
        'search_warrant_details': 'SW/SEBI/WRO/2026/990'
    })
    p_mum_id = excel_db.save_premise({
        'search_id': s_mum_id,
        'premise_name': 'Trident Tower Office',
        'address': 'Nariman Point, Mumbai - 400021',
        'team': 'Team Mumbai-1',
        'date': '2026-09-28',
        'time': '10:00 AM'
    })

    # Create Delhi Search & Premise
    s_del_id = excel_db.save_search({
        'case_name': 'Operation Connaught (SEBI/DEL/2026/01)',
        'jurisdiction': 'Delhi',
        'date': '2026-09-28',
        'authorised_officer': 'Shri A. Sharma, GM',
        'remarks': 'Market manipulation'
    })
    p_del_id = excel_db.save_premise({
        'search_id': s_del_id,
        'premise_name': 'Barakhamba Office',
        'address': 'Connaught Place, New Delhi - 110001',
        'team': 'Team Delhi-1',
        'date': '2026-09-28',
        'time': '10:30 AM'
    })

    # Create Gujarat Search & Premise
    s_guj_id = excel_db.save_search({
        'case_name': 'Operation GIFT City (SEBI/GUJ/2026/01)',
        'jurisdiction': 'Gujarat',
        'date': '2026-09-28',
        'authorised_officer': 'Shri P. Patel, DGM',
        'remarks': 'Front running'
    })
    p_guj_id = excel_db.save_premise({
        'search_id': s_guj_id,
        'premise_name': 'GIFT One Tower',
        'address': 'GIFT City, Gandhinagar - 382355',
        'team': 'Team Guj-1',
        'date': '2026-09-28',
        'time': '11:00 AM'
    })

    # Check Mumbai Form GET
    resp_mum = client.get(f'/premise/{p_mum_id}/evidence/new')
    assert resp_mum.status_code == 200
    assert b"MUMBAI SPECIFIC" in resp_mum.data
    assert b"Where Found / Location in Premise" in resp_mum.data
    assert b"notes_by_officer" not in resp_mum.data, "notes_by_officer should be removed from Mumbai form!"
    print("  [PASS] Mumbai form renders dedicated template without notes_by_officer.")

    # Check Delhi Form GET
    resp_del = client.get(f'/premise/{p_del_id}/evidence/new')
    assert resp_del.status_code == 200
    assert b"DELHI SPECIFIC" in resp_del.data
    assert b"sub-evidences-section" in resp_del.data, "Delhi form must include sub-evidences section!"
    assert b"notes_by_officer" not in resp_del.data, "notes_by_officer should be removed from Delhi form!"
    print("  [PASS] Delhi form renders dedicated template with sub-evidences and no notes_by_officer.")

    # Check Gujarat Form GET
    resp_guj = client.get(f'/premise/{p_guj_id}/evidence/new')
    assert resp_guj.status_code == 200
    assert b"GUJARAT SPECIFIC" in resp_guj.data
    assert b"sub-evidences-section" not in resp_guj.data, "Gujarat form should not have Delhi sub-evidences clutter!"
    assert b"notes_by_officer" not in resp_guj.data, "notes_by_officer should be removed from Gujarat form!"
    print("  [PASS] Gujarat form renders dedicated template without notes_by_officer.")

    # 2. Test Evidence Creation with Shared Master/Working Hard Drives in Mumbai
    print("\n--- 2. Creating Mumbai Evidences (Testing Unique HDD Sharing) ---")
    
    # Evidence 1: Laptop - Owner is YES (Name: Suresh Rao)
    # Target Copies: Master S/N: WD-134567890, Working S/N: SG-2345678909
    ev1_id = excel_db.save_evidence({
        'premise_id': p_mum_id,
        'assessee_name': 'Trident Capital LLP',
        'evidence_category': 'Laptop',
        'device_name': 'Managing Director Laptop',
        'location_found': 'MD Cabin, 4th Floor',
        'make': 'Dell',
        'model': 'Latitude 7420',
        'serial_number': 'DELL-LAT-7420-SN1',
        'capacity': '1 TB SSD',
        'color': 'Carbon Black',
        'os': 'Windows 11 Enterprise',
        'owner_user': 'Suresh Rao',
        'is_owner': 'Yes',
        'owner_relation': 'Managing Director',
        'md5_hash': '11112222333344445555666677778888',
        'sha1_hash': 'aaaabbbbccccddddeeeeffff0000111122223333',
        'sha256_hash_1': 'abcdef1234567890abcdef1234567890abcdef1234567890abcdef1234567890',
        'master_copy_make': 'Western Digital',
        'master_copy_model': 'Ultrastar HDD',
        'master_copy_sn': '134567890',
        'master_copy_capacity': '4 TB',
        'working_copy_make': 'Seagate',
        'working_copy_model': 'Exos HDD',
        'working_copy_sn': '2345678909',
        'working_copy_capacity': '4 TB',
        'collection_date': '2026-09-28',
        'collection_time': '10:15 AM',
        'collection_address': 'Nariman Point, Mumbai - 400021',
        'expert_name': 'Kunal Sen',
        'expert_father_name': 'Shri D. N. Sen',
        'expert_designation': 'Digital Forensic Expert Grade I',
        'cert_place': 'Nariman Point, Mumbai',
        'party_name': 'Suresh Rao',
        'authorised_officer': 'Shri R. K. Nair, General Manager'
    })

    # Evidence 2: Desktop - Owner is NO (in possession of trader Ramesh, owner is Company)
    # Target Copies: SHARES THE SAME Master S/N: 134567890, Working S/N: 2345678909!
    ev2_id = excel_db.save_evidence({
        'premise_id': p_mum_id,
        'assessee_name': 'Trident Capital LLP',
        'evidence_category': 'Desktop',
        'device_name': 'Trading Terminal 01',
        'location_found': 'Trading Floor, Desk 12',
        'make': 'HP',
        'model': 'EliteDesk 800',
        'serial_number': 'HP-ED-0012',
        'capacity': '512 GB SSD',
        'color': 'Black',
        'os': 'Ubuntu Linux 24.04',
        'owner_user': 'Ramesh Shinde',
        'is_owner': 'No',
        'owner_relation': 'Derivatives Trader',
        'md5_hash': '22223333444455556666777788889999',
        'sha1_hash': 'bbbbccccddddeeeeffff00001111222233334444',
        'sha256_hash_1': 'bcdef1234567890abcdef1234567890abcdef1234567890abcdef1234567890a',
        'master_copy_make': 'Western Digital',
        'master_copy_model': 'Ultrastar HDD',
        'master_copy_sn': '134567890',
        'master_copy_capacity': '4 TB',
        'working_copy_make': 'Seagate',
        'working_copy_model': 'Exos HDD',
        'working_copy_sn': '2345678909',
        'working_copy_capacity': '4 TB',
        'collection_date': '2026-09-28',
        'collection_time': '10:30 AM',
        'collection_address': 'Nariman Point, Mumbai - 400021',
        'expert_name': 'Kunal Sen',
        'expert_father_name': 'Shri D. N. Sen',
        'expert_designation': 'Digital Forensic Expert Grade I',
        'cert_place': 'Nariman Point, Mumbai',
        'party_name': 'Suresh Rao',
        'authorised_officer': 'Shri R. K. Nair, General Manager'
    })

    # Evidence 3: Cloud Evidence - ONLY Email ID!
    # Target Copies: A DIFFERENT hard drive pair: Master S/N: 234567890, Working S/N: 456789078
    ev3_id = excel_db.save_evidence({
        'premise_id': p_mum_id,
        'assessee_name': 'Trident Capital LLP',
        'evidence_category': 'Cloud',
        'device_name': 'Director Cloud Mailbox (M365)',
        'location_found': 'Microsoft Cloud Tenant',
        'email_id': 'suresh.rao@tridentcapital.in',
        'owner_user': 'Suresh Rao',
        'is_owner': 'Yes',
        'owner_relation': 'Managing Director',
        'md5_hash': '33334444555566667777888899990000',
        'sha1_hash': 'ccccddddeeeeffff000011112222333344445555',
        'sha256_hash_1': 'cdef1234567890abcdef1234567890abcdef1234567890abcdef1234567890ab',
        'master_copy_make': 'SanDisk',
        'master_copy_model': 'Extreme PRO SSD',
        'master_copy_sn': '234567890',
        'master_copy_capacity': '2 TB',
        'working_copy_make': 'Crucial',
        'working_copy_model': 'X9 Pro SSD',
        'working_copy_sn': '456789078',
        'working_copy_capacity': '2 TB',
        'collection_date': '2026-09-28',
        'collection_time': '10:45 AM',
        'collection_address': 'Nariman Point, Mumbai - 400021',
        'expert_name': 'Kunal Sen',
        'expert_father_name': 'Shri D. N. Sen',
        'expert_designation': 'Digital Forensic Expert Grade I',
        'cert_place': 'Nariman Point, Mumbai',
        'party_name': 'Suresh Rao',
        'authorised_officer': 'Shri R. K. Nair, General Manager'
    })

    print(f"  Created 3 Mumbai evidences: Laptop={ev1_id}, Desktop={ev2_id}, Cloud={ev3_id}")

    # 3. Test Annexure 1 Rules
    print("\n--- 3. Testing Mumbai Annexure 1 (Physical Inventory) ---")
    a1_file = doc_generator.generate_mumbai_annexure_1(p_mum_id)
    doc_a1 = docx.Document(a1_file)
    t_a1 = doc_a1.tables[0]
    
    # Locate Laptop row
    lap_row = None
    desk_row = None
    cloud_row = None
    for r in t_a1.rows[1:]:
        c2 = r.cells[2].text
        if "Latitude 7420" in c2:
            lap_row = r
        elif "EliteDesk 800" in c2:
            desk_row = r
        elif "suresh.rao@tridentcapital.in" in c2:
            cloud_row = r

    assert lap_row is not None, "Laptop row missing in Annexure 1"
    assert desk_row is not None, "Desktop row missing in Annexure 1"
    assert cloud_row is not None, "Cloud row missing in Annexure 1"

    # RULE 1: Nature of Evidence / Where Found in Column 1
    print(f"  Laptop Col 1 (Nature/Where): {lap_row.cells[1].text}")
    assert "MD Cabin, 4th Floor" in lap_row.cells[1].text, "Column 1 must have where found location!"
    print(f"  Desktop Col 1 (Nature/Where): {desk_row.cells[1].text}")
    assert "Trading Floor, Desk 12" in desk_row.cells[1].text, "Column 1 must have where found location!"
    print("  [PASS] Column 1 ('Nature Of Evidence Found') correctly contains Where Found location.")

    # RULE 2: Model Details in Column 2 must contain Evidence Name
    print(f"  Laptop Col 2: \n{lap_row.cells[2].text}")
    assert "Evidence Name: Managing Director Laptop" in lap_row.cells[2].text, "Evidence Name must be in Column 2!"
    assert "Evidence Name: Trading Terminal 01" in desk_row.cells[2].text, "Evidence Name must be in Column 2!"
    assert "Evidence Name: Director Cloud Mailbox (M365)" in cloud_row.cells[2].text, "Evidence Name must be in Column 2!"
    print("  [PASS] Column 2 ('Model Details') starts with Evidence Name.")

    # RULE 3: Is Owner: If Yes -> Name; If No -> 'No'
    print(f"  Laptop Col 4 (Is Owner): '{lap_row.cells[4].text}' (Expected: Suresh Rao)")
    assert lap_row.cells[4].text == "Suresh Rao", "Is Owner must show Name when Yes!"
    print(f"  Desktop Col 4 (Is Owner): '{desk_row.cells[4].text}' (Expected: No)")
    assert desk_row.cells[4].text == "No", "Is Owner must show 'No' when No!"
    print("  [PASS] Column 4 ('Is the Owner of the Device?') displays Name if Yes, and 'No' if No.")

    # RULE 4: Cloud in Annexure 1 has ONLY Email ID (no tenant, no IP, no domain)
    print(f"  Cloud Col 2 Details: \n{cloud_row.cells[2].text}")
    assert "Email ID: suresh.rao@tridentcapital.in" in cloud_row.cells[2].text
    assert "Cloud ID:" not in cloud_row.cells[2].text, "Cloud ID must NOT be present!"
    assert "IP:" not in cloud_row.cells[2].text, "IP must NOT be present!"
    assert "Domain:" not in cloud_row.cells[2].text, "Domain must NOT be present!"
    assert "Color:" not in cloud_row.cells[2].text, "Color must NOT be present!"
    print("  [PASS] Cloud evidence contains ONLY Email ID in Annexure 1.")

    # RULE 5: Remarks in Column 6 is empty
    assert lap_row.cells[6].text == ""
    assert desk_row.cells[6].text == ""
    assert cloud_row.cells[6].text == ""
    print("  [PASS] Remarks in Column 6 is left blank for manual on-site notation.")

    # 4. Test Annexure 3 COC Unique Master & Working Copies
    print("\n--- 4. Testing Mumbai Annexure 3 (Unique COC Register) ---")
    a3_file = doc_generator.generate_mumbai_annexure_3_coc(p_mum_id)
    doc_a3 = docx.Document(a3_file)
    t1_coc = doc_a3.tables[1]
    
    t1_text = t1_coc.rows[2].cells[0].text
    t2_text = t1_coc.rows[3].cells[0].text
    print(f"  Transfer 1 Evidence Devices Content:\n{t1_text}\n")

    # There were 3 evidences, but ONLY 2 UNIQUE Master/Working copy pairs:
    # Pair 1: WD-134567890 / SG-2345678909 (shared by Laptop and Desktop)
    # Pair 2: 234567890 / 456789078 (Cloud)
    assert "Evidence 1:" in t1_text
    assert "Master copy S/N: 134567890" in t1_text
    assert "Working copy S/N: 2345678909" in t1_text
    assert "Evidence 2:" in t1_text
    assert "Master copy S/N: 234567890" in t1_text
    assert "Working copy S/N: 456789078" in t1_text
    assert "Evidence 3:" not in t1_text, "There are only 2 UNIQUE HDD copy pairs; Evidence 3 must NOT be present!"
    
    assert "Master copy S/N: 134567890" in t2_text
    assert "Working copy S/N: 2345678909" in t2_text
    assert "Master copy S/N: 234567890" in t2_text
    assert "Working copy S/N: 456789078" in t2_text
    print("  [PASS] Annexure 3 COC Table 1 lists EXACTLY the Unique Master & Working copy hard drives!")

    # 5. Test Removal of notes_by_officer from generated documents
    print("\n--- 5. Testing Document Generation without notes_by_officer ---")
    doc_res1 = doc_generator.generate_evidence_docs(ev1_id)
    zip_path = doc_res1['zip_path']
    import zipfile
    with zipfile.ZipFile(zip_path, 'r') as zf:
        file_names = zf.namelist()
        print("  ZIP contents:", file_names)
        for fn in file_names:
            content = zf.read(fn).decode('latin1', errors='ignore')
            assert "notes_by_officer" not in content, f"Placeholder notes_by_officer still present in {fn}!"
    print("  [PASS] notes_by_officer is completely removed from document templates and generated output.")

    print("\n" + "=" * 60)
    print("ALL NEW MUMBAI SPECIFIC RULES VERIFIED WITH 100% SUCCESS!")
    print("=" * 60)

if __name__ == '__main__':
    test_new_mumbai_rules()
