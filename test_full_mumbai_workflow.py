import os
import sys
# pyrefly: ignore [missing-import]
import docx

sys.stdout.reconfigure(encoding='utf-8')

import config
import excel_db
import doc_generator

def run_test():
    print("Initializing Database...")
    excel_db.init_db()

    # Create Mumbai Search
    search_data = {
        'case_name': 'Operation Marine Drive (SEBI/MUM/2026/01)',
        'jurisdiction': 'Mumbai',
        'date': '2026-09-26',
        'authorised_officer': 'Shri V. K. Malhotra, General Manager',
        'remarks': 'Investigation in front-running and unauthorized trading algorithms',
        'search_warrant_details': 'SW/SEBI/WRO/MUM/2026/009 dt. 25/09/2026'
    }
    search_id = excel_db.save_search(search_data)
    print(f"Created Mumbai Search: {search_id}")

    # Create Premise
    premise_data = {
        'search_id': search_id,
        'premise_name': 'Corporate HQ (Marine Equities Pvt Ltd)',
        'address': 'Plot C-54, G Block, Bandra Kurla Complex (BKC), Mumbai - 400051',
        'team': 'Team Alpha-MUM',
        'date': '2026-09-26',
        'time': '10:00 AM'
    }
    premise_id = excel_db.save_premise(premise_data)
    print(f"Created Premise: {premise_id}")

    # Create Evidence 1: Laptop
    ev1_data = {
        'premise_id': premise_id,
        'assessee_name': 'Marine Equities Pvt Ltd',
        'evidence_category': 'Laptop',
        'device_name': 'MD Trading Laptop (Lenovo)',
        'make': 'Lenovo',
        'model': 'ThinkPad P1 Gen 6',
        'serial_number': 'SN-LEN-99120',
        'capacity': '2 TB NVMe SSD',
        'color': 'Midnight Black',
        'owner_user': 'Vikramaditya Singhania',
        'is_owner': 'Yes',
        'owner_relation': 'Managing Director & Promoter',
        'os': 'Windows 11 Pro 64-bit',
        'email_id': 'vikram@marineequities.in',
        'cloud_id': 'marine_m365_001',
        'ip_address': '10.0.1.44',
        'domain_name': 'marineequities.in',
        'hibernation_sleep': 'On',
        'visible_on_screen': 'Bloomberg Terminal and Trading Console',
        'shutdown_type': 'PowerPlugPulled',
        'encryption_present': 'Y',
        'encryption_software': 'BitLocker (TPM 2.0)',
        'device_handling': 'Seizure, Imaging',
        'imaging_tools_details': 'Tableau Forensic TX1 v23.4 / FTK Imager v4.7',
        'md5_hash': 'a1b2c3d4e5f60718293a4b5c6d7e8f90',
        'sha1_hash': '1234567890abcdef1234567890abcdef12345678',
        'sha256_hash_1': 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
        'master_copy_make': 'Western Digital',
        'master_copy_model': 'My Passport Enterprise',
        'master_copy_sn': 'WD-MUM-001-M',
        'master_copy_capacity': '4 TB USB 3.2',
        'working_copy_make': 'Seagate',
        'working_copy_model': 'One Touch SSD',
        'working_copy_sn': 'SG-MUM-001-W',
        'working_copy_capacity': '4 TB USB-C',
        'device_returned': 'N',
        'notes_by_officer': 'Tamper-evident barcode seal #SEBI-MUM-1001 placed on port enclosures in presence of panchas.',
        'examiner_name': 'Ketan Parekh',
        'examiner_details': 'Senior Digital Forensic Examiner, Cert. EnCE, CHFI',
        'expert_name': 'Ketan Parekh',
        'expert_designation': 'Senior Digital Forensic Examiner',
        'cert_place': 'BKC, Mumbai',
        'collection_date': '2026-09-26',
        'collection_time': '10:30 AM',
        'collection_address': 'Plot C-54, G Block, BKC, Mumbai - 400051',
        'authorised_officer': 'Shri V. K. Malhotra, General Manager',
        'hash_algorithm': 'SHA256'
    }
    ev1_id = excel_db.save_evidence(ev1_data)
    print(f"Created Evidence 1: {ev1_id} (Laptop)")

    # Create Evidence 2: Mobile
    ev2_data = {
        'premise_id': premise_id,
        'assessee_name': 'Marine Equities Pvt Ltd',
        'evidence_category': 'Mobile',
        'device_name': 'Promoter iPhone 15 Pro Max',
        'make': 'Apple',
        'model': 'iPhone 15 Pro Max',
        'serial_number': 'APL-IPH-88219',
        'capacity': '512 GB',
        'color': 'Blue Titanium',
        'owner_user': 'Vikramaditya Singhania',
        'is_owner': 'Yes',
        'owner_relation': 'Managing Director & Promoter',
        'os': 'iOS 17.5.1',
        'imei1': '359123456789012',
        'imei2': '359123456789013',
        'email_id': 'vikram.personal@icloud.com',
        'cloud_id': 'apple_id_vikram_singhania',
        'ip_address': '10.0.1.89',
        'domain_name': '',
        'sim_card_present': 'Y',
        'sim_provider': 'Airtel 5G eSIM',
        'mobile_state': 'Offline',
        'hibernation_sleep': 'Off',
        'visible_on_screen': 'Lock screen with WhatsApp message notifications',
        'device_handling': 'Seizure, Imaging',
        'imaging_tools_details': 'Cellebrite UFED 4PC v7.68 / Premium Advanced Logical',
        'md5_hash': 'b2c3d4e5f60718293a4b5c6d7e8f90a1',
        'sha1_hash': '234567890abcdef1234567890abcdef123456789',
        'sha256_hash_1': 'ca978112ca1bbdcafac231b39a23dc4da786eff8147c4e72b9807785afee48bb',
        'sha256_hash_2': '88d4266fd4e6338d13b845fcf289579d209c897823b9217da3e161936f031589',
        'master_copy_make': 'Western Digital',
        'master_copy_model': 'My Passport Enterprise',
        'master_copy_sn': 'WD-MUM-002-M',
        'master_copy_capacity': '2 TB USB 3.2',
        'working_copy_make': 'Seagate',
        'working_copy_model': 'One Touch SSD',
        'working_copy_sn': 'SG-MUM-002-W',
        'working_copy_capacity': '2 TB USB-C',
        'device_returned': 'N',
        'notes_by_officer': 'Faraday bag #SEBI-FB-202 sealed immediately upon seizure.',
        'examiner_name': 'Ketan Parekh',
        'examiner_details': 'Senior Digital Forensic Examiner, Cert. EnCE, CHFI',
        'expert_name': 'Ketan Parekh',
        'expert_designation': 'Senior Digital Forensic Examiner',
        'cert_place': 'BKC, Mumbai',
        'collection_date': '2026-09-26',
        'collection_time': '10:45 AM',
        'collection_address': 'Plot C-54, G Block, BKC, Mumbai - 400051',
        'authorised_officer': 'Shri V. K. Malhotra, General Manager',
        'hash_algorithm': 'SHA256'
    }
    ev2_id = excel_db.save_evidence(ev2_data)
    print(f"Created Evidence 2: {ev2_id} (Mobile)")

    # Create Evidence 3: Other / Cloud Email
    ev3_data = {
        'premise_id': premise_id,
        'assessee_name': 'Marine Equities Pvt Ltd',
        'evidence_category': 'Other',
        'device_name': 'CEO Exchange Mailbox Backup',
        'make': 'Microsoft M365',
        'model': 'Exchange Online Cloud Archive',
        'serial_number': 'M365-EXCH-DIR-09',
        'capacity': '85 GB PST/EDRM',
        'color': 'Cloud Asset',
        'owner_user': 'Vikramaditya Singhania',
        'is_owner': 'No',
        'owner_relation': 'User / Account Holder',
        'os': 'Cloud M365 Service',
        'email_id': 'vikram.singhania@marineequities.in',
        'cloud_id': 'marineequities_tenant_guid_8892',
        'ip_address': '40.97.12.33',
        'domain_name': 'marineequities.in',
        'hibernation_sleep': 'Off',
        'visible_on_screen': 'Admin Compliance Center eDiscovery Export',
        'device_handling': 'Imaging, Backup',
        'imaging_tools_details': 'FTK Imager v4.7 Cloud Agent / Purview Export Tool',
        'md5_hash': 'c3d4e5f60718293a4b5c6d7e8f90a1b2',
        'sha1_hash': '34567890abcdef1234567890abcdef1234567890',
        'sha256_hash_1': '4b227777d4dd1fc61c6f884f48641d02b4d121d3fd328cb08b5531fcacdabf8a',
        'master_copy_make': 'Western Digital',
        'master_copy_model': 'My Passport Enterprise',
        'master_copy_sn': 'WD-MUM-003-M',
        'master_copy_capacity': '1 TB',
        'working_copy_make': 'Seagate',
        'working_copy_model': 'One Touch SSD',
        'working_copy_sn': 'SG-MUM-003-W',
        'working_copy_capacity': '1 TB',
        'device_returned': 'N',
        'notes_by_officer': 'Live cloud mailbox export completed under panchnama.',
        'examiner_name': 'Ketan Parekh',
        'examiner_details': 'Senior Digital Forensic Examiner, Cert. EnCE, CHFI',
        'expert_name': 'Ketan Parekh',
        'expert_designation': 'Senior Digital Forensic Examiner',
        'cert_place': 'BKC, Mumbai',
        'collection_date': '2026-09-26',
        'collection_time': '11:15 AM',
        'collection_address': 'Plot C-54, G Block, BKC, Mumbai - 400051',
        'authorised_officer': 'Shri V. K. Malhotra, General Manager',
        'hash_algorithm': 'SHA256'
    }
    ev3_id = excel_db.save_evidence(ev3_data)
    print(f"Created Evidence 3: {ev3_id} (Other / Mail)")

    print("\n--- Testing Device-Level Document Generation for Evidence 1 (Laptop) ---")
    res1 = doc_generator.generate_evidence_docs(ev1_id)
    print("Evidence 1 generated files:", len(res1['files']), "Zip:", res1.get('zip_path'))
    assert os.path.exists(res1['zip_path']), "Zip file not found for Laptop"

    print("\n--- Testing Device-Level Document Generation for Evidence 2 (Mobile) ---")
    res2 = doc_generator.generate_evidence_docs(ev2_id)
    print("Evidence 2 generated files:", len(res2['files']), "Zip:", res2.get('zip_path'))
    assert os.path.exists(res2['zip_path']), "Zip file not found for Mobile"

    print("\n--- Testing Device-Level Document Generation for Evidence 3 (Other) ---")
    res3 = doc_generator.generate_evidence_docs(ev3_id)
    print("Evidence 3 generated files:", len(res3['files']), "Zip:", res3.get('zip_path'))
    assert os.path.exists(res3['zip_path']), "Zip file not found for Other"

    print("\n--- Testing Premise-Level Annexure 1 (Physical Inventory) ---")
    a1_path = doc_generator.generate_mumbai_annexure_1(premise_id)
    print("Annexure 1 Path:", a1_path)
    assert os.path.exists(a1_path), "Annexure 1 docx not found"
    doc_a1 = docx.Document(a1_path)
    print(f"Annexure 1 has {len(doc_a1.tables)} table(s). Expected: 1 (single consolidated table)")
    assert len(doc_a1.tables) == 1, f"Expected 1 table in Annexure 1, got {len(doc_a1.tables)}"
    t_a1 = doc_a1.tables[0]
    print(f"Annexure 1 Table rows: {len(t_a1.rows)}")
    for r_i, r in enumerate(t_a1.rows):
        print(f"  R{r_i}: C0={repr(r.cells[0].text)} | C1={repr(r.cells[1].text)} | C2={repr(r.cells[2].text.replace(chr(10), '; ')[:45])}")

    print("\n--- Testing Premise-Level Annexure 3 (Mumbai Chain of Custody) ---")
    a3_path = doc_generator.generate_mumbai_annexure_3_coc(premise_id)
    print("Annexure 3 Path:", a3_path)
    assert os.path.exists(a3_path), "Annexure 3 docx not found"
    doc_a3 = docx.Document(a3_path)
    print(f"Annexure 3 Table 0 rows: {len(doc_a3.tables[0].rows)}")
    t1_coc = doc_a3.tables[1]
    print(f"Annexure 3 Table 1 rows: {len(t1_coc.rows)}, cols: {len(t1_coc.columns)}")
    assert len(t1_coc.columns) == 6, f"Expected 6 columns in Table 1, got {len(t1_coc.columns)}"
    
    t1_text = t1_coc.rows[2].cells[0].text
    t2_text = t1_coc.rows[3].cells[0].text
    print("\n--- COC Transfer 1 Text Sample ---\n" + t1_text[:250])
    assert "Master copy S/N:" in t1_text, "Master copy S/N: missing in Transfer 1"
    assert "Working copy S/N:" in t1_text, "Working copy S/N: missing in Transfer 1"
    assert "Master copy S/N:" in t2_text, "Master copy S/N: missing in Transfer 2"
    assert "Working copy S/N:" in t2_text, "Working copy S/N: missing in Transfer 2"
    
    print("\nSUCCESS! All Mumbai tests passed with 100% precision!")

if __name__ == '__main__':
    run_test()
