import os
import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')

import config
import excel_db
import doc_generator

def test_detailed_rules():
    print("Initializing Database...")
    excel_db.init_db()

    # Create Search
    search_data = {
        'case_name': 'Operation Dalal Street (SEBI/MUM/2026/02)',
        'jurisdiction': 'Mumbai',
        'date': '2026-09-26',
        'authorised_officer': 'Shri S. Raman, Executive Director',
        'remarks': 'Investigation in algorithmic spoofing',
        'search_warrant_details': 'SW/SEBI/WRO/2026/778'
    }
    search_id = excel_db.save_search(search_data)
    print(f"Created Mumbai Search: {search_id}")

    # Create Premise
    premise_data = {
        'search_id': search_id,
        'premise_name': 'Algorithmic Trading Floor',
        'address': 'Level 14, Tower 2, One International Centre, Prabhadevi, Mumbai - 400013',
        'team': 'Team Charlie-MUM',
        'date': '2026-09-26',
        'time': '10:30 AM'
    }
    premise_id = excel_db.save_premise(premise_data)
    print(f"Created Premise: {premise_id}")

    # 1. Create 3 Desktops (to test rowspan 3 and single document inventory)
    desktop_ids = []
    for d_idx in range(1, 4):
        d_data = {
            'premise_id': premise_id,
            'assessee_name': 'Dalal Algorithms LLP',
            'evidence_category': 'Desktop',
            'device_name': f'Algo Trading Terminal 0{d_idx}',
            'make': 'Dell',
            'model': 'OptiPlex 7090 Tower',
            'serial_number': f'DELL-TRD-00{d_idx}',
            'capacity': '1 TB NVMe SSD',
            'color': 'Charcoal Black',
            'owner_user': f'Trader {d_idx}',
            'is_owner': 'No',
            'owner_relation': 'Proprietary Trader',
            'os': 'Ubuntu 24.04 LTS / Low Latency Kernel',
            'hibernation_sleep': 'Off',
            'shutdown_type': 'Normal',
            'visible_on_screen': 'Low latency trading engine terminal',
            'imaging_tools_details': 'Tableau TX1 Hardware Imager',
            'md5_hash': f'd{d_idx}d{d_idx}d{d_idx}d{d_idx}000011112222333344445555',
            'sha1_hash': f'11223344556677889900aabbccddeeff0011223{d_idx}',
            'sha256_hash_1': f'aaabbbcccdddeeefff000111222333444555666777888999000aaabbbcccddd{d_idx}',
            'master_copy_make': 'Western Digital',
            'master_copy_model': 'Ultrastar Enterprise',
            'master_copy_sn': f'WD-MUM-MC-00{d_idx}',
            'master_copy_capacity': '2 TB',
            'working_copy_make': 'Seagate',
            'working_copy_model': 'Exos 7E8',
            'working_copy_sn': f'SG-MUM-WC-00{d_idx}',
            'working_copy_capacity': '2 TB',
            'notes_by_officer': f'Found connected to DMA trading switch port {d_idx}',
            'other_specific_notes': f'Algorithmic trading node {d_idx} running automated strategy',
            'expert_name': 'Rajeev Verma',
            'expert_father_name': 'Shri Rameshwar Verma',
            'expert_designation': 'Forensic Expert Grade I',
            'cert_place': 'Prabhadevi, Mumbai',
            'collection_date': '2026-09-26',
            'collection_time': '10:45 AM',
            'collection_address': 'Level 14, Tower 2, One International Centre, Mumbai - 400013',
            'authorised_officer': 'Shri S. Raman, Executive Director',
            'hash_algorithm': 'SHA256'
        }
        ev_d_id = excel_db.save_evidence(d_data)
        desktop_ids.append(ev_d_id)
    print(f"Created 3 Desktops: {desktop_ids}")

    # 2. Create Laptop (to test no IMEI in Annexure 1)
    laptop_data = {
        'premise_id': premise_id,
        'assessee_name': 'Dalal Algorithms LLP',
        'evidence_category': 'Laptop',
        'device_name': 'CTO Quantitative Laptop',
        'make': 'Lenovo',
        'model': 'ThinkPad X1 Carbon Gen 11',
        'serial_number': 'LEN-CTO-8819',
        'capacity': '2 TB SSD',
        'color': 'Carbon Weave Black',
        'owner_user': 'Dr. K. Srinivas',
        'is_owner': 'Yes',
        'owner_relation': 'Chief Technology Officer & Partner',
        'os': 'Windows 11 Pro 64-bit',
        'notes_by_officer': 'Retrieved from executive backpack with tamper-evident seal #7781',
        'other_specific_notes': 'Contains model backtesting scripts and execution code',
        'master_copy_sn': 'WD-MUM-MC-004',
        'working_copy_sn': 'SG-MUM-WC-004',
        'expert_name': 'Rajeev Verma',
        'expert_designation': 'Forensic Expert Grade I',
        'collection_date': '2026-09-26',
        'collection_time': '11:00 AM'
    }
    lap_id = excel_db.save_evidence(laptop_data)
    print(f"Created Laptop: {lap_id}")

    # 3. Create Mobile (to test IMEI in Annexure 1 & Mobile form)
    mobile_data = {
        'premise_id': premise_id,
        'assessee_name': 'Dalal Algorithms LLP',
        'evidence_category': 'Mobile',
        'device_name': 'CTO iPhone 15 Pro',
        'make': 'Apple',
        'model': 'iPhone 15 Pro',
        'serial_number': 'APL-CTO-IP15',
        'capacity': '512 GB',
        'color': 'Blue Titanium',
        'os': 'iOS 17.6',
        'imei1': '865432098765432',
        'imei2': '865432098765431',
        'sim_card_present': 'Y',
        'sim_provider': 'Airtel Postpaid',
        'imsi': '404450123456789',
        'mobile_state': 'Offline',
        'owner_user': 'Dr. K. Srinivas',
        'is_owner': 'Yes',
        'owner_relation': 'Partner / CTO',
        'master_copy_sn': 'WD-MUM-MC-005',
        'working_copy_sn': 'SG-MUM-WC-005',
        'notes_by_officer': 'Faraday pouch sealed with barcoded tamper tape #9901',
        'other_specific_notes': 'Primary mobile used for multi-factor authentication',
        'expert_name': 'Rajeev Verma',
        'expert_designation': 'Forensic Expert Grade I',
        'collection_date': '2026-09-26',
        'collection_time': '11:15 AM'
    }
    mob_id = excel_db.save_evidence(mobile_data)
    print(f"Created Mobile: {mob_id}")

    # 4. Create Cloud Evidence (to test Cloud checkbox in Annexure E vs Other in Non-Mobile form)
    cloud_data = {
        'premise_id': premise_id,
        'assessee_name': 'Dalal Algorithms LLP',
        'evidence_category': 'Cloud',
        'device_name': 'AWS S3 Algorithmic Archive',
        'make': 'Amazon Web Services (AWS)',
        'model': 'S3 Bucket (asia-south1)',
        'serial_number': 'arn:aws:s3:::dalal-algo-archive-2026',
        'capacity': '350 GB',
        'email_id': 'aws-admin@dalalalgo.com',
        'cloud_id': 'aws_account_id_990182746192',
        'domain_name': 'dalalalgo.com',
        'owner_user': 'Dr. K. Srinivas',
        'is_owner': 'No',
        'owner_relation': 'AWS IAM Root Admin',
        'master_copy_sn': 'WD-MUM-MC-006',
        'working_copy_sn': 'SG-MUM-WC-006',
        'notes_by_officer': 'Cloud IAM log export and bucket mirror completed under panchnama',
        'other_specific_notes': 'Cloud repository storing live tick-by-tick order placement logs',
        'expert_name': 'Rajeev Verma',
        'expert_designation': 'Forensic Expert Grade I',
        'collection_date': '2026-09-26',
        'collection_time': '11:30 AM'
    }
    cloud_id = excel_db.save_evidence(cloud_data)
    print(f"Created Cloud Evidence: {cloud_id}")

    # 5. Create File/Folder Evidence (to test File/Folder checkbox in Non-Mobile form vs Other in Annexure E)
    file_data = {
        'premise_id': premise_id,
        'assessee_name': 'Dalal Algorithms LLP',
        'evidence_category': 'File/Folder',
        'device_name': 'Strategy Source Code Directory',
        'make': 'Git Repository Export',
        'model': 'Directory tree (C:/Strategies)',
        'serial_number': 'DIR-HASH-GIT-4401',
        'capacity': '12 GB',
        'owner_user': 'Senior Quant Team',
        'is_owner': 'Yes',
        'owner_relation': 'Internal Proprietary Models',
        'master_copy_sn': 'WD-MUM-MC-007',
        'working_copy_sn': 'SG-MUM-WC-007',
        'notes_by_officer': 'Extracted folder verified with SHA-256 manifest',
        'other_specific_notes': 'Target strategy folder containing Python/C++ HFT execution modules',
        'expert_name': 'Rajeev Verma',
        'expert_designation': 'Forensic Expert Grade I',
        'collection_date': '2026-09-26',
        'collection_time': '11:45 AM'
    }
    file_id = excel_db.save_evidence(file_data)
    print(f"Created File/Folder Evidence: {file_id}")

    print("\n=======================================================")
    print("VERIFYING RULE 1: Annexure 1 Single Document & Rowspan")
    print("=======================================================")
    a1_path = doc_generator.generate_mumbai_annexure_1(premise_id)
    assert os.path.exists(a1_path), "Annexure 1 docx not found"
    doc_a1 = docx.Document(a1_path)
    assert len(doc_a1.tables) == 1, f"Expected 1 table in Annexure 1, got {len(doc_a1.tables)}"
    t_a1 = doc_a1.tables[0]
    print(f"Annexure 1 Table has {len(t_a1.rows)} rows.")
    
    # Check that Row 1, 2, 3 have Desktop
    # Row 1 is start of Desktop, Row 2 and 3 are in the same Desktop merge group
    d1_cell0 = t_a1.cell(1, 0).text
    d2_cell0 = t_a1.cell(2, 0).text
    d3_cell0 = t_a1.cell(3, 0).text
    print(f"Desktop cell 0 texts: R1={repr(d1_cell0)}, R2={repr(d2_cell0)}, R3={repr(d3_cell0)}")
    assert d1_cell0 == 'Desktop', f"Expected 'Desktop', got {repr(d1_cell0)}"
    assert d2_cell0 == 'Desktop', f"Expected merged 'Desktop', got {repr(d2_cell0)}"
    assert d3_cell0 == 'Desktop', f"Expected merged 'Desktop', got {repr(d3_cell0)}"
    
    # Check that Evidence Name is in Model Details (column 2) and Where Found is in column 1
    print(f"Desktop Col 1 (Where found): R1={t_a1.cell(1, 1).text}")
    print(f"Desktop Col 2 (Model Details with Name): R1={t_a1.cell(1, 2).text}")
    assert "Terminal 01" in t_a1.cell(1, 2).text
    assert "Terminal 02" in t_a1.cell(2, 2).text
    assert "Terminal 03" in t_a1.cell(3, 2).text
    assert "Evidence Name: Algo Trading Terminal 01" in t_a1.cell(1, 2).text
    # Check Is Owner: For Desktops is_owner was No -> shows 'No'
    assert t_a1.cell(1, 4).text == "No"

    # Check Model Details of Laptop (NO IMEI)
    laptop_row = None
    for r in t_a1.rows:
        if "ThinkPad X1" in r.cells[2].text:
            laptop_row = r
            break
    assert laptop_row is not None, "Laptop row not found in Annexure 1"
    lap_details = laptop_row.cells[2].text
    print("\nLaptop Model Details in Annexure 1:\n" + lap_details)
    assert "IMEI" not in lap_details, "CRITICAL: Laptop details MUST NOT contain IMEI!"
    assert "Make/Model: Lenovo ThinkPad" in lap_details
    assert "S/N: LEN-CTO-8819" in lap_details
    assert "Evidence Name: CTO Quantitative Laptop" in lap_details
    assert laptop_row.cells[4].text == 'Dr. K. Srinivas', "Is Owner column MUST show owner name when is_owner is Yes!"

    # Check Model Details of Mobile (MUST HAVE IMEI)
    mobile_row = None
    for r in t_a1.rows:
        if "iPhone 15 Pro" in r.cells[2].text:
            mobile_row = r
            break
    assert mobile_row is not None, "Mobile row not found in Annexure 1"
    mob_details = mobile_row.cells[2].text
    print("\nMobile Model Details in Annexure 1:\n" + mob_details)
    assert "IMEI 1: 865432098765432" in mob_details, "IMEI 1 missing from Mobile details!"
    assert "IMEI 2: 865432098765431" in mob_details, "IMEI 2 missing from Mobile details!"

    # Check Model Details of Cloud (MUST HAVE Email ID & Cloud ID, NO physical color)
    cloud_row = None
    for r in t_a1.rows:
        if "aws-admin@dalalalgo.com" in r.cells[2].text:
            cloud_row = r
            break
    assert cloud_row is not None, "Cloud row not found in Annexure 1"
    cloud_details = cloud_row.cells[2].text
    print("\nCloud Model Details in Annexure 1:\n" + cloud_details)
    assert "Email ID: aws-admin@dalalalgo.com" in cloud_details
    assert "Cloud ID:" not in cloud_details, "Cloud ID should not appear for Cloud!"
    assert "IP:" not in cloud_details, "IP should not appear for Cloud!"
    assert "Domain:" not in cloud_details, "Domain should not appear for Cloud!"
    assert "Color:" not in cloud_details, "Color should not appear for Cloud!"

    print("\n=======================================================")
    print("VERIFYING RULE 2: Cloud & File/Folder Checkboxes & Notes")
    print("=======================================================")
    # Generate Cloud docs
    res_cloud = doc_generator.generate_evidence_docs(cloud_id)
    cloud_ctx = doc_generator._build_evidence_context(excel_db.get_evidence(cloud_id))
    print("Cloud in Non-Mobile Form:")
    print(f"  is_other_sym: {cloud_ctx['is_other_sym']} (Expected: ☑)")
    print(f"  non_mobile_other_specify: {cloud_ctx['non_mobile_other_specify']} (Expected: aws-admin@dalalalgo.com)")
    assert cloud_ctx['is_other_sym'] == '☑', "Non-Mobile form Other checkbox should be ☑ for Cloud"
    assert "aws-admin@dalalalgo.com" in cloud_ctx['non_mobile_other_specify']
    print("Cloud in Annexure E / BSA Certificate:")
    print(f"  cb_cloud: {cloud_ctx['cb_cloud']} (Expected: ☑)")
    print(f"  cb_other: {cloud_ctx['cb_other']} (Expected: ☐)")
    print(f"  other_specific_notes: {cloud_ctx['other_specific_notes']}")
    assert cloud_ctx['cb_cloud'] == '☑', "Annexure E Cloud checkbox MUST be ☑ for Cloud!"
    assert cloud_ctx['cb_other'] == '☐', "Annexure E Other checkbox MUST NOT be ☑ for Cloud!"
    assert "tick-by-tick" in cloud_ctx['other_specific_notes']

    # Generate File/Folder docs
    res_file = doc_generator.generate_evidence_docs(file_id)
    file_ctx = doc_generator._build_evidence_context(excel_db.get_evidence(file_id))
    print("\nFile/Folder in Non-Mobile Form:")
    print(f"  is_filefolder_sym: {file_ctx['is_filefolder_sym']} (Expected: ☑)")
    print(f"  is_other_sym: {file_ctx['is_other_sym']} (Expected: ☐)")
    assert file_ctx['is_filefolder_sym'] == '☑', "Non-Mobile form FILE/FOLDER checkbox MUST be ☑ for File/Folder!"
    assert file_ctx['is_other_sym'] == '☐', "Non-Mobile form OTHER checkbox MUST be ☐ for File/Folder!"
    print("File/Folder in Annexure E / BSA Certificate:")
    print(f"  cb_other: {file_ctx['cb_other']} (Expected: ☑)")
    print(f"  cb_cloud: {file_ctx['cb_cloud']} (Expected: ☐)")
    print(f"  annexure_e_other_specify: {file_ctx['annexure_e_other_specify']} (Expected: File/Folder)")
    assert file_ctx['cb_other'] == '☑', "Annexure E Other checkbox MUST be ☑ for File/Folder!"
    assert file_ctx['cb_cloud'] == '☐', "Annexure E Cloud checkbox MUST be ☐ for File/Folder!"
    assert file_ctx['annexure_e_other_specify'] == 'File/Folder', "Annexure E Other specify MUST be 'File/Folder'!"

    print("\n=======================================================")
    print("VERIFYING RULE 3: COC Table 1 Evidence Names & Formatting")
    print("=======================================================")
    a3_path = doc_generator.generate_mumbai_annexure_3_coc(premise_id)
    assert os.path.exists(a3_path), "Annexure 3 COC docx not found"
    doc_a3 = docx.Document(a3_path)
    t1_coc = doc_a3.tables[1]
    assert len(t1_coc.columns) == 6, f"Expected 6 columns in COC Table 1, got {len(t1_coc.columns)}"
    
    t1_text = t1_coc.rows[2].cells[0].text
    t2_text = t1_coc.rows[3].cells[0].text
    print("\nTransfer 1 (After Imaging):\n" + t1_text[:400])
    
    # Check that Unique Master and Working copies are listed in Transfer 1 and Transfer 2
    assert "Master copy S/N: WD-MUM-MC-001" in t1_text
    assert "Working copy S/N: SG-MUM-WC-001" in t1_text
    assert "Master copy S/N: WD-MUM-MC-001" in t2_text
    assert "Working copy S/N: SG-MUM-WC-001" in t2_text

    # Transfer 1 has first entry date/time, Transfer 2 is empty for pen
    t1_dt = t1_coc.rows[2].cells[4].text.strip()
    t2_dt = t1_coc.rows[3].cells[4].text.strip()
    print(f"\nTransfer 1 Date & Time: '{t1_dt}'")
    print(f"Transfer 2 Date & Time: '{t2_dt}' (Expected empty string)")
    assert "10:30 AM" in t1_dt, "Transfer 1 should contain premise first entry time"
    assert t2_dt == "", "Transfer 2 Date & Time MUST be empty for pen entry!"

    print("\n=======================================================")
    print("VERIFYING RULE 4: Annexure 1 Remarks Column is Empty")
    print("=======================================================")
    for r_i, r in enumerate(t_a1.rows[1:], 1):
        rem_text = r.cells[6].text.strip()
        # Non-empty rows should have empty remarks
        if r.cells[1].text != '-':
            assert rem_text == "", f"Row {r_i} Remarks should be empty, got: '{rem_text}'"
    print("All Annexure 1 data row Remarks are completely blank as requested!")

    print("\n=======================================================")
    print("VERIFYING RULE 5: Single Device ZIP Excludes Annexure 1")
    print("=======================================================")
    import zipfile
    res_d1 = doc_generator.generate_evidence_docs(desktop_ids[0])
    for f in res_d1['files']:
        assert "Annexure_1" not in os.path.basename(f), f"Annexure 1 found in single-device files: {f}"
    with zipfile.ZipFile(res_d1['zip_path'], 'r') as zf:
        namelist = zf.namelist()
        print(f"Single device ZIP contains: {namelist}")
        assert any("Device_Collection_" in n for n in namelist), "Device collection form missing in zip"
        assert any("BSA_Certificate_" in n for n in namelist), "BSA Certificate missing in zip"
        assert all("Annexure_1" not in n for n in namelist), "Annexure 1 MUST NOT be in single-device ZIP!"

    print("\n=======================================================")
    print("VERIFYING RULE 6: Annexure E Father Name, No Company, Hashes & Checkboxes")
    print("=======================================================")
    # 1. Inspect generated BSA Certificate
    bsa_file = [f for f in res_d1['files'] if 'BSA_Certificate_' in f][0]
    doc_bsa = docx.Document(bsa_file)
    bsa_full_text = "\n".join([p.text for p in doc_bsa.paragraphs])
    
    assert "Son/daughter/spouse of Shri Rameshwar Verma" in bsa_full_text, "Expert Father Name not found in Part B!"
    assert "do hereby solemnly affirm and sincerely state and submit as follows: —" in bsa_full_text
    
    # Check that Part B has residing/employed at R S Patel & Co per user specification
    assert "residing/employed at R S Patel & Co" in bsa_full_text, "Part B must state 'residing/employed at R S Patel & Co'!"
    print("Part B contains 'Son/daughter/spouse of Shri Rameshwar Verma' and 'residing/employed at R S Patel & Co.'!")

    # 2. Check Hashes Sequence (MD5, SHA1, SHA256)
    ctx_d1 = doc_generator._build_evidence_context(excel_db.get_evidence(desktop_ids[0]))
    h_seq = ctx_d1['hash_value']
    print(f"\nHash Sequence in Annexure E:\n{h_seq}")
    assert h_seq.startswith("MD5:"), "Hash sequence must start with MD5!"
    assert ", SHA1:" in h_seq, "Hash sequence must have SHA1 second!"
    assert ", SHA256:" in h_seq, "Hash sequence must have SHA256 third!"
    assert ctx_d1['hash_md5'] == '☑', "MD5 checkbox must be ☑"
    assert ctx_d1['hash_sha1'] == '☑', "SHA1 checkbox must be ☑"
    assert ctx_d1['hash_sha256'] == '☑', "SHA256 checkbox must be ☑"

    # 3. Check partial hashes (e.g. only MD5 & SHA256, no SHA1)
    ev_partial = {
        'evidence_id': 'TEST_HASH',
        'evidence_category': 'Laptop',
        'md5_hash': 'd41d8cd98f00b204e9800998ecf8427e',
        'sha1_hash': '',
        'sha256_hash_1': 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
    }
    ctx_partial = doc_generator._build_evidence_context(ev_partial)
    print(f"\nPartial Hashes (MD5 & SHA256 only): {ctx_partial['hash_value']}")
    assert ctx_partial['hash_value'] == "MD5: d41d8cd98f00b204e9800998ecf8427e, SHA256: e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    assert ctx_partial['hash_md5'] == '☑'
    assert ctx_partial['hash_sha1'] == '☐'
    assert ctx_partial['hash_sha256'] == '☑'
    print("Whichever hash is provided has its checkbox checked ☑ and sequence is MD5, SHA256!")

    print("\n=======================================================")
    print("ALL DETAILED FUNCTIONAL RULES VERIFIED WITH 100% SUCCESS!")
    print("=======================================================")

if __name__ == '__main__':
    test_detailed_rules()
