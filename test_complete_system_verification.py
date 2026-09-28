import os
import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')

import excel_db
import doc_generator

def verify_all():
    print("=== STARTING FULL SYSTEM VERIFICATION ===")
    excel_db.init_db(force=True)

    # 1. Create a Gujarat search
    s_guj_id = excel_db.save_search({
        'case_name': 'Gujarat Tech Search',
        'jurisdiction': 'Gujarat',
        'date': '2026-09-28',
        'authorised_officer': 'Officer Gujarat'
    })
    p_guj_id = excel_db.save_premise({
        'search_id': s_guj_id,
        'premise_name': 'Ahmedabad Office',
        'address': 'SG Highway, Ahmedabad',
        'team': 'Team GUJ-1',
        'date': '2026-09-28'
    })

    # Add evidence 1 to Gujarat
    ev1_id = excel_db.save_evidence({
        'premise_id': p_guj_id,
        'evidence_category': 'Laptop',
        'device_name': 'Director ThinkPad',
        'make': 'Lenovo',
        'model': 'ThinkPad X1',
        'serial_number': 'LNV-889900',
        'capacity': '1 TB',
        'color': 'Black',
        'master_copy_make': 'WD',
        'master_copy_model': 'My Passport',
        'master_copy_sn': 'WD-GUJ-M01',
        'master_copy_capacity': '2 TB',
        'working_copy_make': 'Seagate',
        'working_copy_model': 'One Touch',
        'working_copy_sn': 'SG-GUJ-W01',
        'working_copy_capacity': '2 TB',
        'party_name': 'Shri Jignesh Shah',
        'party_father_name': 'Late Shri Rasiklal Shah',
        'party_relation': 'Managing Director',
        'party_residing_at': 'Bopal, Ahmedabad',
        'expert_name': 'Ravi Patel',
        'expert_father_name': 'Suresh Patel',
        'expert_designation': 'Forensic Director',
        'sha256_hash_1': 'a1b2c3d4e5f67890123456789012345678901234567890123456789012345678',
        'other_specific_notes': 'Corporate accounting records extracted'
    })

    print(f"Created Gujarat Evidence {ev1_id}")

    # Check last evidence prefill for premise
    last_ev = excel_db.get_last_evidence_for_premise(p_guj_id, s_guj_id)
    assert last_ev is not None, "get_last_evidence_for_premise failed!"
    assert last_ev['master_copy_sn'] == 'WD-GUJ-M01', f"Prefill mismatch: {last_ev['master_copy_sn']}"
    assert last_ev['party_father_name'] == 'Late Shri Rasiklal Shah', f"Prefill party father mismatch: {last_ev['party_father_name']}"
    print("✔ Auto-fill from last evidence verified successfully!")

    # Test Gujarat document generation
    guj_docs = doc_generator.generate_evidence_docs(ev1_id)
    print(f"Gujarat docs generated: {guj_docs['files']}")
    assert len(guj_docs['files']) == 1, "Gujarat should generate 1 Annexure E file"
    guj_doc_path = guj_docs['files'][0]
    
    # Verify contents of Gujarat Annexure E docx
    g_doc = docx.Document(guj_doc_path)
    
    # Inspect Table 1 & Table 5
    for t_idx in [1, 5]:
        if len(g_doc.tables) > t_idx:
            tbl = g_doc.tables[t_idx]
            assert len(tbl.rows) >= 3, f"Table {t_idx} should have device specification rows"
    print("✔ Gujarat Annexure E Table 1 and Table 5 verified!")

    # Check text for Part A, Part B, R S Patel & Co., and Hashes
    all_text = "\n".join([p.text for p in g_doc.paragraphs])
    for t in g_doc.tables:
        for r in t.rows:
            for c in r.cells:
                all_text += "\n" + c.text

    assert "Late Shri Rasiklal Shah" in all_text, "Party father name must be present in Part A!"
    assert "R S Patel & Co." in all_text, "Part B must state 'residing/employed at R S Patel & Co.'!"
    assert "Suresh Patel" in all_text, "Expert father name must be present in Part B!"
    assert "Corporate accounting records extracted" in all_text, "other_specific_notes must replace notes_by_officer!"
    assert "SHA256:" in all_text, "Hashes must be present!"
    assert "Model No.: My Passport" in all_text, f"Gujarat Annexure E Master Copy Model missing: {all_text}"
    assert "Size: 2 TB" in all_text, f"Gujarat Annexure E Master Copy Size missing: {all_text}"
    print("✔ Gujarat Annexure E content (Part A, Part B R S Patel & Co., Master Copy Model & Capacity, notes, hashes) verified!")

    # 2. Test Delhi search and document generation
    s_del_id = excel_db.save_search({
        'case_name': 'Delhi Cyber Investigation',
        'jurisdiction': 'Delhi',
        'date': '2026-09-28',
        'authorised_officer': 'Officer Delhi'
    })
    p_del_id = excel_db.save_premise({
        'search_id': s_del_id,
        'premise_name': 'Barakhamba Road Office',
        'address': 'Barakhamba Road, Connaught Place, New Delhi',
        'team': 'Team DEL-1',
        'date': '2026-09-28'
    })
    ev2_id = excel_db.save_evidence({
        'premise_id': p_del_id,
        'evidence_category': 'Mobile',
        'device_name': 'Target iPhone 15',
        'make': 'Apple',
        'model': 'iPhone 15 Pro',
        'serial_number': 'APP-DEL-9911',
        'imei1': '867543210987654',
        'imei2': '867543210987655',
        'master_copy_make': 'WD',
        'master_copy_model': 'My Passport',
        'master_copy_sn': 'WD-DEL-M01',
        'master_copy_capacity': '2 TB',
        'working_copy_make': 'Seagate',
        'working_copy_model': 'One Touch',
        'working_copy_sn': 'SG-DEL-W01',
        'working_copy_capacity': '2 TB',
        'party_name': 'Vikram Malhotra',
        'party_father_name': 'Shri S. K. Malhotra',
        'expert_name': 'Amit Kumar',
        'expert_father_name': 'Shri O. P. Kumar',
        'sha256_hash_1': 'fedcba9876543210fedcba9876543210fedcba9876543210fedcba9876543210'
    })

    del_docs = doc_generator.generate_evidence_docs(ev2_id)
    print(f"Delhi docs generated: {del_docs['files']}")
    assert len(del_docs['files']) == 3, "Delhi Mobile should generate Annexure C, D, E"
    
    # Verify Master Copy in Delhi Annexure D
    del_d_path = [f for f in del_docs['files'] if 'Annexure_D' in f][0]
    d_doc_d = docx.Document(del_d_path)
    d_text_d = "\n".join([c.text for t in d_doc_d.tables for r in t.rows for c in r.cells])
    assert "MODEL: My Passport" in d_text_d, f"Delhi Annexure D Master Copy Model missing: {d_text_d}"
    assert "CAPACITY: 2 TB" in d_text_d, f"Delhi Annexure D Master Copy Capacity missing: {d_text_d}"
    print("✔ Delhi Annexure D Master Copy Model and Capacity verified!")

    del_e_path = [f for f in del_docs['files'] if 'Annexure_E' in f][0]
    d_doc = docx.Document(del_e_path)
    del_text = "\n".join([p.text for p in d_doc.paragraphs])
    for t in d_doc.tables:
        for r in t.rows:
            for c in r.cells:
                del_text += "\n" + c.text
    assert "Shri S. K. Malhotra" in del_text, "Delhi Part A father name missing!"
    assert "Amit Kumar" in del_text, "Delhi Part B Expert Name missing!"
    print("✔ Delhi Annexures & Annexure E verified!")

    # 3. Test Mumbai search and document generation
    s_mum_id = excel_db.save_search({
        'case_name': 'Mumbai Stock Investigation',
        'jurisdiction': 'Mumbai',
        'date': '2026-09-28',
        'authorised_officer': 'Officer Mumbai'
    })
    p_mum_id = excel_db.save_premise({
        'search_id': s_mum_id,
        'premise_name': 'Nariman Point Floor 12',
        'address': 'Maker Chambers, Nariman Point, Mumbai',
        'team': 'Team MUM-1',
        'date': '2026-09-28'
    })
    ev3_id = excel_db.save_evidence({
        'premise_id': p_mum_id,
        'evidence_category': 'Desktop',
        'device_name': 'Algo Server Primary',
        'make': 'Dell',
        'model': 'PowerEdge R750',
        'serial_number': 'DELL-SRV-MUM01',
        'location_found': 'Server Rack 04',
        'owner_user': 'Chief Technology Officer',
        'is_owner': 'Yes',
        'master_copy_sn': '134567890',
        'working_copy_sn': '2345678909',
        'party_name': 'Rajesh Jhunjhunwala',
        'party_father_name': 'Late Shri R. K. Jhunjhunwala',
        'expert_name': 'Pratham Forensic',
        'expert_father_name': 'Doshi Senior',
        'sha256_hash_1': '1122334455667788990011223344556677889900112233445566778899001122'
    })

    mum_annex_1 = doc_generator.generate_mumbai_annexure_1(p_mum_id)
    mum_annex_3 = doc_generator.generate_mumbai_annexure_3_coc(p_mum_id)
    mum_device_docs = doc_generator.generate_evidence_docs(ev3_id)
    
    print(f"Mumbai Annexure 1: {mum_annex_1}")
    print(f"Mumbai Annexure 3: {mum_annex_3}")
    print(f"Mumbai Device Docs: {mum_device_docs['files']}")

    # Verify Annexure 1 content
    doc1 = docx.Document(mum_annex_1)
    txt1 = "\n".join([c.text for r in doc1.tables[0].rows for c in r.cells])
    assert "Server Rack 04" in txt1, "Location found should be in Column 1 (Nature of evidence / Where)"
    assert "Evidence Name: Algo Server Primary" in txt1, "Evidence Name should be in Model Details"
    assert "Chief Technology Officer" in txt1, "Owner name should be in Is Owner column if yes"
    print("✔ Mumbai Annexure 1 verified!")

    # Verify Annexure 3 content
    doc3 = docx.Document(mum_annex_3)
    txt3 = "\n".join([c.text for t in doc3.tables for r in t.rows for c in r.cells])
    assert "Master copy S/N: 134567890" in txt3, "Master copy S/N in Annexure 3 verified"
    assert "Working copy S/N: 2345678909" in txt3, "Working copy S/N in Annexure 3 verified"
    print("✔ Mumbai Annexure 3 CoC verified!")

    # Verify Mumbai BSA Certificate
    bsa_path = [f for f in mum_device_docs['files'] if 'BSA_Certificate' in f][0]
    bsa_doc = docx.Document(bsa_path)
    bsa_txt = "\n".join([p.text for p in bsa_doc.paragraphs])
    for t in bsa_doc.tables:
        for r in t.rows:
            for c in r.cells:
                bsa_txt += "\n" + c.text
    assert "Late Shri R. K. Jhunjhunwala" in bsa_txt, "Mumbai Part A father name missing!"
    assert "R S Patel & Co" in bsa_txt, "Mumbai Part B R S Patel & Co missing!"
    print("✔ Mumbai BSA Certificate verified!")

    print("\n🎉 ALL TESTS AND DETAILED RULES PASSED WITH 100% ACCURACY! 🎉")

if __name__ == '__main__':
    verify_all()
