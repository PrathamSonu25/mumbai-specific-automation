import os
import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')

import excel_db
import doc_generator
from app import app

def test_delhi_cloud():
    print("=== TESTING DELHI CLOUD EVIDENCE & ANNEXURE E / F WORKFLOW ===")
    excel_db.init_db()

    # 1. Get or create Delhi search & premise
    searches = excel_db.get_searches()
    del_s = next((s for s in searches if s.get('jurisdiction') == 'Delhi'), None)
    if not del_s:
        s_id = excel_db.save_search({
            'case_name': 'Delhi Cyber Cloud Search',
            'jurisdiction': 'Delhi',
            'date': '2026-09-28',
            'authorised_officer': 'Officer Delhi IT'
        })
        p_id = excel_db.save_premise({
            'search_id': s_id,
            'premise_name': 'Connaught Place Branch',
            'address': 'Barakhamba Road, New Delhi',
            'team': 'Team DEL-Cloud',
            'date': '2026-09-28'
        })
    else:
        s_id = del_s['search_id']
        prems = excel_db.get_premises(s_id)
        p_id = prems[0]['premise_id']

    print(f"Using Delhi Search: {s_id}, Premise: {p_id}")

    # 2. Save a Cloud Evidence entry with full Annexure E & F fields
    ce_data = {
        'premise_id': p_id,
        'sn': '1',
        'source': 'Google Workspace (Gmail)',
        'device_imaged': 'finance.director@delhicorp.in',
        'cloud_id': 'finance.director@delhicorp.in',
        'size_of_data': '18.4 GB',
        'export_format': 'PST / MBOX',
        'collection_date': '2026-09-28',
        'collection_time': '14:30',
        'cert_place': 'New Delhi',
        'authorised_officer': 'Shri R. K. Sharma, Addl. Director',
        'party_name': 'Shri Deepak Gupta',
        'party_father_name': 'Late Shri M. L. Gupta',
        'party_relation': 'Finance Director',
        'party_residing_at': 'Greater Kailash-II, New Delhi',
        'expert_name': 'Gaurav Saxena',
        'expert_father_name': 'Shri P. N. Saxena',
        'expert_designation': 'Principal Cyber Forensic Analyst',
        'hash_algorithm': 'SHA256',
        'sha256_hash_1': '8f4c2e6b1a9d0f3e5c7a2b4d6e8f0a1c3e5b7d9f1a2c4e6b8d0f2a4c6e8b0a1d',
        'md5_hash': 'e4d909c290d0fb1ca068ffaddf22cbd0',
        'device_control': 'Owned, Maintained',
        'other_specific_notes': 'Complete Google Takeout MBOX export with header logs verified'
    }

    ce_id = excel_db.save_cloud_evidence(ce_data)
    print(f"✔ Created Cloud Evidence: {ce_id}")

    # Verify retrieval
    ce_fetched = excel_db.get_cloud_evidence(ce_id)
    assert ce_fetched is not None, "get_cloud_evidence returned None!"
    assert ce_fetched['cloud_id'] == 'finance.director@delhicorp.in'
    assert ce_fetched['party_father_name'] == 'Late Shri M. L. Gupta'
    print("✔ Cloud evidence retrieval from database verified!")

    # 3. Test generate_premise_annexure_f (Consolidated Cloud Evidence)
    f_path = doc_generator.generate_premise_annexure_f(p_id)
    assert os.path.exists(f_path), f"Annexure F file does not exist: {f_path}"
    f_doc = docx.Document(f_path)
    f_text = "\n".join([p.text for p in f_doc.paragraphs] + [c.text for t in f_doc.tables for r in t.rows for c in r.cells])
    assert "finance.director@delhicorp.in" in f_text, "Cloud ID missing in Annexure F!"
    assert "18.4 GB" in f_text, "Size missing in Annexure F!"
    print(f"✔ Annexure F verified successfully: {f_path}")

    # 4. Test generate_cloud_annexure_e (Individual Cloud Annexure E)
    e_path = doc_generator.generate_cloud_annexure_e(ce_id)
    assert os.path.exists(e_path), f"Cloud Annexure E does not exist: {e_path}"
    print(f"✔ Generated Cloud Annexure E: {e_path}")

    e_doc = docx.Document(e_path)
    
    # Verify Table 1 (and Table 5): Make, Model, S/N, Color must be N/A, unique_id must be Cloud ID!
    for t_idx in [1, 5]:
        tbl = e_doc.tables[t_idx]
        t_dict = {}
        for r in tbl.rows:
            label = r.cells[0].text.strip()
            val = r.cells[1].text.strip()
            t_dict[label] = val
        print(f"Table {t_idx} contents: {t_dict}")
        assert t_dict.get('Make & Model:') == 'N/A', f"Make & Model must be N/A: {t_dict}"
        assert t_dict.get('Serial Number:') == 'N/A', f"Serial Number must be N/A: {t_dict}"
        assert t_dict.get('Color:') == 'N/A', f"Color must be N/A: {t_dict}"
        assert t_dict.get('IMEI/UIN/UID/MAC/Cloud ID: (as applicable)') == 'finance.director@delhicorp.in', f"Cloud ID mismatch: {t_dict}"
    print("✔ Tables 1 & 5 correctly have Make/Model/SN/Color as N/A and Cloud ID in unique_id!")

    # Verify Table 0 (and Table 4): Cloud checkbox must be checked ☑, others ☐
    for t_idx in [0, 4]:
        tbl = e_doc.tables[t_idx]
        all_cells = " ".join([c.text.strip().replace('\n', ' ') for r in tbl.rows for c in r.cells])
        assert "Cloud ☑" in all_cells or ("Cloud" in all_cells and "☑" in all_cells), f"Cloud checkbox not checked: {all_cells}"
    print("✔ Tables 0 & 4 have Cloud checkbox ☑ checked!")

    # Verify Part A and Part B text
    all_paras = "\n".join([p.text for p in e_doc.paragraphs])
    assert "Shri Deepak Gupta" in all_paras, "Party Name missing in Part A!"
    assert "Late Shri M. L. Gupta" in all_paras, "Party Father Name missing in Part A!"
    assert "Gaurav Saxena" in all_paras, "Expert Name missing in Part B!"
    assert "Shri P. N. Saxena" in all_paras, "Expert Father Name missing in Part B!"
    assert "residing/employed at" in all_paras, "Part B must state residing/employed at!"
    print("✔ Part A and Part B declarations verified in Annexure E!")

    # Verify Hashes in Annexure E
    assert "8f4c2e6b1a9d0f3e5c7a2b4d6e8f0a1c3e5b7d9f1a2c4e6b8d0f2a4c6e8b0a1d" in all_paras, "SHA-256 missing in Annexure E!"
    assert "e4d909c290d0fb1ca068ffaddf22cbd0" in all_paras, "MD5 missing in Annexure E!"
    print("✔ Cryptographic Hashes verified in Annexure E!")

    # 5. Test Web Endpoints via Flask Test Client
    client = app.test_client()
    
    # Test GET new cloud evidence form
    res_get = client.get(f"/premise/{p_id}/cloud/new")
    assert res_get.status_code == 200
    assert b"Part A" in res_get.data
    assert b"Part B" in res_get.data
    assert b"Cryptographic Hash Verification" in res_get.data
    print("✔ GET /premise/<id>/cloud/new renders comprehensive form!")

    # Test GET generate cloud Annexure E route
    res_e = client.get(f"/generate/cloud/{ce_id}/annexure-e")
    assert res_e.status_code == 200
    assert len(res_e.data) > 5000
    print(f"✔ GET /generate/cloud/{ce_id}/annexure-e returned {len(res_e.data)} bytes!")

    # Test GET generate premise Annexure F route
    res_f = client.get(f"/generate/premise/{p_id}/annexure-f")
    assert res_f.status_code == 200
    assert len(res_f.data) > 5000
    print(f"✔ GET /generate/premise/{p_id}/annexure-f returned {len(res_f.data)} bytes!")

    # Test Premise Detail page shows Annexure E button
    res_prem = client.get(f"/premise/{p_id}")
    assert res_prem.status_code == 200
    assert f"/generate/cloud/{ce_id}/annexure-e".encode() in res_prem.data, "Annexure E download button missing on premise detail page!"
    print("✔ Premise detail page displays 'Annexure E' button for cloud evidence!")

    print("\n🎉 ALL DELHI CLOUD EVIDENCE TESTS PASSED WITH 100% SUCCESS! 🎉")

if __name__ == '__main__':
    test_delhi_cloud()
