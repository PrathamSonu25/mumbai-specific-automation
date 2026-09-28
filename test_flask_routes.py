import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

from app import app
import excel_db

def test_routes():
    print("Testing Flask Application Routes...")
    client = app.test_client()

    # 1. Index page
    res = client.get('/')
    print(f"GET / -> Status {res.status_code}")
    assert res.status_code == 200, "Index page failed"

    # Find searches by jurisdiction
    searches = excel_db.get_searches()
    mum_s = next(s for s in searches if s.get('jurisdiction') == 'Mumbai')
    del_s = next(s for s in searches if s.get('jurisdiction') == 'Delhi')
    guj_s = next(s for s in searches if s.get('jurisdiction') == 'Gujarat')

    mum_p = excel_db.get_premises(mum_s['search_id'])[0]
    del_p = excel_db.get_premises(del_s['search_id'])[0]
    guj_p = excel_db.get_premises(guj_s['search_id'])[0]

    mum_evs = excel_db.get_evidences(mum_p['premise_id'])
    
    # 2. Search detail (Mumbai)
    res = client.get(f"/search/{mum_s['search_id']}")
    print(f"GET /search/{mum_s['search_id']} -> Status {res.status_code}")
    assert res.status_code == 200, "Search detail failed"
    assert b'Mumbai' in res.data, "Mumbai badge missing from search detail"

    # 3. Premise detail (Mumbai)
    res = client.get(f"/premise/{mum_p['premise_id']}")
    print(f"GET /premise/{mum_p['premise_id']} -> Status {res.status_code}")
    assert res.status_code == 200, "Premise detail failed"
    assert b'Generate Inventory (Annexure 1)' in res.data, "Annexure 1 button missing"
    assert b'Generate COC (Annexure 3)' in res.data, "Annexure 3 COC button missing"

    # 4. Generate Premise Mumbai Annexure 1
    res = client.get(f"/generate/premise/{mum_p['premise_id']}/mumbai-annexure-1")
    print(f"GET /generate/premise/{mum_p['premise_id']}/mumbai-annexure-1 -> Status {res.status_code}, Length: {len(res.data)}")
    assert res.status_code == 200, "Download Annexure 1 failed"
    assert len(res.data) > 1000, "Annexure 1 content too small"

    # 5. Generate Premise Mumbai Annexure 3 COC
    res = client.get(f"/generate/premise/{mum_p['premise_id']}/mumbai-annexure-3")
    print(f"GET /generate/premise/{mum_p['premise_id']}/mumbai-annexure-3 -> Status {res.status_code}, Length: {len(res.data)}")
    assert res.status_code == 200, "Download Annexure 3 COC failed"
    assert len(res.data) > 1000, "Annexure 3 COC content too small"

    # 6. Generate Device Docs for Mumbai Evidence
    if mum_evs:
        ev_id = mum_evs[0]['evidence_id']
        res = client.get(f"/generate/evidence/{ev_id}")
        print(f"GET /generate/evidence/{ev_id} -> Status {res.status_code}, Length: {len(res.data)}")
        assert res.status_code == 200, "Download device docs failed"
        assert 'zip' in res.headers.get('Content-Type', '').lower() or 'zip' in res.headers.get('Content-Disposition', '').lower()

    # 7. Check Delhi still works properly
    res_delhi = client.get(f"/premise/{del_p['premise_id']}")
    print(f"GET /premise/{del_p['premise_id']} (Delhi) -> Status {res_delhi.status_code}")
    assert res_delhi.status_code == 200
    assert b'Generate Inventory (Annexure G)' in res_delhi.data

    # 8. Check Gujarat still works properly
    res_gujarat = client.get(f"/search/{guj_s['search_id']}")
    print(f"GET /search/{guj_s['search_id']} (Gujarat) -> Status {res_gujarat.status_code}")
    assert res_gujarat.status_code == 200
    assert b'Generate Annexure E (All Evidences)' in res_gujarat.data

    # 9. Test New Evidence form prefill route (GET)
    res_new_ev = client.get(f"/premise/{mum_p['premise_id']}/evidence/new")
    print(f"GET /premise/{mum_p['premise_id']}/evidence/new -> Status {res_new_ev.status_code}")
    assert res_new_ev.status_code == 200
    assert b'134567890' in res_new_ev.data, "Master copy SN must be prefilled from last evidence!"
    assert b'2345678909' in res_new_ev.data, "Working copy SN must be prefilled from last evidence!"
    assert b'Late Shri R. K. Jhunjhunwala' in res_new_ev.data, "Party father name must be prefilled from last evidence!"
    print("✔ New evidence prefill verified in HTML form!")

    print("\nALL FLASK ROUTE TESTS PASSED SUCCESSFULLY!")

if __name__ == '__main__':
    test_routes()
