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

    # 2. Search detail (S004 - Mumbai)
    res = client.get('/search/S004')
    print(f"GET /search/S004 -> Status {res.status_code}")
    assert res.status_code == 200, "Search detail failed"
    assert b'Mumbai' in res.data, "Mumbai badge missing from search detail"

    # 3. Premise detail (P004 - Mumbai)
    res = client.get('/premise/P004')
    print(f"GET /premise/P004 -> Status {res.status_code}")
    assert res.status_code == 200, "Premise detail failed"
    assert b'Generate Inventory (Annexure 1)' in res.data, "Annexure 1 button missing"
    assert b'Generate COC (Annexure 3)' in res.data, "Annexure 3 COC button missing"

    # 4. Generate Premise Mumbai Annexure 1
    res = client.get('/generate/premise/P004/mumbai-annexure-1')
    print(f"GET /generate/premise/P004/mumbai-annexure-1 -> Status {res.status_code}, Length: {len(res.data)}")
    assert res.status_code == 200, "Download Annexure 1 failed"
    assert len(res.data) > 1000, "Annexure 1 content too small"

    # 5. Generate Premise Mumbai Annexure 3 COC
    res = client.get('/generate/premise/P004/mumbai-annexure-3')
    print(f"GET /generate/premise/P004/mumbai-annexure-3 -> Status {res.status_code}, Length: {len(res.data)}")
    assert res.status_code == 200, "Download Annexure 3 COC failed"
    assert len(res.data) > 1000, "Annexure 3 COC content too small"

    # 6. Generate Device Docs for E006 (Laptop)
    res = client.get('/generate/evidence/E006')
    print(f"GET /generate/evidence/E006 -> Status {res.status_code}, Length: {len(res.data)}")
    assert res.status_code == 200, "Download Laptop docs failed"
    assert 'zip' in res.headers.get('Content-Type', '').lower() or 'zip' in res.headers.get('Content-Disposition', '').lower()

    # 7. Generate Device Docs for E007 (Mobile)
    res = client.get('/generate/evidence/E007')
    print(f"GET /generate/evidence/E007 -> Status {res.status_code}, Length: {len(res.data)}")
    assert res.status_code == 200, "Download Mobile docs failed"
    assert 'zip' in res.headers.get('Content-Type', '').lower() or 'zip' in res.headers.get('Content-Disposition', '').lower()

    # 8. Check Delhi still works properly (S002 / P002 / E002)
    res_delhi = client.get('/premise/P002')
    print(f"GET /premise/P002 (Delhi) -> Status {res_delhi.status_code}")
    assert res_delhi.status_code == 200
    assert b'Generate Inventory (Annexure G)' in res_delhi.data

    # 9. Check Gujarat still works properly (S001)
    res_gujarat = client.get('/search/S001')
    print(f"GET /search/S001 (Gujarat) -> Status {res_gujarat.status_code}")
    assert res_gujarat.status_code == 200
    assert b'Generate Annexure E (All Evidences)' in res_gujarat.data

    print("\nALL FLASK ROUTE TESTS PASSED SUCCESSFULLY!")

if __name__ == '__main__':
    test_routes()
