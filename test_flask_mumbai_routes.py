import app
import excel_db

app.app.config['TESTING'] = True
client = app.app.test_client()

# Look up Mumbai premise and evidence dynamically
searches = excel_db.get_searches()
mum_s = next(s for s in searches if s.get('jurisdiction') == 'Mumbai')
mum_p = excel_db.get_premises(mum_s['search_id'])[0]
mum_evs = excel_db.get_evidences(mum_p['premise_id'])
ev_id = mum_evs[0]['evidence_id'] if mum_evs else 'E001'

# Test new evidence page for Mumbai
resp = client.get(f"/premise/{mum_p['premise_id']}/evidence/new")
assert resp.status_code == 200
html = resp.data.decode('utf-8')

assert 'party_father_name' in html, 'party_father_name field missing from Mumbai evidence form'
assert 'expert_father_name' in html, 'expert_father_name field missing from Mumbai evidence form'
assert 'Device Return & Verification Status' not in html, 'Device Return & Verification Status must be removed from Mumbai form'
assert 'device_control_opts' not in html, 'device_control_opts must be removed from Mumbai form'
assert 'Chain of Custody Logs' not in html, 'Chain of Custody Logs should be hidden for Mumbai'
assert 'id="custody-section"' not in html, 'custody-section should be hidden for Mumbai'
assert 'type="hidden" id="collection_time"' in html, 'collection_time should be hidden input for Mumbai'
print('Evidence Form UI verification PASSED!')

# Test evidence generate route
resp_gen = client.get(f'/generate/evidence/{ev_id}')
assert resp_gen.status_code == 200
print('Single device ZIP route PASSED!')

# Test premise annexure 1 route
resp_a1 = client.get(f"/generate/premise/{mum_p['premise_id']}/mumbai-annexure-1")
assert resp_a1.status_code == 200
print('Premise Mumbai Annexure 1 download PASSED!')

# Test premise mumbai annexure 3 route
resp_a3 = client.get(f"/generate/premise/{mum_p['premise_id']}/mumbai-annexure-3")
assert resp_a3.status_code == 200
print('Premise Mumbai Annexure 3 download PASSED!')
print('\nALL FLASK ENDPOINTS VERIFIED SUCCESSFULLY!')
