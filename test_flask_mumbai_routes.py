import app

app.app.config['TESTING'] = True
client = app.app.test_client()

# Test new evidence page for Mumbai
resp = client.get('/premise/P011/evidence/new')
assert resp.status_code == 200
html = resp.data.decode('utf-8')

assert 'expert_father_name' in html, 'expert_father_name field missing from evidence form'
assert 'Chain of Custody Logs' not in html, 'Chain of Custody Logs should be hidden for Mumbai'
assert 'id="custody-section"' not in html, 'custody-section should be hidden for Mumbai'
assert 'type="hidden" id="collection_time"' in html, 'collection_time should be hidden input for Mumbai'
print('Evidence Form UI verification PASSED!')

# Test evidence generate route
resp_gen = client.get('/generate/evidence/E032')
assert resp_gen.status_code == 200
print('Single device ZIP route PASSED!')

# Test premise annexure 1 route
resp_a1 = client.get('/generate/premise/P011/mumbai-annexure-1')
assert resp_a1.status_code == 200
print('Premise Mumbai Annexure 1 download PASSED!')

# Test premise mumbai annexure 3 route
resp_a3 = client.get('/generate/premise/P011/mumbai-annexure-3')
assert resp_a3.status_code == 200
print('Premise Mumbai Annexure 3 download PASSED!')
print('\nALL FLASK ENDPOINTS VERIFIED SUCCESSFULLY!')
