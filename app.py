import os
# pyrefly: ignore [missing-import]
from flask import (
    Flask, render_template, request, redirect, url_for, flash, send_file
)

import config
import excel_db
import doc_generator

app = Flask(__name__)
app.secret_key = config.SECRET_KEY

# Ensure DB is initialized
excel_db.init_db()

@app.route('/')
def index():
    searches = excel_db.get_searches()
    for s in searches:
        premises = excel_db.get_premises(s['search_id'])
        s['premise_count'] = len(premises)
    return render_template('index.html', searches=searches)

@app.route('/search/new', methods=['GET', 'POST'])
def new_search():
    if request.method == 'POST':
        case_name = request.form.get('case_name', '').strip()
        jurisdiction = request.form.get('jurisdiction', 'Gujarat')
        date = request.form.get('date', '').strip()
        officer = request.form.get('authorised_officer', '').strip()
        remarks = request.form.get('remarks', '').strip()

        search_warrant_details = request.form.get('search_warrant_details', '').strip()

        if not case_name or not date:
            flash('Case Name and Date are required.', 'danger')
            return redirect(url_for('new_search'))

        search_id = excel_db.save_search({
            'case_name': case_name,
            'jurisdiction': jurisdiction,
            'date': date,
            'authorised_officer': officer,
            'remarks': remarks,
            'search_warrant_details': search_warrant_details
        })
        flash(f'Search {search_id} created successfully.', 'success')
        return redirect(url_for('search_detail', search_id=search_id))

    return render_template('search_form.html', search=None, jurisdictions=config.JURISDICTIONS)

@app.route('/search/<search_id>')
def search_detail(search_id):
    search = excel_db.get_search(search_id)
    if not search:
        flash('Search not found.', 'danger')
        return redirect(url_for('index'))

    premises = excel_db.get_premises(search_id)
    for p in premises:
        evs = excel_db.get_evidences(p['premise_id'])
        clouds = excel_db.get_cloud_evidences(p['premise_id'])
        p['evidence_count'] = len(evs)
        p['cloud_count'] = len(clouds)

    return render_template('search_detail.html', search=search, premises=premises)

@app.route('/search/<search_id>/edit', methods=['GET', 'POST'])
def edit_search(search_id):
    search = excel_db.get_search(search_id)
    if not search:
        flash('Search not found.', 'danger')
        return redirect(url_for('index'))

    if request.method == 'POST':
        search['case_name'] = request.form.get('case_name', '').strip()
        search['jurisdiction'] = request.form.get('jurisdiction', 'Gujarat')
        search['date'] = request.form.get('date', '').strip()
        search['authorised_officer'] = request.form.get('authorised_officer', '').strip()
        search['remarks'] = request.form.get('remarks', '').strip()
        search['search_warrant_details'] = request.form.get('search_warrant_details', '').strip()

        excel_db.save_search(search)
        flash('Search updated successfully.', 'success')
        return redirect(url_for('search_detail', search_id=search_id))

    return render_template('search_form.html', search=search, jurisdictions=config.JURISDICTIONS)

@app.route('/search/<search_id>/delete', methods=['POST'])
def delete_search(search_id):
    excel_db.delete_search_cascade(search_id)
    flash(f'Search {search_id} and all related premises/evidences deleted.', 'info')
    return redirect(url_for('index'))

# ----------------- PREMISE ROUTES -----------------
@app.route('/search/<search_id>/premise/new', methods=['GET', 'POST'])
def new_premise(search_id):
    search = excel_db.get_search(search_id)
    if not search:
        flash('Search not found.', 'danger')
        return redirect(url_for('index'))

    if request.method == 'POST':
        premise_name = request.form.get('premise_name', '').strip()
        address = request.form.get('address', '').strip()
        team = request.form.get('team', '').strip()
        date = request.form.get('date', '').strip()
        time = request.form.get('time', '').strip()

        if not premise_name or not address:
            flash('Premise Name and Address are required.', 'danger')
            return redirect(url_for('new_premise', search_id=search_id))

        premise_id = excel_db.save_premise({
            'search_id': search_id,
            'premise_name': premise_name,
            'address': address,
            'team': team,
            'date': date,
            'time': time
        })
        flash(f'Premise {premise_id} created successfully.', 'success')
        return redirect(url_for('premise_detail', premise_id=premise_id))

    return render_template('premise_form.html', search=search, premise=None)

@app.route('/premise/<premise_id>')
def premise_detail(premise_id):
    premise = excel_db.get_premise(premise_id)
    if not premise:
        flash('Premise not found.', 'danger')
        return redirect(url_for('index'))

    search = excel_db.get_search(premise['search_id'])
    evidences = excel_db.get_evidences(premise_id)
    for ev in evidences:
        sub_evs = excel_db.get_sub_evidences(ev['evidence_id'])
        custodies = excel_db.get_custody_rows(ev['evidence_id'])
        ev['sub_count'] = len(sub_evs)
        ev['custody_count'] = len(custodies)

    cloud_evs = excel_db.get_cloud_evidences(premise_id)

    return render_template(
        'premise_detail.html',
        search=search,
        premise=premise,
        evidences=evidences,
        cloud_evidences=cloud_evs
    )

@app.route('/premise/<premise_id>/edit', methods=['GET', 'POST'])
def edit_premise(premise_id):
    premise = excel_db.get_premise(premise_id)
    if not premise:
        flash('Premise not found.', 'danger')
        return redirect(url_for('index'))

    search = excel_db.get_search(premise['search_id'])

    if request.method == 'POST':
        premise['premise_name'] = request.form.get('premise_name', '').strip()
        premise['address'] = request.form.get('address', '').strip()
        premise['team'] = request.form.get('team', '').strip()
        premise['date'] = request.form.get('date', '').strip()
        premise['time'] = request.form.get('time', '').strip()

        excel_db.save_premise(premise)
        flash('Premise updated successfully.', 'success')
        return redirect(url_for('premise_detail', premise_id=premise_id))

    return render_template('premise_form.html', search=search, premise=premise)

@app.route('/premise/<premise_id>/delete', methods=['POST'])
def delete_premise(premise_id):
    premise = excel_db.get_premise(premise_id)
    search_id = premise.get('search_id') if premise else None
    excel_db.delete_premise_cascade(premise_id)
    flash(f'Premise {premise_id} deleted.', 'info')
    if search_id:
        return redirect(url_for('search_detail', search_id=search_id))
    return redirect(url_for('index'))

# ----------------- EVIDENCE ROUTES -----------------
@app.route('/premise/<premise_id>/evidence/new', methods=['GET', 'POST'])
def new_evidence(premise_id):
    premise = excel_db.get_premise(premise_id)
    if not premise:
        flash('Premise not found.', 'danger')
        return redirect(url_for('index'))
    search = excel_db.get_search(premise['search_id'])

    if request.method == 'POST':
        form_data = request.form.to_dict()
        form_data['premise_id'] = premise_id
        form_data['location_found'] = request.form.get('location_found', '').strip()
        form_data['notes_by_officer'] = ''

        # For Cloud: if serial_number is empty, use email_id
        if form_data.get('evidence_category') == 'Cloud' and not form_data.get('serial_number'):
            form_data['serial_number'] = form_data.get('email_id', 'CLOUD-MAIL')

        # If is_owner is Yes and owner_name is supplied, update owner_user
        if form_data.get('is_owner') == 'Yes' and request.form.get('owner_name'):
            owner_n = request.form.get('owner_name', '').strip()
            if owner_n:
                form_data['owner_user'] = owner_n

        # Multi-checkboxes handling
        handling_list = request.form.getlist('device_handling_opts')
        form_data['device_handling'] = ', '.join(handling_list)

        control_list = request.form.getlist('device_control_opts')
        form_data['device_control'] = ', '.join(control_list)

        ev_id = excel_db.save_evidence(form_data)

        # Save sub-evidences if any
        sub_types = request.form.getlist('sub_type[]')
        sub_makes = request.form.getlist('sub_make[]')
        sub_models = request.form.getlist('sub_model[]')
        sub_sns = request.form.getlist('sub_sn[]')
        sub_caps = request.form.getlist('sub_capacity[]')
        sub_owners = request.form.getlist('sub_owner[]')

        for i in range(len(sub_makes)):
            if sub_makes[i].strip() or sub_sns[i].strip():
                excel_db.save_sub_evidence({
                    'evidence_id': ev_id,
                    'sub_type': sub_types[i] if i < len(sub_types) else 'SSD',
                    'make': sub_makes[i],
                    'model': sub_models[i] if i < len(sub_models) else '',
                    'sn': sub_sns[i] if i < len(sub_sns) else '',
                    'capacity': sub_caps[i] if i < len(sub_caps) else '',
                    'owner_user': sub_owners[i] if i < len(sub_owners) else ''
                })

        # Save custody rows if any
        c_reasons = request.form.getlist('custody_reason[]')
        c_givens = request.form.getlist('custody_given[]')
        c_receiveds = request.form.getlist('custody_received[]')
        c_dates = request.form.getlist('custody_date[]')
        c_times = request.form.getlist('custody_time[]')

        for j in range(len(c_reasons)):
            if c_reasons[j].strip() or c_givens[j].strip():
                excel_db.save_custody_row({
                    'evidence_id': ev_id,
                    'reason_action': c_reasons[j],
                    'given_by_name': c_givens[j] if j < len(c_givens) else '',
                    'given_by_signature': '[Signed]',
                    'received_by_name': c_receiveds[j] if j < len(c_receiveds) else '',
                    'received_by_signature': '[Signed]',
                    'date': c_dates[j] if j < len(c_dates) else premise.get('date', ''),
                    'time': c_times[j] if j < len(c_times) else ''
                })

        flash(f'Evidence {ev_id} recorded successfully.', 'success')
        return redirect(url_for('premise_detail', premise_id=premise_id))

    jur = search.get('jurisdiction', '')
    if jur == 'Mumbai':
        template_name = 'evidence_form_mumbai.html'
    elif jur == 'Gujarat':
        template_name = 'evidence_form_gujarat.html'
    elif jur == 'Delhi':
        template_name = 'evidence_form_delhi.html'
    else:
        template_name = 'evidence_form.html'

    return render_template(
        template_name,
        search=search,
        premise=premise,
        evidence=None,
        sub_evidences=[],
        custody_rows=[],
        categories=config.EVIDENCE_CATEGORIES,
        handling_options=config.DEVICE_HANDLING_OPTIONS,
        control_options=config.DEVICE_CONTROL_OPTIONS,
        hash_algorithms=config.HASH_ALGORITHMS
    )

@app.route('/evidence/<evidence_id>/edit', methods=['GET', 'POST'])
def edit_evidence(evidence_id):
    evidence = excel_db.get_evidence(evidence_id)
    if not evidence:
        flash('Evidence not found.', 'danger')
        return redirect(url_for('index'))

    premise = excel_db.get_premise(evidence['premise_id'])
    search = excel_db.get_search(premise['search_id'])

    if request.method == 'POST':
        form_data = request.form.to_dict()
        form_data['evidence_id'] = evidence_id
        form_data['premise_id'] = premise['premise_id']
        form_data['location_found'] = request.form.get('location_found', '').strip()
        form_data['notes_by_officer'] = ''

        # For Cloud: if serial_number is empty, use email_id
        if form_data.get('evidence_category') == 'Cloud' and not form_data.get('serial_number'):
            form_data['serial_number'] = form_data.get('email_id', 'CLOUD-MAIL')

        # If is_owner is Yes and owner_name is supplied, update owner_user
        if form_data.get('is_owner') == 'Yes' and request.form.get('owner_name'):
            owner_n = request.form.get('owner_name', '').strip()
            if owner_n:
                form_data['owner_user'] = owner_n

        handling_list = request.form.getlist('device_handling_opts')
        form_data['device_handling'] = ', '.join(handling_list)

        control_list = request.form.getlist('device_control_opts')
        form_data['device_control'] = ', '.join(control_list)

        excel_db.save_evidence(form_data)

        # Update sub-evidences: remove existing and re-insert
        existing_subs = excel_db.get_sub_evidences(evidence_id)
        for es in existing_subs:
            excel_db.delete_sub_evidence(es['sub_evidence_id'])

        sub_types = request.form.getlist('sub_type[]')
        sub_makes = request.form.getlist('sub_make[]')
        sub_models = request.form.getlist('sub_model[]')
        sub_sns = request.form.getlist('sub_sn[]')
        sub_caps = request.form.getlist('sub_capacity[]')
        sub_owners = request.form.getlist('sub_owner[]')

        for i in range(len(sub_makes)):
            if sub_makes[i].strip() or sub_sns[i].strip():
                excel_db.save_sub_evidence({
                    'evidence_id': evidence_id,
                    'sub_type': sub_types[i] if i < len(sub_types) else 'SSD',
                    'make': sub_makes[i],
                    'model': sub_models[i] if i < len(sub_models) else '',
                    'sn': sub_sns[i] if i < len(sub_sns) else '',
                    'capacity': sub_caps[i] if i < len(sub_caps) else '',
                    'owner_user': sub_owners[i] if i < len(sub_owners) else ''
                })

        # Update custody rows: remove existing and re-insert
        existing_custodies = excel_db.get_custody_rows(evidence_id)
        for ec in existing_custodies:
            excel_db.delete_custody_row(ec['custody_id'])

        c_reasons = request.form.getlist('custody_reason[]')
        c_givens = request.form.getlist('custody_given[]')
        c_receiveds = request.form.getlist('custody_received[]')
        c_dates = request.form.getlist('custody_date[]')
        c_times = request.form.getlist('custody_time[]')

        for j in range(len(c_reasons)):
            if c_reasons[j].strip() or c_givens[j].strip():
                excel_db.save_custody_row({
                    'evidence_id': evidence_id,
                    'reason_action': c_reasons[j],
                    'given_by_name': c_givens[j] if j < len(c_givens) else '',
                    'given_by_signature': '[Signed]',
                    'received_by_name': c_receiveds[j] if j < len(c_receiveds) else '',
                    'received_by_signature': '[Signed]',
                    'date': c_dates[j] if j < len(c_dates) else premise.get('date', ''),
                    'time': c_times[j] if j < len(c_times) else ''
                })

        flash(f'Evidence {evidence_id} updated successfully.', 'success')
        return redirect(url_for('premise_detail', premise_id=premise['premise_id']))

    sub_evs = excel_db.get_sub_evidences(evidence_id)
    custody_rows = excel_db.get_custody_rows(evidence_id)

    jur = search.get('jurisdiction', '')
    if jur == 'Mumbai':
        template_name = 'evidence_form_mumbai.html'
    elif jur == 'Gujarat':
        template_name = 'evidence_form_gujarat.html'
    elif jur == 'Delhi':
        template_name = 'evidence_form_delhi.html'
    else:
        template_name = 'evidence_form.html'

    return render_template(
        template_name,
        search=search,
        premise=premise,
        evidence=evidence,
        sub_evidences=sub_evs,
        custody_rows=custody_rows,
        categories=config.EVIDENCE_CATEGORIES,
        handling_options=config.DEVICE_HANDLING_OPTIONS,
        control_options=config.DEVICE_CONTROL_OPTIONS,
        hash_algorithms=config.HASH_ALGORITHMS
    )

@app.route('/evidence/<evidence_id>/delete', methods=['POST'])
def delete_evidence(evidence_id):
    ev = excel_db.get_evidence(evidence_id)
    premise_id = ev.get('premise_id') if ev else None
    excel_db.delete_evidence_cascade(evidence_id)
    flash(f'Evidence {evidence_id} deleted.', 'info')
    if premise_id:
        return redirect(url_for('premise_detail', premise_id=premise_id))
    return redirect(url_for('index'))

# ----------------- CLOUD EVIDENCE ROUTES (DELHI) -----------------
@app.route('/premise/<premise_id>/cloud/new', methods=['GET', 'POST'])
def new_cloud_evidence(premise_id):
    premise = excel_db.get_premise(premise_id)
    if not premise:
        flash('Premise not found.', 'danger')
        return redirect(url_for('index'))
    search = excel_db.get_search(premise['search_id'])

    if request.method == 'POST':
        sn = request.form.get('sn', '').strip()
        device_imaged = request.form.get('device_imaged', '').strip()
        size_of_data = request.form.get('size_of_data', '').strip()
        export_format = request.form.get('export_format', '').strip()
        source = request.form.get('source', '').strip()

        if not device_imaged:
            flash('Device / Account Imaged is required.', 'danger')
            return redirect(url_for('new_cloud_evidence', premise_id=premise_id))

        ce_id = excel_db.save_cloud_evidence({
            'premise_id': premise_id,
            'sn': sn,
            'device_imaged': device_imaged,
            'size_of_data': size_of_data,
            'export_format': export_format,
            'source': source
        })
        flash(f'Cloud evidence {ce_id} added.', 'success')
        return redirect(url_for('premise_detail', premise_id=premise_id))

    existing_clouds = excel_db.get_cloud_evidences(premise_id)
    next_sn = str(len(existing_clouds) + 1)

    return render_template(
        'cloud_evidence_form.html',
        search=search,
        premise=premise,
        cloud_evidence=None,
        next_sn=next_sn
    )

@app.route('/cloud/<cloud_id>/edit', methods=['GET', 'POST'])
def edit_cloud_evidence(cloud_id):
    ce = excel_db.get_by_id('CloudEvidences', cloud_id)
    if not ce:
        flash('Cloud evidence not found.', 'danger')
        return redirect(url_for('index'))
    premise = excel_db.get_premise(ce['premise_id'])
    search = excel_db.get_search(premise['search_id'])

    if request.method == 'POST':
        ce['sn'] = request.form.get('sn', '').strip()
        ce['device_imaged'] = request.form.get('device_imaged', '').strip()
        ce['size_of_data'] = request.form.get('size_of_data', '').strip()
        ce['export_format'] = request.form.get('export_format', '').strip()
        ce['source'] = request.form.get('source', '').strip()

        excel_db.save_cloud_evidence(ce)
        flash('Cloud evidence updated.', 'success')
        return redirect(url_for('premise_detail', premise_id=premise['premise_id']))

    return render_template(
        'cloud_evidence_form.html',
        search=search,
        premise=premise,
        cloud_evidence=ce,
        next_sn=ce.get('sn', '1')
    )

@app.route('/cloud/<cloud_id>/delete', methods=['POST'])
def delete_cloud_evidence_route(cloud_id):
    ce = excel_db.get_by_id('CloudEvidences', cloud_id)
    premise_id = ce.get('premise_id') if ce else None
    excel_db.delete_cloud_evidence(cloud_id)
    flash(f'Cloud evidence {cloud_id} deleted.', 'info')
    if premise_id:
        return redirect(url_for('premise_detail', premise_id=premise_id))
    return redirect(url_for('index'))

# ----------------- DOCUMENT GENERATION ROUTES -----------------
@app.route('/generate/evidence/<evidence_id>')
def generate_evidence_doc(evidence_id):
    try:
        result = doc_generator.generate_evidence_docs(evidence_id)
        if result.get('zip_path') and os.path.exists(result['zip_path']):
            return send_file(
                result['zip_path'],
                as_attachment=True,
                download_name=os.path.basename(result['zip_path'])
            )
        elif result.get('files'):
            return send_file(
                result['files'][0],
                as_attachment=True,
                download_name=os.path.basename(result['files'][0])
            )
        else:
            flash('No files generated.', 'warning')
            return redirect(request.referrer or url_for('index'))
    except Exception as e:
        flash(f'Error generating documents: {str(e)}', 'danger')
        return redirect(request.referrer or url_for('index'))

@app.route('/generate/search/<search_id>/annexure-e')
def generate_search_annexure_e_doc(search_id):
    try:
        out_path = doc_generator.generate_search_annexure_e(search_id)
        return send_file(
            out_path,
            as_attachment=True,
            download_name=os.path.basename(out_path)
        )
    except Exception as e:
        flash(f'Error generating Gujarat Annexure E: {str(e)}', 'danger')
        return redirect(request.referrer or url_for('index'))

@app.route('/generate/premise/<premise_id>/annexure-g')
def generate_premise_g(premise_id):
    try:
        premise = excel_db.get_premise(premise_id) or {}
        search = excel_db.get_search(premise.get('search_id', '')) or {}
        if search.get('jurisdiction') == 'Mumbai':
            out_path = doc_generator.generate_mumbai_annexure_1(premise_id)
        else:
            out_path = doc_generator.generate_premise_annexure_g(premise_id)
        return send_file(
            out_path,
            as_attachment=True,
            download_name=os.path.basename(out_path)
        )
    except Exception as e:
        flash(f'Error generating Annexure: {str(e)}', 'danger')
        return redirect(request.referrer or url_for('index'))

@app.route('/generate/premise/<premise_id>/mumbai-annexure-1')
def generate_mumbai_annexure_1_route(premise_id):
    try:
        out_path = doc_generator.generate_mumbai_annexure_1(premise_id)
        return send_file(
            out_path,
            as_attachment=True,
            download_name=os.path.basename(out_path)
        )
    except Exception as e:
        flash(f'Error generating Mumbai Annexure 1: {str(e)}', 'danger')
        return redirect(request.referrer or url_for('index'))

@app.route('/generate/premise/<premise_id>/mumbai-annexure-3')
def generate_mumbai_annexure_3_route(premise_id):
    try:
        out_path = doc_generator.generate_mumbai_annexure_3_coc(premise_id)
        return send_file(
            out_path,
            as_attachment=True,
            download_name=os.path.basename(out_path)
        )
    except Exception as e:
        flash(f'Error generating Mumbai Annexure 3 COC: {str(e)}', 'danger')
        return redirect(request.referrer or url_for('index'))

@app.route('/generate/premise/<premise_id>/annexure-f')
def generate_premise_f(premise_id):
    try:
        out_path = doc_generator.generate_premise_annexure_f(premise_id)
        return send_file(
            out_path,
            as_attachment=True,
            download_name=os.path.basename(out_path)
        )
    except Exception as e:
        flash(f'Error generating Annexure F: {str(e)}', 'danger')
        return redirect(request.referrer or url_for('index'))

@app.route('/database/download')
def download_database():
    if os.path.exists(config.DATABASE_PATH):
        return send_file(
            config.DATABASE_PATH,
            as_attachment=True,
            download_name='master_database.xlsx'
        )
    flash('Database file not found.', 'danger')
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("SEBI Digital Evidence Automation starting on http://127.0.0.1:5000 ...")
    app.run(debug=True, host='127.0.0.1', port=5000)
