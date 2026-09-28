# SEBI Digital Evidence Documentation & Annexure Automation System

A local, offline Python + Flask web application that automates the generation of legally compliant SEBI search and seizure digital evidence documentation, Chain of Custody (COC) registers, and Bharatiya Sakshya Adhiniyam (BSA) 2023 Sec 63(4)(c) electronic certificates in Microsoft Word (`.docx`) and `.zip` packages.

The system utilizes an offline, thread-safe master Excel workbook (`database/master_database.xlsx`) as its relational database and single source of truth, eliminating the need for external database servers or internet connectivity during search and seizure operations.

---

## 1. Quick Setup & Run

### Prerequisites
- Python 3.10+ (tested on Python 3.10 – 3.12)
- Microsoft Word or compatible viewer for inspecting generated `.docx` documents

### Local Setup
```bash
# 1. Create and activate a virtual environment
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the application
python app.py
```

Once running, access the web portal in your browser at:
**[http://127.0.0.1:5000](http://127.0.0.1:5000)**

### Standalone Executable Packaging (Optional)
The project includes a PyInstaller build specification (`build.spec`) to build a standalone, zero-dependency executable:
```bash
pip install pyinstaller
pyinstaller build.spec
```
The compiled binary and required assets will be placed in `dist/EvidenceAutomation/`.

---

## 2. Supported Jurisdictions

The application dynamically routes document generation, form structures, and validation rules based on the selected jurisdiction:

| Jurisdiction | Evidence Formats & Document Sets | Primary Outputs |
|---|---|---|
| **Mumbai** | Specialized Annexure 1 (Physical Inventory with rowspan grouping), Annexure 3 (Chain of Custody with 6-column handoff tables), Mobile & Non-Mobile Collection Forms, and BSA 2023 Sec 63(4)(c) Certificate with expert paternity details. | `.docx` and `.zip` |
| **Delhi** | Annexure B (Non-Mobile Device Collection), Annexure C (Chain of Custody), Annexure D (Mobile Collection), Annexure E (BSA 2023 Certificate), Annexure F (Cloud Evidence), Annexure G (Physical Inventory). | `.zip` per device, `.docx` for Premise F & G |
| **Gujarat** | Batch Annexure E (BSA 2023 Certificate) consolidated at the Search level for all evidences across all premises, with clean page breaks. Per-premise Annexure G inventory. | `.docx` |

---

## 3. Project Architecture & Directory Structure

```
sebi_evidence_automation/
├── app.py                         # Flask application entrypoint, route handlers, file dispatchers
├── config.py                      # Global configuration, ID prefixes, categories, and paths
├── excel_db.py                    # Thread-safe Excel database engine (auto-schema migration & formatting)
├── doc_generator.py               # Jurisdiction routing, context assembly, table styling & docxtpl rendering
├── build_templates.py             # Template builder for Gujarat & Delhi Word templates
├── build_mumbai_templates.py      # Template builder for the 5 specialized Mumbai Word templates
├── build.spec                     # PyInstaller specification for standalone executable builds
├── requirements.txt               # Python package dependencies (Flask, openpyxl, pandas, python-docx, docxtpl)
├── README.md                      # Complete system documentation and reference guide
├── database/
│   └── master_database.xlsx       # Relational Excel database containing 6 relational sheets
├── templates_docx/                # Compiled Word templates containing Jinja2 docxtpl placeholders
│   ├── annexure_b.docx            # Non-Mobile Collection Form (Delhi)
│   ├── annexure_c.docx            # Chain of Custody Form (Delhi)
│   ├── annexure_d.docx            # Mobile Collection Form (Delhi)
│   ├── annexure_e.docx            # BSA 2023 Sec 63(4)(c) Certificate (Delhi & Gujarat)
│   ├── annexure_f.docx            # Cloud / Email Evidence Form (Delhi premise level)
│   ├── annexure_g.docx            # Physical Inventory of Devices (Delhi & Gujarat premise level)
│   ├── mumbai_annexure_1.docx     # Physical Inventory of Relevant Devices (Mumbai premise level)
│   ├── mumbai_annexure_3_coc.docx # Chain of Custody with 6-column transfer table (Mumbai premise level)
│   ├── mumbai_device_mobile.docx  # Mobile Collection Form (Mumbai device level)
│   ├── mumbai_device_non_mobile.docx # Non-Mobile Collection Form (Mumbai device level)
│   └── mumbai_per_device_bsa.docx # BSA 2023 Sec 63(4)(c) Certificate (Mumbai device level)
├── templates/                     # Flask Jinja2 HTML templates
│   ├── base.html                  # Shared layout, navigation header, and notification alerts
│   ├── index.html                 # Search operations overview & summary dashboard
│   ├── search_form.html           # Create / edit search operation
│   ├── search_detail.html         # Search details, premises list, Gujarat batch export
│   ├── premise_form.html          # Create / edit premise location
│   ├── premise_detail.html        # Premise details, evidence table, inventory & COC generation buttons
│   ├── evidence_form.html         # Dynamic evidence capture form (category toggles, guidance banners)
│   └── cloud_evidence_form.html   # Cloud / email evidence form (Delhi specific)
├── static/
│   ├── css/style.css              # Modern UI styling, cards, status badges, responsive tables
│   └── js/script.js               # Category toggles, dynamic input validation, table row management
├── Documents/                     # Base source Word reference templates
└── generated_docs/                # Directory where generated .docx and .zip files are written
```

---

## 4. Master Excel Database Schema (`database/master_database.xlsx`)

The system uses `database/master_database.xlsx` as its single authoritative datastore. All primary keys are short, human-readable identifiers automatically generated by the engine (e.g. `S001`, `P001`, `E001`, `SE001`, `C001`, `CE001`).

> **Automatic Schema Migration:** When `excel_db.init_db()` runs, it checks existing workbooks against `SHEET_SCHEMAS`. Any missing columns are appended with official styling (dark blue header `#1F4E78`, white bold Calibri font, thin borders, and freeze panes) without overwriting existing data.

### Sheet 1: `Searches`
- **PK**: `search_id` (e.g. `S001`)
- **Columns**: `search_id`, `case_name`, `jurisdiction`, `date`, `authorised_officer`, `remarks`, `search_warrant_details`

### Sheet 2: `Premises`
- **PK**: `premise_id` (e.g. `P001`)
- **FK**: `search_id`
- **Columns**: `premise_id`, `search_id`, `premise_name`, `address`, `team`, `date`, `time`

### Sheet 3: `Evidences`
- **PK**: `evidence_id` (e.g. `E001`)
- **FK**: `premise_id`
- **Columns**:
  - **Identity & Category:** `evidence_id`, `premise_id`, `assessee_name`, `evidence_category`, `device_name`, `make`, `model`, `serial_number`, `capacity`, `color`, `owner_user`, `is_owner`, `owner_relation`, `os`
  - **Digital & Network Identifiers:** `email_id`, `cloud_id`, `ip_address`, `domain_name`
  - **Mobile Specifics:** `imei1`, `imei2`, `imsi`, `sim_card_present`, `sim_provider`, `media_card_present`, `media_card_make`, `media_card_model`, `media_card_sn`, `mobile_state`
  - **Boot & Physical State:** `hibernation_sleep`, `visible_on_screen`, `shutdown_type`
  - **Security & Handling:** `encryption_present`, `encryption_software`, `device_handling`, `imaging_tools_details`
  - **Forensic Hashes:** `hash_algorithm`, `md5_hash`, `sha1_hash`, `sha256_hash_1`, `sha256_hash_2`, `sha256_hash_3`, `sha256_hash_4`, `sha256_hash_5`
  - **Target Copies (Drives):** `master_copy_make`, `master_copy_model`, `master_copy_sn`, `master_copy_capacity`, `working_copy_make`, `working_copy_model`, `working_copy_sn`, `working_copy_capacity`
  - **Return & Verification:** `device_returned`, `returned_date`, `returned_time`, `witness_signature_taken`, `notes_by_officer`
  - **Signatory & BSA Fields:** `examiner_name`, `examiner_details`, `collection_date`, `collection_time`, `collection_address`, `authorised_officer`, `device_control`, `cert_place`, `party_name`, `party_relation`, `party_residing_at`, `expert_name`, `expert_father_name`, `expert_designation`, `other_specific_notes`

### Sheet 4: `SubEvidences` (Internal Storage Devices)
- **PK**: `sub_evidence_id` (e.g. `SE001`)
- **FK**: `evidence_id` (Parent Laptop, Desktop, or Server)
- **Columns**: `sub_evidence_id`, `evidence_id`, `sub_type`, `make`, `model`, `sn`, `capacity`, `owner_user`

### Sheet 5: `ChainOfCustody`
- **PK**: `custody_id` (e.g. `C001`)
- **FK**: `evidence_id`
- **Columns**: `custody_id`, `evidence_id`, `reason_action`, `given_by_name`, `given_by_signature`, `received_by_name`, `received_by_signature`, `date`, `time`

### Sheet 6: `CloudEvidences` (Delhi Specific)
- **PK**: `cloud_evidence_id` (e.g. `CE001`)
- **FK**: `premise_id`
- **Columns**: `cloud_evidence_id`, `premise_id`, `sn`, `device_imaged`, `size_of_data`, `export_format`, `source`

---

## 5. Detailed Jurisdiction Rules & Logic

### 5.1 Mumbai Specific Rules (Complete Implementation)

The Mumbai workflow implements strict, court-ready formatting standards across all annexures:

#### 1. Annexure 1 — Physical Inventory of Relevant Devices
- **Scope**: Generated at the **Premise Level** via the **"Generate Inventory (Annexure 1)"** button.
- **Single Consolidated Table**: All devices across the premise are consolidated into a single table with unified column widths, custom borders (`#CCCCCC`), cell padding, and light grey headers (`#EAEFF5`).
- **Standard Categories & Vertical Rowspan Merge**:
  - Devices are grouped in standard order: `Desktop`, `Laptop`, `Mobile`, `Pendrive`, `Harddisk`, `Server`, followed by any extra categories (`Cloud`, `File/Folder`, `Other`).
  - If multiple devices share a category (e.g., three Desktops), the "Evidence Particulars" category cell (Column 0) is **vertically merged (rowspan)** across all rows for that category with centered bold text.
  - If a standard category has no devices present, a single row is rendered with `-` in columns 1 through 6.
- **Column 1 ("Nature Of Evidence Found")**: Displays the location where the evidence was found on-site in the premise (e.g. `Trading Desk 4`, `MD Cabin`, `Accounts Dept`).
- **Column 2 ("Model Details")**:
  - **Evidence Name Header**: Precludes with `Evidence Name: <device_name>`.
  - For physical hardware: `Make/Model`, `S/N`, `Capacity`, `Color`, `OS`.
  - **IMEI Filtering**: IMEI 1 and IMEI 2 are **strictly displayed for Mobile phones only**. IMEIs are never printed for laptops, desktops, or storage media.
  - **Cloud / Email Records**: Strictly displays `Evidence Name` and `Email ID` only. Physical attributes, tenant IDs, IP addresses, and domain names are suppressed.
- **Column 4 ("Is the Owner of the Device?")**: If affirmative/Yes, dynamically outputs the owner's Name (e.g. `Suresh Rao`); if negative/No, outputs `No`.
- **Empty Remarks Column**: Column 6 ("Remarks") is intentionally left blank for manual on-site officer notations and physical signatures.
- **Header Details**: Team Code, Date, and Sign of the Authorised Officer are dynamically rendered in the document header.

#### 2. Annexure 3 — Mumbai Chain of Custody (COC)
- **Scope**: Generated at the **Premise Level** via the **"Generate COC (Annexure 3)"** button.
- **Table 0 (Device & Hash Inventory)**:
  - Repeating rows for every evidence item found at the premise.
  - Column 0: Evidence index, device name, make/model, S/N, OS, size, IMEI, email.
  - Column 2: Master Copy details (Make, Model, S/N, Capacity) and MD5/SHA1 hashes.
  - Column 4: Working Copy details (Make, Model, S/N, Capacity) and MD5/SHA1 hashes.
- **Table 1 (6-Column Custody Handoff Documentation)**:
  - Columns: `EVIDENCE DEVICES` | `PURPOSE OF TRANSFER OF CUSTODY` | `RELEASED BY (Name & Sign)` | `RECEIVED BY (Name & Sign)` | `DATE & TIME` | `REMARKS`
  - **Unique Target Copies Register**: In Table 1, the target hard drives (Master Copy and Working Copy) represent the evidence being handed over. The system aggregates all devices and generates a deduplicated register of **Unique Master and Working copy serial numbers**:
    - If 1 unique drive pair: `Evidence: \n Master copy S/N: <sn> \n Working copy S/N: <sn>`
    - If multiple unique drive pairs: `Evidence 1: \n Master copy S/N: <sn> \n Working copy S/N: <sn> \n\n Evidence 2: ...`
  - **Transfer 1 (After Forensic Backup & Hashing)**: Released by Authorised Officer, received by Forensic Expert. Date & Time populated automatically from premise entry.
  - **Transfer 2 (Conclusion of Search)**: Released by Forensic Expert, received by Custodian. Date & Time left blank for manual pen entry.

#### 3. Dedicated Jurisdiction Forms & Field Toggling
- **Jurisdiction-Specific Evidence Forms**: Mumbai (`evidence_form_mumbai.html`), Delhi (`evidence_form_delhi.html`), and Gujarat (`evidence_form_gujarat.html`) feature tailored layouts without clutter.
- **Dynamic JavaScript Show/Hide**: Category selector dynamically exposes only relevant fields (e.g. Mobile specs for Mobile; Boot/Shutdown for Computers; Email ID for Cloud).
- **Notes Removal**: `notes_by_officer` has been completely removed from all input forms and document templates.
- **Cloud Minimal Mode**: Cloud evidence captures exclusively the Email ID (no tenant ID, no IP address, no domain name).

---

### 5.2 Delhi Specific Rules
- **Device-Level Generation**:
  - `Mobile`: Packaged in `.zip` containing Annexure C (Chain of Custody), Annexure D (Mobile Collection), and Annexure E (BSA Certificate).
  - `Non-Mobile`: Packaged in `.zip` containing Annexure B (Non-Mobile Collection), Annexure C, and Annexure E.
- **Premise-Level Annexures**:
  - **Annexure G (Inventory)**: Consolidated physical device inventory including internal storage components (Sub-Evidences).
  - **Annexure F (Cloud Evidence)**: Consolidated table listing all cloud mailboxes, drives, and repositories imaged at that premise.

---

### 5.3 Gujarat Specific Rules
- **Search-Level Annexure E**:
  - Accessible via **"Generate Annexure E (All Evidences)"** on the Search Detail page.
  - Combines all evidences across all premises into one master Word document with clean page breaks between certificates.
- **Premise-Level Annexure G**: Physical inventory generated per premise.

---

## 6. Checkbox & Styling Standards

Per SEBI forensic documentation standards, checkboxes across all Word documents render as:
- **☑** (`\u2611`) when true / affirmative
- **☐** (`\u2610`) when false / unselected

All tables utilize standardized cell margins (top/bottom: 80 dxa, left/right: 80–100 dxa), subtle grid borders (`#CCCCCC`), and Calibri typography.

---

## 7. Web Application Endpoints

| Method | Endpoint | Description | Output |
|---|---|---|---|
| `GET` | `/` | Home dashboard; lists all search operations and summary statistics | HTML |
| `GET`, `POST` | `/search/new` | Create a new search operation (with jurisdiction selection) | HTML / Redirect |
| `GET` | `/search/<search_id>` | Search detail view (premises list, Gujarat batch action) | HTML |
| `GET`, `POST` | `/search/<search_id>/edit` | Edit search details | HTML / Redirect |
| `POST` | `/search/<search_id>/delete` | Cascade delete search and all child records | Redirect |
| `GET`, `POST` | `/search/<search_id>/premise/new` | Add a new premise to a search operation | HTML / Redirect |
| `GET` | `/premise/<premise_id>` | Premise detail view (evidences list, inventory & COC actions) | HTML |
| `GET`, `POST` | `/premise/<premise_id>/edit` | Edit premise details | HTML / Redirect |
| `POST` | `/premise/<premise_id>/delete` | Cascade delete premise and all child evidences | Redirect |
| `GET`, `POST` | `/premise/<premise_id>/evidence/new` | Record new digital evidence (with dynamic category toggling) | HTML / Redirect |
| `GET`, `POST` | `/evidence/<evidence_id>/edit` | Edit existing digital evidence record | HTML / Redirect |
| `POST` | `/evidence/<evidence_id>/delete` | Cascade delete evidence record | Redirect |
| `GET`, `POST` | `/premise/<premise_id>/cloud/new` | Record Delhi cloud evidence | HTML / Redirect |
| `GET`, `POST` | `/cloud/<cloud_id>/edit` | Edit Delhi cloud evidence | HTML / Redirect |
| `POST` | `/cloud/<cloud_id>/delete` | Delete Delhi cloud evidence | Redirect |
| `GET` | `/generate/evidence/<evidence_id>` | Generate per-device documentation | `.zip` or `.docx` |
| `GET` | `/generate/search/<search_id>/annexure-e` | Generate Gujarat batch Annexure E | `.docx` |
| `GET` | `/generate/premise/<premise_id>/mumbai-annexure-1` | Generate Mumbai Annexure 1 (Physical Inventory) | `.docx` |
| `GET` | `/generate/premise/<premise_id>/mumbai-annexure-3` | Generate Mumbai Annexure 3 (Chain of Custody) | `.docx` |
| `GET` | `/generate/premise/<premise_id>/annexure-g` | Generate Annexure G (Delhi / Gujarat inventory) | `.docx` |
| `GET` | `/generate/premise/<premise_id>/annexure-f` | Generate Annexure F (Delhi cloud evidence) | `.docx` |
| `GET` | `/database/download` | Download current `master_database.xlsx` workbook | `.xlsx` |

---

## 8. Template Rebuilding Guide

If base documents in `Documents/` are modified, the Word templates in `templates_docx/` can be recompiled using the dedicated template builder scripts:

### Rebuild Mumbai Templates
```bash
python build_mumbai_templates.py
```
Recompiles:
- `templates_docx/mumbai_annexure_1.docx`
- `templates_docx/mumbai_annexure_3_coc.docx`
- `templates_docx/mumbai_device_non_mobile.docx`
- `templates_docx/mumbai_device_mobile.docx`
- `templates_docx/mumbai_per_device_bsa.docx`

### Rebuild Gujarat & Delhi Templates
```bash
python build_templates.py
```
Recompiles Annexures B, C, D, E, F, and G.

---

## 9. Automated Testing & Verification

The repository includes a comprehensive test suite covering routes, document generation, and functional rules:

```bash
# 1. Full end-to-end Mumbai workflow test
python test_full_mumbai_workflow.py

# 2. Detailed functional rule validation (rowspan, IMEIs, hashes, checkboxes, COC 6-cols)
python test_mumbai_detailed_rules.py

# 3. Flask web routes test (Gujarat, Delhi, Mumbai)
python test_flask_routes.py

# 4. Mumbai-specific Flask route test
python test_flask_mumbai_routes.py

# 5. Standalone Annexure 1 generator test
python test_generate_annexure_1.py
```

All tests execute synchronously against the local database and verify output files in `generated_docs/`.

---

## 10. Extending to New Jurisdictions

To add another jurisdiction (e.g. `Karnataka` or `Kolkata`):
1. **Register Jurisdiction**: Add the jurisdiction name to `JURISDICTIONS` in `config.py`.
2. **Implement Routing Rules**: In `doc_generator.py`, add a branch in `generate_evidence_docs(evidence_id)` defining which annexure templates to render and whether to output `.zip` or `.docx`.
3. **Add Premise/Search Level UI Actions**: If the jurisdiction requires premise-level or search-level batch exports, add buttons in `templates/premise_detail.html` or `templates/search_detail.html` using Jinja conditional tags (e.g., `{% if search.jurisdiction == 'Karnataka' %}`).
4. **Excel Compatibility**: No database schema migrations or code rewrites are required—the master Excel engine automatically supports any valid jurisdiction string.
