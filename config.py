import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATABASE_PATH = os.path.join(BASE_DIR, 'database', 'master_database.xlsx')
TEMPLATES_DOCX_DIR = os.path.join(BASE_DIR, 'templates_docx')
GENERATED_DOCS_DIR = os.path.join(BASE_DIR, 'generated_docs')

SECRET_KEY = os.environ.get('SECRET_KEY', 'sebi-forensic-automation-secret-key-2026')

JURISDICTIONS = ['Gujarat', 'Delhi', 'Mumbai']

EVIDENCE_CATEGORIES = [
    'Laptop',
    'Desktop',
    'Server',
    'Mobile',
    'HDD',
    'SSD',
    'Pendrive',
    'File/Folder',
    'Cloud',
    'Other'
]

SUB_EVIDENCE_TYPES = ['HDD', 'SSD', 'Other']

SHUTDOWN_TYPES = ['Normal', 'PowerPlugPulled', 'BatteryRemoved']

MOBILE_STATES = ['On', 'Off', 'Offline']

DEVICE_HANDLING_OPTIONS = ['Seizure', 'Imaging', 'Backup', 'Reporting']

DEVICE_CONTROL_OPTIONS = ['Owned', 'Maintained', 'Managed', 'Operated']

HASH_ALGORITHMS = ['SHA256', 'SHA1', 'MD5', 'Other']

ID_PREFIXES = {
    'Searches': 'S',
    'Premises': 'P',
    'Evidences': 'E',
    'SubEvidences': 'SE',
    'ChainOfCustody': 'C',
    'CloudEvidences': 'CE'
}

ID_LENGTHS = {
    'Searches': 3,
    'Premises': 3,
    'Evidences': 3,
    'SubEvidences': 3,
    'ChainOfCustody': 3,
    'CloudEvidences': 3
}
