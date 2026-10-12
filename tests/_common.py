import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
TEMPLATE = os.path.join(ROOT, 'assist_report_template.docx')
