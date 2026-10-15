import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
TEMPLATE = os.path.join(ROOT, '保育巡查回報單.docx')
TATAKA_TEMPLATE = os.path.join(ROOT, '塔塔加保育宣導回報單.docx')
ECO_TEMPLATE = os.path.join(ROOT, '生態廁所回報單.docx')
