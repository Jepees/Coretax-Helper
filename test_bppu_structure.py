"""Verify BPPU XML structure is correct (tags and order)."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
import re, zipfile

from modules.excel_reader import read_bppu
from modules.xml_generator import generate_bppu_xml

base = os.path.join(os.path.dirname(__file__), "Sample Data", "Data Format and Template From Coretax")

# Read Excel and generate XML
with open(os.path.join(base, "BPPU Excel to XML v.3.xlsx"), "rb") as f:
    result = read_bppu(f.read())
xml_output = generate_bppu_xml(result["tin"], result["rows"])

# Read expected from zip
with zipfile.ZipFile(os.path.join(base, "bppu.zip"), "r") as z:
    with z.open("bppu.xml") as f:
        expected = f.read().decode("utf-8")

# Extract tag structure (just tag names, ignoring content)
def get_tags(xml_str):
    return re.findall(r'</?(\w+)[> /]', xml_str)

our_tags = get_tags(xml_output)
exp_tags = get_tags(expected)

# Get unique tag sequence for first Bpu entry
def get_entry_tags(xml_str, entry_tag):
    m = re.search(f'<{entry_tag}>(.+?)</{entry_tag}>', xml_str, re.DOTALL)
    if m:
        return re.findall(r'<(\w+)>', m.group(1))
    return []

our_entry = get_entry_tags(xml_output, "Bpu")
exp_entry = get_entry_tags(expected, "Bpu")

print("=== BPPU Entry Tag Order ===")
print(f"Ours:     {our_entry}")
print(f"Expected: {exp_entry}")
print(f"Match: {our_entry == exp_entry}")

print(f"\n=== Root/List Tags ===")
# Check root element
print(f"Has BpuBulk: {'BpuBulk' in xml_output}")
print(f"Has TIN: {'<TIN>' in xml_output}")
print(f"Has ListOfBpu: {'<ListOfBpu>' in xml_output}")
print(f"Has SP2DNumber nil: {'xsi:nil' in xml_output}")
print(f"Bpu count ours: {xml_output.count('<Bpu>')}")
print(f"Bpu count expected: {expected.count('<Bpu>')}")
print(f"\nTIN ours: {result['tin']}")
print(f"Rows ours: {len(result['rows'])}")
