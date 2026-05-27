"""Test BPA1 and BPPU conversion against official XML outputs."""
import sys
import os
import zipfile
import re

sys.path.insert(0, os.path.dirname(__file__))

from modules.excel_reader import read_bpa1, read_bppu
from modules.xml_generator import generate_bpa1_xml, generate_bppu_xml

base = os.path.join(os.path.dirname(__file__), "Sample Data", "Data Format and Template From Coretax")

def test_type(name, excel_file, zip_file, reader_fn, generator_fn):
    print(f"\n{'='*60}")
    print(f"TESTING: {name}")
    print(f"{'='*60}")
    
    # Read Excel
    with open(os.path.join(base, excel_file), "rb") as f:
        result = reader_fn(f.read())
    
    print(f"TIN: {result['tin']}")
    print(f"Rows: {len(result['rows'])}")
    print(f"Errors: {result['errors']}")
    
    # Generate XML
    xml_output = generator_fn(result["tin"], result["rows"])
    
    # Read expected XML from zip
    with zipfile.ZipFile(os.path.join(base, zip_file), "r") as z:
        xml_name = z.namelist()[0]
        with z.open(xml_name) as f:
            expected = f.read().decode("utf-8")
    
    # Normalize: strip \r and trailing whitespace
    our_lines = [l.strip() for l in xml_output.splitlines() if l.strip()]
    exp_lines = [l.strip() for l in expected.splitlines() if l.strip()]
    
    print(f"\nOur lines: {len(our_lines)}")
    print(f"Expected lines: {len(exp_lines)}")
    
    # Compare
    diffs = []
    max_lines = max(len(our_lines), len(exp_lines))
    for i in range(max_lines):
        our = our_lines[i] if i < len(our_lines) else "<MISSING>"
        exp = exp_lines[i] if i < len(exp_lines) else "<MISSING>"
        if our != exp:
            # Ignore cosmetic space before /> 
            if our.replace(' />', '/>') == exp.replace(' />', '/>'):
                continue
            diffs.append((i+1, our, exp))
    
    if diffs:
        print(f"\nDifferences: {len(diffs)}")
        for line_num, ours, exp in diffs[:20]:
            print(f"  Line {line_num}:")
            print(f"    Ours:     {ours}")
            print(f"    Expected: {exp}")
    else:
        print(f"\n✅ PASS — XML output matches expected (ignoring cosmetic spacing)!")
    
    return len(diffs) == 0

# Test BPA1
ok1 = test_type("BPA1", "BPA1 Excel to XML.xlsx", "bpa1.zip", read_bpa1, generate_bpa1_xml)

# Test BPPU
ok2 = test_type("BPPU", "BPPU Excel to XML v.3.xlsx", "bppu.zip", read_bppu, generate_bppu_xml)

print(f"\n{'='*60}")
print(f"RESULTS: BPA1={'PASS' if ok1 else 'FAIL'}, BPPU={'PASS' if ok2 else 'FAIL'}")
print(f"{'='*60}")
