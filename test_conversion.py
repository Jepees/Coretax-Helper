"""Quick test: Read sample Excel files and verify XML structure."""

import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

from modules.excel_reader import read_bp21, read_bpmp
from modules.xml_generator import generate_bp21_xml, generate_bpmp_xml


def test_bp21():
    print("=" * 60)
    print("Testing BP21")
    print("=" * 60)

    with open(r"sample data\BP21 Excel to XML v.4.xlsx", "rb") as f:
        file_bytes = f.read()

    result = read_bp21(file_bytes)
    print(f"TIN: {result['tin']}")
    print(f"Rows: {len(result['rows'])}")
    print(f"Errors: {result['errors']}")

    # Print summary per row
    for i, row in enumerate(result["rows"]):
        print(f"  Row {i+1}: {row['CounterpartTin']} | {row['TaxObjectCode']} | Gross={row['Gross']} | Deemed={row['Deemed']} | Rate={row['Rate']}")

    # Generate XML
    xml = generate_bp21_xml(result["tin"], result["rows"])
    print(f"\nGenerated XML length: {len(xml)} chars")

    # Check structure
    assert "<Bp21Bulk" in xml, "Missing Bp21Bulk root"
    assert "<ListOfBp21>" in xml, "Missing ListOfBp21"
    assert "<Bp21>" in xml, "Missing Bp21 element"
    assert f"<TIN>{result['tin']}</TIN>" in xml, "Missing TIN"
    assert '<Bp21Bulk xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' in xml
    print("[OK] BP21 XML structure is correct")

    # Verify field ordering matches expected (compare with sample)
    expected_path = r"sample data\bp21_extracted\bp21.xml"
    if os.path.exists(expected_path):
        with open(expected_path, "r", encoding="utf-8") as f:
            expected = f.read()
        # Check that field order matches
        expected_tags = []
        for line in expected.split("\n"):
            line = line.strip()
            if line.startswith("<") and not line.startswith("<?") and not line.startswith("</") and not line.startswith("<Bp21Bulk") and not line.startswith("<ListOfBp21") and not line.startswith("<TIN") and not line.startswith("<Bp21>"):
                tag = line.split(">")[0].replace("<", "")
                if tag not in expected_tags:
                    expected_tags.append(tag)

        generated_tags = []
        for line in xml.split("\n"):
            line = line.strip()
            if line.startswith("<") and not line.startswith("<?") and not line.startswith("</") and not line.startswith("<Bp21Bulk") and not line.startswith("<ListOfBp21") and not line.startswith("<TIN") and not line.startswith("<Bp21>"):
                tag = line.split(">")[0].replace("<", "")
                if tag not in generated_tags:
                    generated_tags.append(tag)

        print(f"\nExpected field order:  {expected_tags}")
        print(f"Generated field order: {generated_tags}")
        if expected_tags == generated_tags:
            print("[OK] Field order matches!")
        else:
            print("[WARN] Field order differs")


def test_bpmp():
    print("\n" + "=" * 60)
    print("Testing BPMP")
    print("=" * 60)

    with open(r"sample data\BPMP Excel to XML v.3.xlsx", "rb") as f:
        file_bytes = f.read()

    result = read_bpmp(file_bytes)
    print(f"TIN: {result['tin']}")
    print(f"Rows: {len(result['rows'])}")
    print(f"Errors: {result['errors']}")

    for i, row in enumerate(result["rows"]):
        print(f"  Row {i+1}: {row['CounterpartTin']} | {row['TaxObjectCode']} | Gross={row['Gross']} | Rate={row['Rate']} | Opt={row['CounterpartOpt']}")

    # Generate XML
    xml = generate_bpmp_xml(result["tin"], result["rows"])
    print(f"\nGenerated XML length: {len(xml)} chars")

    # Check structure
    assert "<MmPayrollBulk" in xml, "Missing MmPayrollBulk root"
    assert "<ListOfMmPayroll>" in xml, "Missing ListOfMmPayroll"
    assert "<MmPayroll>" in xml, "Missing MmPayroll element"
    assert f"<TIN>{result['tin']}</TIN>" in xml, "Missing TIN"
    assert 'xsi:nil="true"' in xml, "Missing xsi:nil for passport"
    print("[OK] BPMP XML structure is correct")

    # Compare with expected
    expected_path = r"sample data\bpmp_extracted\bpmp.xml"
    if os.path.exists(expected_path):
        with open(expected_path, "r", encoding="utf-8") as f:
            expected = f.read()

        expected_tags = []
        for line in expected.split("\n"):
            line = line.strip()
            if line.startswith("<") and not line.startswith("<?") and not line.startswith("</") and not line.startswith("<MmPayrollBulk") and not line.startswith("<ListOfMmPayroll") and not line.startswith("<TIN") and not line.startswith("<MmPayroll>"):
                tag = line.split(">")[0].split(" ")[0].replace("<", "")
                if tag not in expected_tags:
                    expected_tags.append(tag)

        generated_tags = []
        for line in xml.split("\n"):
            line = line.strip()
            if line.startswith("<") and not line.startswith("<?") and not line.startswith("</") and not line.startswith("<MmPayrollBulk") and not line.startswith("<ListOfMmPayroll") and not line.startswith("<TIN") and not line.startswith("<MmPayroll>"):
                tag = line.split(">")[0].split(" ")[0].replace("<", "")
                if tag not in generated_tags:
                    generated_tags.append(tag)

        print(f"\nExpected field order:  {expected_tags}")
        print(f"Generated field order: {generated_tags}")
        if expected_tags == generated_tags:
            print("[OK] Field order matches!")
        else:
            print("[WARN] Field order differs")


if __name__ == "__main__":
    test_bp21()
    test_bpmp()
    print("\n" + "=" * 60)
    print("All tests complete!")
