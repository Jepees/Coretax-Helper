"""Compare all 4 types: App Output vs Official Excel Macro Output."""
import re

base_expected = r"D:\Dani's Projects\AUTOMATION (By Dani)\Coretax Helper\Sample Data\Data XML Ril Practice"
base_ours = r"D:\Dani's Projects\AUTOMATION (By Dani)\Coretax Helper\Sample Data\App Output"

comparisons = [
    ("BP21", "bp21_20260527_090551.xml", "bp21_test1.xml"),
    ("BPMP", "bpmp_20260527_090647.xml", "bpmp_test1.xml"),
    ("BPA1", "bpa1_20260527_090711.xml", "bpa1_test1.xml"),
    ("BPPU", "bppu_20260527_090736.xml", "bppu_test1.xml"),
]

results = {}

for label, our_file, exp_file in comparisons:
    print(f"\n{'='*70}")
    print(f"  {label}")
    print(f"{'='*70}")
    
    with open(f"{base_ours}\\{our_file}", "r", encoding="utf-8") as f:
        ours = f.read()
    with open(f"{base_expected}\\{exp_file}", "r", encoding="utf-8") as f:
        expected = f.read()
    
    our_lines = ours.splitlines()
    exp_lines = expected.splitlines()
    
    print(f"  Ours:     {len(our_lines)} lines, {len(ours)} chars")
    print(f"  Expected: {len(exp_lines)} lines, {len(expected)} chars")
    
    # Count entries
    entry_tags = {"BP21": "Bp21", "BPMP": "MmPayroll", "BPA1": "A1", "BPPU": "Bpu"}
    tag = entry_tags[label]
    our_count = len(re.findall(f"<{tag}>", ours))
    exp_count = len(re.findall(f"<{tag}>", expected))
    print(f"  Entries:  Ours={our_count}, Expected={exp_count}")
    
    # TIN
    our_tin = re.search(r"<TIN>(.*?)</TIN>", ours)
    exp_tin = re.search(r"<TIN>(.*?)</TIN>", expected)
    tin_match = our_tin and exp_tin and our_tin.group(1) == exp_tin.group(1)
    print(f"  TIN:      Ours={our_tin.group(1) if our_tin else 'N/A'}, Expected={exp_tin.group(1) if exp_tin else 'N/A'} {'OK' if tin_match else 'MISMATCH'}")
    
    # Line-by-line comparison (strip whitespace)
    diffs_content = []  # Real content differences
    diffs_cosmetic = 0  # Cosmetic (space before />)
    
    max_lines = max(len(our_lines), len(exp_lines))
    for i in range(max_lines):
        our_line = our_lines[i].strip() if i < len(our_lines) else "<MISSING>"
        exp_line = exp_lines[i].strip() if i < len(exp_lines) else "<MISSING>"
        
        if our_line != exp_line:
            # Check if it's just the cosmetic space before />
            if our_line.replace(' />', '/>') == exp_line.replace(' />', '/>'):
                diffs_cosmetic += 1
            else:
                diffs_content.append((i+1, our_line, exp_line))
    
    print(f"  Cosmetic diffs (space before />): {diffs_cosmetic}")
    print(f"  Content diffs: {len(diffs_content)}")
    
    if diffs_content:
        print(f"\n  First 15 content differences:")
        for line_num, ours_l, exp_l in diffs_content[:15]:
            print(f"    Line {line_num}:")
            print(f"      Ours:     {ours_l[:120]}")
            print(f"      Expected: {exp_l[:120]}")
    
    results[label] = {
        "entries_match": our_count == exp_count,
        "tin_match": tin_match,
        "content_diffs": len(diffs_content),
        "cosmetic_diffs": diffs_cosmetic,
        "pass": len(diffs_content) == 0,
    }

# Summary
print(f"\n{'='*70}")
print(f"  SUMMARY")
print(f"{'='*70}")
for label, r in results.items():
    status = "PASS" if r["pass"] else "FAIL"
    print(f"  {label}: {status} | Entries={r['entries_match']} | TIN={r['tin_match']} | ContentDiffs={r['content_diffs']} | CosmeticDiffs={r['cosmetic_diffs']}")
