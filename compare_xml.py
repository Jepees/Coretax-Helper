"""Compare our app's XML output with the official Excel export."""

our_file = r"C:\Users\User\Downloads\bpmp_20260526_233009.xml"
expected_file = r"D:\Dani's Projects\AUTOMATION (By Dani)\Coretax Helper\Sample Data\Data Test\bpmp_test1.xml"

with open(our_file, "r", encoding="utf-8") as f:
    ours = f.read()

with open(expected_file, "r", encoding="utf-8") as f:
    expected = f.read()

print(f"Our file:      {len(ours)} chars, {len(ours.splitlines())} lines")
print(f"Expected file: {len(expected)} chars, {len(expected.splitlines())} lines")

our_lines = ours.splitlines()
exp_lines = expected.splitlines()

# Compare line by line
diffs = []
max_lines = max(len(our_lines), len(exp_lines))
for i in range(max_lines):
    our_line = our_lines[i] if i < len(our_lines) else "<MISSING>"
    exp_line = exp_lines[i] if i < len(exp_lines) else "<MISSING>"
    if our_line.strip() != exp_line.strip():
        diffs.append((i+1, our_line.strip(), exp_line.strip()))

print(f"\nTotal differences: {len(diffs)}")

if diffs:
    print("\nFirst 30 differences:")
    for line_num, ours_l, exp_l in diffs[:30]:
        print(f"  Line {line_num}:")
        print(f"    Ours:     {ours_l}")
        print(f"    Expected: {exp_l}")
else:
    print("\nFILES ARE IDENTICAL!")

# Also check: same number of MmPayroll entries?
import re
our_count = len(re.findall(r"<MmPayroll>", ours))
exp_count = len(re.findall(r"<MmPayroll>", expected))
print(f"\nMmPayroll entries - Ours: {our_count}, Expected: {exp_count}")

# Check TIN
our_tin = re.search(r"<TIN>(.*?)</TIN>", ours)
exp_tin = re.search(r"<TIN>(.*?)</TIN>", expected)
print(f"TIN - Ours: {our_tin.group(1) if our_tin else 'N/A'}, Expected: {exp_tin.group(1) if exp_tin else 'N/A'}")
