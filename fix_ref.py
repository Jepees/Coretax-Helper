"""Truncate reference_data.py to remove orphaned duplicate data."""
import os
path = os.path.join(os.path.dirname(__file__), "modules", "reference_data.py")
with open(path, "r", encoding="utf-8") as f:
    lines = f.readlines()
print(f"Before: {len(lines)} lines")
# Keep only lines 1-392 (index 0-391)
with open(path, "w", encoding="utf-8") as f:
    f.writelines(lines[:392])
with open(path, "r", encoding="utf-8") as f:
    new_lines = f.readlines()
print(f"After: {len(new_lines)} lines")
