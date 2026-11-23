#!/usr/bin/env python3
"""Decode accent-skill.zip from parts."""
import base64, os, sys

parts = []
i = 1
while True:
    fn = f"accent-skill.zip.part{i}"
    if not os.path.exists(fn):
        break
    with open(fn, 'r') as f:
        parts.append(f.read())
    i += 1

if not parts:
    print("No parts found! Download all accent-skill.zip.part* files first.")
    sys.exit(1)

print(f"Found {len(parts)} parts, combining...")
data = base64.b64decode(''.join(parts))
with open('accent-skill.zip', 'wb') as f:
    f.write(data)
print(f"Created accent-skill.zip ({len(data):,} bytes)")
print("Now extract: unzip accent-skill.zip")
