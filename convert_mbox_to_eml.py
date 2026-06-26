#!/usr/bin/env python3
"""
Convert Thunderbird mbox files to individual EML files.

Thunderbird stores emails in mbox format (no extension, .msf index files alongside).
This script extracts each message to a separate .eml file, preserving folder structure.

Usage:
    1. Edit src and dst paths below
    2. Run: python3 convert_mbox_to_eml.py
    3. Copy output (dst) to Windows staging path
"""

import mailbox
import os

# Source: Thunderbird Local Folders directory
# Linux example:
src = '/home/ruben/Downloads/ThunderbirdMail/Local Folders/'
# Windows example:
# src = 'C:\\Users\\YOURNAME\\AppData\\Roaming\\Thunderbird\\Profiles\\YOURPROFILE.default\\Mail\\Local Folders'

# Destination: where EML files will be written
# Linux example:
dst = '/home/ruben/Desktop/WorkMailEML/'
# Windows example:
# dst = 'C:\\Users\\Public\\Thunderbird_Staging'

for root, dirs, files in os.walk(src):
    for f in files:
        # Skip .msf index files (Thunderbird metadata, not actual emails)
        if not f.endswith('.msf') and os.path.isfile(os.path.join(root, f)):
            mbox_path = os.path.join(root, f)

            # Preserve folder structure in output
            rel_path = os.path.relpath(mbox_path, src)
            out_dir = os.path.join(dst, rel_path)
            os.makedirs(out_dir, exist_ok=True)

            try:
                box = mailbox.mbox(mbox_path)
                for i, msg in enumerate(box):
                    with open(os.path.join(out_dir, f'{i}.eml'), 'wb') as out:
                        out.write(msg.as_bytes())
                print(f"[OK] {rel_path} -> {len(box)} messages")
            except Exception as e:
                print(f"[FAIL] {rel_path}: {e}")
