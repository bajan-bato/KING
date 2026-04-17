#!/usr/bin/env python3
"""
Check joined.xlsx for duplicate otpremnice or serijski numbers across different (EL, ID) pairs.
- Warning: same Otpremnica used for multiple distinct (EL, ID) pairs (shows Serijski).
- Error: same Serijski used for multiple distinct (EL, ID) pairs (shows Otpremnica).
"""

import pandas as pd
from pathlib import Path
from collections import defaultdict

JOINED_FILE = "out/joined.xlsx"
OUTPUT_REPORT = "out/duplicate_report.txt"

def main():
    if not Path(JOINED_FILE).exists():
        print(f"Error: {JOINED_FILE} not found.")
        return

    df = pd.read_excel(JOINED_FILE, engine='openpyxl')
    required_cols = ['EL', 'ID', 'Otpremnica', 'Serijski']
    for col in required_cols:
        if col not in df.columns:
            print(f"Error: Column '{col}' missing in {JOINED_FILE}")
            return

    # For each otpremnica, store set of (EL, ID, Serijski)
    otp_to_entries = defaultdict(set)
    for _, row in df.iterrows():
        el = str(row['EL']).strip()
        id_val = str(row['ID']).strip()
        otp = str(row['Otpremnica']).strip()
        ser = str(row['Serijski']).strip()
        if otp:
            otp_to_entries[otp].add((el, id_val, ser))

    # For each serijski, store set of (EL, ID, Otpremnica)
    ser_to_entries = defaultdict(set)
    for _, row in df.iterrows():
        el = str(row['EL']).strip()
        id_val = str(row['ID']).strip()
        otp = str(row['Otpremnica']).strip()
        ser = str(row['Serijski']).strip()
        if ser:
            ser_to_entries[ser].add((el, id_val, otp))

    # Find duplicates: more than one distinct (EL, ID) pair (ignore third field)
    otp_warnings = {}
    for otp, entries in otp_to_entries.items():
        distinct_pairs = set((el, id_val) for (el, id_val, _) in entries)
        if len(distinct_pairs) > 1:
            otp_warnings[otp] = entries

    ser_errors = {}
    for ser, entries in ser_to_entries.items():
        distinct_pairs = set((el, id_val) for (el, id_val, _) in entries)
        if len(distinct_pairs) > 1:
            ser_errors[ser] = entries

    # Print summary
    print("="*80)
    print("DUPLICATE CHECK REPORT (distinct EL-ID pairs)")
    print("="*80)
    print(f"Total rows in joined.xlsx: {len(df)}")
    print()

    if otp_warnings:
        print(f"⚠️  WARNINGS: {len(otp_warnings)} Otpremnice used for multiple distinct (EL, ID) pairs:")
        for otp, entries in sorted(otp_warnings.items()):
            print(f"\n  Otpremnica: {otp}")
            for el, id_val, ser in sorted(entries):
                print(f"      - {el}-{id_val} : Serijski = {ser}")
        print()
    else:
        print("✅ No duplicate Otpremnice found.\n")

    if ser_errors:
        print(f"❌ ERRORS: {len(ser_errors)} Serijski numbers used for multiple distinct (EL, ID) pairs:")
        for ser, entries in sorted(ser_errors.items()):
            print(f"\n  Serijski: {ser}")
            for el, id_val, otp in sorted(entries):
                print(f"      - {el}-{id_val} : Otpremnica = {otp}")
        print()
    else:
        print("✅ No duplicate Serijski numbers found.\n")

    # Save detailed report
    with open(OUTPUT_REPORT, 'w', encoding='utf-8') as f:
        f.write("DUPLICATE CHECK REPORT (distinct EL-ID pairs)\n")
        f.write("="*80 + "\n")
        if otp_warnings:
            f.write("\nWARNINGS (duplicate Otpremnice):\n")
            for otp, entries in sorted(otp_warnings.items()):
                f.write(f"\nOtpremnica: {otp}\n")
                for el, id_val, ser in sorted(entries):
                    f.write(f"  {el}-{id_val} : Serijski = {ser}\n")
        if ser_errors:
            f.write("\nERRORS (duplicate Serijski):\n")
            for ser, entries in sorted(ser_errors.items()):
                f.write(f"\nSerijski: {ser}\n")
                for el, id_val, otp in sorted(entries):
                    f.write(f"  {el}-{id_val} : Otpremnica = {otp}\n")
        if not otp_warnings and not ser_errors:
            f.write("No duplicates found.\n")
    print(f"Detailed report saved to {OUTPUT_REPORT}")

if __name__ == "__main__":
    main()