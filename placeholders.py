#!/usr/bin/env python3
"""
Održava placeholder datoteke u praznim mapama (i briše ih iz mapa koje više nisu prazne).
Ne koristi .gitkeep nego .placeholder (skrivena datoteka).
"""

import os
import sys

PLACEHOLDER = ".placeholder"   # Ime placeholder datoteke

def is_git_dir(path):
    """Provjerava je li putanja unutar .git mape."""
    return ".git" in path.split(os.sep)

def should_keep_placeholder(dirpath):
    """Vraća True ako mapa nema NITI jedan drugi sadržaj osim placeholder datoteke."""
    try:
        entries = os.listdir(dirpath)
    except PermissionError:
        return False   # Preskoči mape kojima ne možemo pristupiti
    # Ako mapa nema ništa -> prazna je -> treba placeholder
    if not entries:
        return True
    # Ako ima samo jedan entry i to je placeholder -> i dalje prazna (samo placeholder)
    if len(entries) == 1 and entries[0] == PLACEHOLDER:
        return True
    # Ako ima išta drugo (bilo koju datoteku ili podmapu) -> NE treba placeholder
    return False

def manage_placeholders(root_dir):
    """Obilazi sve podmape i dodaje/briše placeholder po potrebi."""
    root_dir = os.path.abspath(root_dir)
    for dirpath, dirnames, filenames in os.walk(root_dir):
        # Preskoči .git mapu i sve ispod nje
        if is_git_dir(dirpath):
            continue
        
        placeholder_path = os.path.join(dirpath, PLACEHOLDER)
        needs_placeholder = should_keep_placeholder(dirpath)
        
        if needs_placeholder:
            if not os.path.exists(placeholder_path):
                # Stvori praznu placeholder datoteku
                with open(placeholder_path, 'w'):
                    pass
                print(f"➕ Dodan placeholder: {placeholder_path}")
        else:
            if os.path.exists(placeholder_path):
                os.remove(placeholder_path)
                print(f"➖ Uklonjen placeholder: {placeholder_path} (mapa više nije prazna)")

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "."
    if not os.path.isdir(target):
        print(f"Greška: '{target}' nije direktorij.")
        sys.exit(1)
    manage_placeholders(target)
    print("\nGotovo. Placeholder datoteke su usklađene sa stvarnim stanjem mapa.")