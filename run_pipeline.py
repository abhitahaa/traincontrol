"""
run_pipeline.py  —  One command to rebuild everything from a CYCLES file.

HOW TO USE:
  1. Put this script in the same folder as:
       - build_master.py
       - populate_testcycle.py
       - testcycle.xlsx  (the template)
       - Your CYCLES .xlsx file  (any file with "CYCLES" in the name)
  2. Double-click run_pipeline.py  OR  run from terminal:
       python run_pipeline.py

OUTPUTS (saved to the same folder):
  - my_master.xlsx          clean intermediate master
  - testcycle_updated.xlsx  ready-to-upload portal file
"""

import os, sys, glob, time, subprocess

def find_cycles_file(folder):
    """Find the CYCLES Excel file in the same folder as this script."""
    matches = glob.glob(os.path.join(folder, "*CYCLES*.xlsx"))
    # Exclude output files that might accidentally match
    matches = [f for f in matches if "updated" not in f.lower() and "master" not in f.lower()]
    return matches

def main():
    base = os.path.dirname(os.path.abspath(__file__))

    build_script    = os.path.join(base, "build_master.py")
    populate_script = os.path.join(base, "populate_testcycle.py")
    template        = os.path.join(base, "testcycle.xlsx")
    master_out      = os.path.join(base, "my_master.xlsx")
    final_out       = os.path.join(base, "testcycle_updated.xlsx")

    # ── Check required scripts exist ─────────────────────────────────────────
    for f in [build_script, populate_script]:
        if not os.path.exists(f):
            print(f"ERROR: Missing script — {os.path.basename(f)}")
            print(f"       Make sure all scripts are in the same folder.")
            input("\nPress Enter to exit...")
            sys.exit(1)

    if not os.path.exists(template):
        print(f"ERROR: Missing template — testcycle.xlsx")
        print(f"       Make sure testcycle.xlsx is in the same folder.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # ── Auto-detect CYCLES file ───────────────────────────────────────────────
    matches = find_cycles_file(base)

    if len(matches) == 0:
        print("ERROR: No CYCLES file found in this folder.")
        print("       Add your CYCLES .xlsx file here and run again.")
        print(f"       Folder: {base}")
        input("\nPress Enter to exit...")
        sys.exit(1)

    elif len(matches) == 1:
        cycles_file = matches[0]
        print(f"Found CYCLES file: {os.path.basename(cycles_file)}")

    else:
        # Multiple CYCLES files — pick the most recently modified
        matches.sort(key=os.path.getmtime, reverse=True)
        cycles_file = matches[0]
        print(f"Multiple CYCLES files found — using most recent:")
        for i, f in enumerate(matches):
            marker = "  -> " if i == 0 else "     "
            print(f"{marker}{os.path.basename(f)}")

    print()

    # ── Step 1: Build master ──────────────────────────────────────────────────
    print("=" * 60)
    print("STEP 1 — Extracting data from CYCLES into master")
    print("=" * 60)
    t0 = time.time()
    result = subprocess.run(
        [sys.executable, build_script, "--cycles", cycles_file, "--output", master_out],
        check=False
    )
    if result.returncode != 0:
        print("\nStep 1 FAILED. Check the error above.")
        input("\nPress Enter to exit...")
        sys.exit(result.returncode)
    print(f"  Done in {time.time()-t0:.1f}s  ->  my_master.xlsx\n")

    # ── Step 2: Populate testcycle ────────────────────────────────────────────
    print("=" * 60)
    print("STEP 2 — Populating testcycle_updated from master")
    print("=" * 60)
    t1 = time.time()
    result = subprocess.run(
        [sys.executable, populate_script,
         "--master",   master_out,
         "--template", template,
         "--output",   final_out],
        check=False
    )
    if result.returncode != 0:
        print("\nStep 2 FAILED. Check the error above.")
        input("\nPress Enter to exit...")
        sys.exit(result.returncode)
    print(f"  Done in {time.time()-t1:.1f}s  ->  testcycle_updated.xlsx\n")

    # ── Done ──────────────────────────────────────────────────────────────────
    print("=" * 60)
    print(f"All done in {time.time()-t0:.1f}s")
    print(f"  my_master.xlsx         — clean reference master")
    print(f"  testcycle_updated.xlsx — ready to upload to portal")
    print("=" * 60)
    input("\nPress Enter to close...")

if __name__ == "__main__":
    main()
