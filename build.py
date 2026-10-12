"""
Foreman 2-Week Lookahead — Build Script
========================================
Usage:
    python build.py [configs/<job_name>.py]

Default config: configs/boiler_house.py

Example:
    python build.py
    python build.py configs/boiler_house.py
    python build.py configs/my_new_job.py

Output files are written to the current directory.
"""

import importlib
import sys
import os

def main():
    # Resolve config module name from CLI arg or default
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        # Accept "configs/boiler_house.py" or "configs.boiler_house"
        mod_name = arg.replace("/", ".").replace("\\", ".").removesuffix(".py")
    else:
        mod_name = "configs.boiler_house"

    # Ensure the template directory is on the path
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)

    print(f"Loading config: {mod_name}")
    cfg = importlib.import_module(mod_name)

    from lookahead_engine import build
    output_dir = here   # write xlsx + bas alongside build.py
    xlsx_path, bas_path = build(cfg, output_dir=output_dir)

    print()
    print("=" * 60)
    print(f"  Project : {cfg.PROJECT_NAME}")
    print(f"  Output  : {xlsx_path}")
    print(f"  VBA     : {bas_path}")
    print("=" * 60)
    print("Next steps:")
    print("  1. Open the .xlsx in Excel")
    print("  2. Save as .xlsm (macro-enabled workbook)")
    print("  3. Alt+F11 → Insert Module → paste the .bas file contents")
    print("  4. Add 3 Form Control buttons: NextWeek / PrevWeek / SaveThisWeek")
    print("=" * 60)


if __name__ == "__main__":
    main()
