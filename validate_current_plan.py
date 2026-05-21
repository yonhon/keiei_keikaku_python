from __future__ import annotations

import argparse
from pathlib import Path

from keiei_plan import (
    calculate,
    check_constraints,
    compare_with_workbook,
    read_assets,
    read_crop_schedules,
    read_current_plan,
)


DEFAULT_WORKBOOK = Path("経営計画_四本20260520.xlsx")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    args = parser.parse_args()
    path = args.workbook
    crops = read_crop_schedules(path)
    assets = read_assets(path)
    plan = read_current_plan(path)
    result = calculate(plan, crops, assets)

    mismatches = compare_with_workbook(path, result)
    print(f"Workbook: {path}")
    print(f"Crops loaded: {len(crops)}")
    print(f"Calculation mismatches: {len(mismatches)}")
    for mismatch in mismatches[:30]:
        print(f"  - {mismatch}")
    if len(mismatches) > 30:
        print(f"  ... and {len(mismatches) - 30} more")

    over_240 = [(i, h) for i, h in enumerate(result.labor_total) if h > 240]
    over_200 = [(i, h) for i, h in enumerate(result.labor_total) if h > 200]
    print(f"Months over 240h: {len(over_240)}")
    print(f"Months over 200h: {len(over_200)}")
    if over_200:
        print("Over-200h months:")
        for index, hours in over_200:
            print(f"  - month_index={index:02d}, labor={hours:.2f}")

    issues = check_constraints(plan, crops)
    print(f"Rotation issues: {len(issues)}")
    for issue in issues[:30]:
        print(f"  - {issue}")
    if len(issues) > 30:
        print(f"  ... and {len(issues) - 30} more")

    return 1 if mismatches else 0


if __name__ == "__main__":
    raise SystemExit(main())
