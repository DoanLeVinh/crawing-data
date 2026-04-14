from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from analysis.recommender_web_data import run_full_recommendation_pipeline


def main() -> None:
    out_file = Path("outputs/rules/association_rules_products.csv")
    out_file.parent.mkdir(parents=True, exist_ok=True)
    result = run_full_recommendation_pipeline(
        csv_path="HD-csv.csv",
        min_support=0.01,
        min_confidence=0.30,
        min_lift=1.20,
        max_len=3,
    )
    rules_export = result["rules_export"]
    rules_export.to_csv(out_file, index=False, encoding="utf-8-sig")

    print(f"Exported rules: {out_file.resolve()}")
    print(f"Rows: {len(rules_export)}")


if __name__ == "__main__":
    main()
