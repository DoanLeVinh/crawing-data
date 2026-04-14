from __future__ import annotations

from pathlib import Path
from typing import Any
from collections import defaultdict

import pandas as pd

from analysis.basket_analysis import AnalysisResult, mine_association_rules


DEFAULT_DATA_FILE = "HD-csv.csv"


def _norm_item(name: str) -> str:
    return " ".join(str(name).strip().lower().split())


def _safe_float(value: str, default: float = 0.0) -> float:
    try:
        return float(str(value).strip())
    except Exception:
        return default


def _is_excluded_product(name: str) -> bool:
    s = _norm_item(name)
    return "san pham bo sung" in s


def _is_reasonable_pair(a: str, b: str, support: float, lift: float, pair_count: int = 0) -> bool:
    # Giu logic thuần association rules: chi can du metric co nghia.
    return bool(a and b and a != b and support > 0 and lift >= 1.0)


def load_invoice_lines(csv_path: str | Path = DEFAULT_DATA_FILE) -> pd.DataFrame:
    path = Path(csv_path)
    if not path.exists():
        raise FileNotFoundError(f"Data file not found: {path}")

    df = pd.read_csv(path, encoding="utf-8-sig")
    required_cols = ["receiptId", "productNames", "quantities", "unitPrices", "lineAmounts"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        products = [x.strip() for x in str(row.get("productNames", "")).split("|")]
        quantities = [x.strip() for x in str(row.get("quantities", "")).split("|")]
        unit_prices = [x.strip() for x in str(row.get("unitPrices", "")).split("|")]
        line_amounts = [x.strip() for x in str(row.get("lineAmounts", "")).split("|")]

        max_len = max(len(products), len(quantities), len(unit_prices), len(line_amounts))
        for i in range(max_len):
            product_name = products[i] if i < len(products) else ""
            if not product_name.strip():
                continue
            if _is_excluded_product(product_name):
                continue

            quantity = _safe_float(quantities[i], 1.0) if i < len(quantities) else 1.0
            unit_price = _safe_float(unit_prices[i], 0.0) if i < len(unit_prices) else 0.0
            line_amount = _safe_float(line_amounts[i], 0.0) if i < len(line_amounts) else 0.0

            rows.append(
                {
                    "receipt_id": str(row.get("receiptId", "")).strip(),
                    "product_name": product_name.strip(),
                    "product_key": _norm_item(product_name),
                    "quantity": quantity,
                    "unit_price": unit_price,
                    "line_amount": line_amount,
                    "platform": str(row.get("platform", "")).strip(),
                    "receipt_datetime": str(row.get("receiptDateTime", "")).strip(),
                }
            )

    if not rows:
        raise ValueError("No invoice line data could be parsed")

    return pd.DataFrame(rows)


def build_product_catalog(invoice_lines: pd.DataFrame) -> pd.DataFrame:
    catalog = (
        invoice_lines.groupby("product_key", as_index=False)
        .agg(
            product_name=("product_name", "first"),
            order_count=("receipt_id", "nunique"),
            line_count=("receipt_id", "size"),
            total_quantity=("quantity", "sum"),
            avg_unit_price=("unit_price", "mean"),
            total_revenue=("line_amount", "sum"),
            platform_example=("platform", "first"),
        )
        .sort_values(["order_count", "line_count"], ascending=False)
        .reset_index(drop=True)
    )
    catalog["product_id"] = catalog.index + 1
    catalog = catalog[
        [
            "product_id",
            "product_key",
            "product_name",
            "order_count",
            "line_count",
            "total_quantity",
            "avg_unit_price",
            "total_revenue",
            "platform_example",
        ]
    ]
    return catalog


def build_recommendation_index(
    raw_rules: pd.DataFrame,
    source_label: str = "association_rule",
    include_reverse: bool = False,
) -> dict[str, list[dict[str, Any]]]:
    rec_idx: dict[str, list[dict[str, Any]]] = {}
    if raw_rules is None or raw_rules.empty:
        return rec_idx

    pair_rules = raw_rules[
        (raw_rules["antecedents"].apply(len) == 1)
        & (raw_rules["consequents"].apply(len) == 1)
    ].copy()

    per_antecedent: dict[str, dict[str, dict[str, Any]]] = {}

    for _, r in pair_rules.iterrows():
        ant = _norm_item(next(iter(r["antecedents"])))
        cons = _norm_item(next(iter(r["consequents"])))
        if not ant or not cons or ant == cons:
            continue
        if _is_excluded_product(ant) or _is_excluded_product(cons):
            continue

        # Chi nhan goi y duoc sinh tu association rules hop le.
        confidence = float(r.get("confidence", 0.0))
        lift = float(r.get("lift", 0.0))
        support = float(r.get("support", 0.0))
        consequent_support = float(r.get("consequent support", support))
        # Loai bo goi y qua pho thong (vd: item xuat hien qua nhieu don) va tin hieu yeu.
        if confidence < 0.03 or lift <= 1.0 or support <= 0.0:
            continue
        if consequent_support >= 0.45:
            continue
        if not _is_reasonable_pair(ant, cons, support, lift):
            continue

        # Xep hang thuần theo do manh luat ket hop.
        lift_capped = min(lift, 8.0)
        support_weight = max(0.01, support ** 0.5)
        novelty_weight = max(0.05, (1.0 - consequent_support) ** 1.5)
        score = confidence * lift_capped * support_weight * novelty_weight

        candidate = {
            "product_key": cons,
            "confidence": confidence,
            "lift": lift,
            "support": support,
            "consequent_support": consequent_support,
            "score": score,
            "source": source_label,
        }

        ant_bucket = per_antecedent.setdefault(ant, {})
        existing = ant_bucket.get(cons)
        if existing is None or candidate["score"] > existing["score"]:
            ant_bucket[cons] = candidate

        if include_reverse:
            # Them chieu nguoc tu cung mot luat de tang do phu, van dua tren association rules.
            reverse_confidence = 0.0
            if consequent_support > 0:
                reverse_confidence = support / consequent_support
            antecedent_support = float(r.get("antecedent support", support))
            reverse_lift_capped = min(lift, 8.0)
            reverse_support_weight = max(0.01, support ** 0.5)
            reverse_novelty_weight = max(0.05, (1.0 - antecedent_support) ** 1.5)
            reverse_score = (
                reverse_confidence
                * reverse_lift_capped
                * reverse_support_weight
                * reverse_novelty_weight
            )
            if reverse_confidence >= 0.03 and antecedent_support < 0.45:
                reverse_candidate = {
                    "product_key": ant,
                    "confidence": reverse_confidence,
                    "lift": lift,
                    "support": support,
                    "consequent_support": antecedent_support,
                    "score": reverse_score,
                    "source": f"{source_label}_reverse",
                }
                rev_bucket = per_antecedent.setdefault(cons, {})
                rev_existing = rev_bucket.get(ant)
                if rev_existing is None or reverse_candidate["score"] > rev_existing["score"]:
                    rev_bucket[ant] = reverse_candidate

    for ant, cons_map in per_antecedent.items():
        rec_idx[ant] = sorted(
            cons_map.values(),
            key=lambda x: (x["score"], x["lift"], x["confidence"]),
            reverse=True,
        )

    return rec_idx


def build_rules_export(result: AnalysisResult) -> pd.DataFrame:
    if result.raw_rules is None or result.raw_rules.empty:
        return pd.DataFrame(
            columns=["antecedent", "consequent", "support", "confidence", "lift", "strategy_tag"]
        )

    rules = result.raw_rules.copy()
    rules = rules[(rules["antecedents"].apply(len) == 1) & (rules["consequents"].apply(len) == 1)].copy()

    if rules.empty:
        return pd.DataFrame(
            columns=["antecedent", "consequent", "support", "confidence", "lift", "strategy_tag"]
        )

    rules["antecedent"] = rules["antecedents"].apply(lambda s: next(iter(s)))
    rules["consequent"] = rules["consequents"].apply(lambda s: next(iter(s)))

    rules["strategy_tag"] = "cross_sell"
    rules.loc[rules["lift"] >= 2.0, "strategy_tag"] = "bundle_candidate"
    rules.loc[rules["lift"] >= 3.0, "strategy_tag"] = "strong_bundle"

    out = rules[["antecedent", "consequent", "support", "confidence", "lift", "strategy_tag"]]
    out = out.sort_values(["lift", "confidence", "support"], ascending=False).reset_index(drop=True)
    return out


def build_pair_rule_fallback_index(
    invoice_lines: pd.DataFrame,
    min_support: float = 0.0005,
    min_lift: float = 1.1,
) -> dict[str, list[dict[str, Any]]]:
    tx_map = (
        invoice_lines.groupby("receipt_id")["product_key"]
        .apply(lambda s: sorted(set([x for x in s.astype(str) if str(x).strip()])))
        .to_dict()
    )

    transactions = [items for items in tx_map.values() if len(items) >= 2]
    n_tx = len(tx_map)
    if n_tx == 0:
        return {}

    item_count: dict[str, int] = defaultdict(int)
    pair_count: dict[tuple[str, str], int] = defaultdict(int)

    for items in tx_map.values():
        unique_items = sorted(set(items))
        for a in unique_items:
            item_count[a] += 1
        for i in range(len(unique_items)):
            for j in range(i + 1, len(unique_items)):
                a = unique_items[i]
                b = unique_items[j]
                pair_count[(a, b)] += 1

    rec_idx: dict[str, dict[str, dict[str, Any]]] = defaultdict(dict)
    for (a, b), c in pair_count.items():
        if _is_excluded_product(a) or _is_excluded_product(b):
            continue
        support = c / n_tx
        if support < min_support:
            continue

        sup_a = item_count[a] / n_tx
        sup_b = item_count[b] / n_tx
        denom = sup_a * sup_b
        if denom <= 0:
            continue
        lift = support / denom
        if lift < min_lift:
            continue
        if not _is_reasonable_pair(a, b, support, lift, c):
            continue

        conf_ab = c / item_count[a] if item_count[a] else 0.0
        conf_ba = c / item_count[b] if item_count[b] else 0.0

        lift_capped = min(lift, 8.0)
        novelty_b = max(0.05, (1.0 - sup_b) ** 1.5)
        novelty_a = max(0.05, (1.0 - sup_a) ** 1.5)

        # Bo cac de xuat qua pho thong hoac confidence qua yeu.
        if sup_b >= 0.45 or conf_ab < 0.03:
            cand_ab = None
        else:
            cand_ab = {
                "product_key": b,
                "confidence": conf_ab,
                "lift": lift,
                "support": support,
                "consequent_support": sup_b,
                "score": conf_ab * lift_capped * (support ** 0.5) * novelty_b,
                "source": "association_rule_pair",
            }

        if sup_a >= 0.45 or conf_ba < 0.03:
            cand_ba = None
        else:
            cand_ba = {
                "product_key": a,
                "confidence": conf_ba,
                "lift": lift,
                "support": support,
                "consequent_support": sup_a,
                "score": conf_ba * lift_capped * (support ** 0.5) * novelty_a,
                "source": "association_rule_pair",
            }

        if cand_ab is not None:
            ex_ab = rec_idx[a].get(b)
            if ex_ab is None or cand_ab["score"] > ex_ab["score"]:
                rec_idx[a][b] = cand_ab

        if cand_ba is not None:
            ex_ba = rec_idx[b].get(a)
            if ex_ba is None or cand_ba["score"] > ex_ba["score"]:
                rec_idx[b][a] = cand_ba

    final_idx: dict[str, list[dict[str, Any]]] = {}
    for ant, rec_map in rec_idx.items():
        final_idx[ant] = sorted(
            rec_map.values(),
            key=lambda x: (x["score"], x["lift"], x["confidence"]),
            reverse=True,
        )

    return final_idx


def run_full_recommendation_pipeline(
    csv_path: str | Path = DEFAULT_DATA_FILE,
    min_support: float = 0.02,
    min_confidence: float = 0.30,
    min_lift: float = 1.20,
    max_len: int = 2,
) -> dict[str, Any]:
    invoice_lines = load_invoice_lines(csv_path)
    catalog = build_product_catalog(invoice_lines)

    analysis_result = mine_association_rules(
        csv_path=csv_path,
        min_support=min_support,
        min_confidence=min_confidence,
        min_lift=min_lift,
        max_len=max_len,
    )

    # Lop 1: quy luat theo bo nguong chinh.
    rec_idx_strict = build_recommendation_index(
        analysis_result.raw_rules,
        source_label="association_rule",
        include_reverse=False,
    )

    # Lop 2: fallback theo cap dong xuat hien (pair rules), toi uu bo nho cho 11k hoa don.
    rec_idx_relaxed = build_pair_rule_fallback_index(
        invoice_lines,
        min_support=max(0.0002, min_support / 80),
        min_lift=1.1,
    )

    rec_idx: dict[str, list[dict[str, Any]]] = {}
    all_keys = set(rec_idx_strict.keys()) | set(rec_idx_relaxed.keys())
    for key in all_keys:
        strict_list = rec_idx_strict.get(key, [])
        if strict_list:
            rec_idx[key] = strict_list
        else:
            rec_idx[key] = rec_idx_relaxed.get(key, [])

    rules_export = build_rules_export(analysis_result)

    return {
        "invoice_lines": invoice_lines,
        "catalog": catalog,
        "analysis": analysis_result,
        "recommendation_index": rec_idx,
        "rules_export": rules_export,
    }
