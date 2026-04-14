from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
from mlxtend.frequent_patterns import apriori, association_rules
from mlxtend.preprocessing import TransactionEncoder


@dataclass
class AnalysisResult:
    summary: dict[str, Any]
    frequent_itemsets: pd.DataFrame
    raw_rules: pd.DataFrame
    rules: pd.DataFrame
    cross_sell: pd.DataFrame
    bundles: pd.DataFrame
    shelf_convenience: pd.DataFrame
    shelf_stimulation: pd.DataFrame


def _clean_item(text: str) -> str:
    item = str(text).strip().lower()
    item = " ".join(item.split())
    return item


def _is_excluded_item(item: str) -> bool:
    s = _clean_item(item)
    return "san pham bo sung" in s


def load_transactions(csv_path: str | Path, product_col: str = "productNames", sep: str = "|") -> list[list[str]]:
    path = Path(csv_path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    df = pd.read_csv(path, encoding="utf-8-sig")
    if product_col not in df.columns:
        raise ValueError(f"Missing column '{product_col}' in {path.name}")

    transactions: list[list[str]] = []
    for raw in df[product_col].fillna(""):
        items = [_clean_item(x) for x in str(raw).split(sep)]
        items = [x for x in items if x and not _is_excluded_item(x)]
        # Keep unique items per order to avoid duplicated support inflation.
        items = sorted(set(items))
        if items:
            transactions.append(items)

    if not transactions:
        raise ValueError("No valid transactions found after preprocessing")

    return transactions


def _to_one_hot(transactions: list[list[str]]) -> pd.DataFrame:
    te = TransactionEncoder()
    sparse_arr = te.fit(transactions).transform(transactions, sparse=True)
    return pd.DataFrame.sparse.from_spmatrix(sparse_arr, columns=te.columns_)


def _format_itemset(itemset: frozenset[str]) -> str:
    return ", ".join(sorted(itemset))


def _singleton_support_map(frequent_itemsets: pd.DataFrame) -> dict[str, float]:
    one_item = frequent_itemsets[frequent_itemsets["itemsets"].apply(lambda s: len(s) == 1)].copy()
    mapping: dict[str, float] = {}
    for _, row in one_item.iterrows():
        item = next(iter(row["itemsets"]))
        mapping[item] = float(row["support"])
    return mapping


def _prepare_rules_table(rules: pd.DataFrame) -> pd.DataFrame:
    out = rules.copy()
    out["antecedents_str"] = out["antecedents"].apply(_format_itemset)
    out["consequents_str"] = out["consequents"].apply(_format_itemset)
    cols = [
        "antecedents_str",
        "consequents_str",
        "support",
        "confidence",
        "lift",
        "leverage",
        "conviction",
    ]
    out = out[cols].sort_values(["lift", "confidence", "support"], ascending=False)
    return out


def mine_association_rules(
    csv_path: str | Path,
    min_support: float = 0.01,
    min_confidence: float = 0.30,
    min_lift: float = 1.20,
    max_len: int = 3,
) -> AnalysisResult:
    transactions = load_transactions(csv_path)
    one_hot = _to_one_hot(transactions)

    frequent_itemsets = apriori(
        one_hot,
        min_support=min_support,
        use_colnames=True,
        max_len=max_len,
    )

    if frequent_itemsets.empty:
        raise ValueError("No frequent itemsets found. Try lowering min_support.")

    rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=min_confidence)
    rules = rules[rules["lift"] >= min_lift].copy()

    if rules.empty:
        prepared_rules = pd.DataFrame(
            columns=[
                "antecedents_str",
                "consequents_str",
                "support",
                "confidence",
                "lift",
                "leverage",
                "conviction",
            ]
        )
    else:
        prepared_rules = _prepare_rules_table(rules)

    singleton_support = _singleton_support_map(frequent_itemsets)

    # Cross-sell: prioritize high confidence.
    cross_sell = prepared_rules.sort_values(["confidence", "lift"], ascending=False).head(10)

    # Bundling: pair rules with high lift where one side has relatively low support.
    pair_rules = rules[
        (rules["antecedents"].apply(len) == 1) & (rules["consequents"].apply(len) == 1)
    ].copy()

    def low_support_flag(row: pd.Series) -> bool:
        a = next(iter(row["antecedents"]))
        c = next(iter(row["consequents"]))
        a_sup = singleton_support.get(a, 1.0)
        c_sup = singleton_support.get(c, 1.0)
        return min(a_sup, c_sup) <= 0.05

    if not pair_rules.empty:
        pair_rules["has_low_support_item"] = pair_rules.apply(low_support_flag, axis=1)
        pair_rules["antecedents_str"] = pair_rules["antecedents"].apply(_format_itemset)
        pair_rules["consequents_str"] = pair_rules["consequents"].apply(_format_itemset)
        bundles = pair_rules[pair_rules["has_low_support_item"]].sort_values(
            ["lift", "confidence"], ascending=False
        )[["antecedents_str", "consequents_str", "support", "confidence", "lift"]].head(10)
    else:
        bundles = pd.DataFrame(columns=["antecedents_str", "consequents_str", "support", "confidence", "lift"])

    # Shelf convenience: both items popular enough + high lift.
    if not pair_rules.empty:
        def both_popular(row: pd.Series) -> bool:
            a = next(iter(row["antecedents"]))
            c = next(iter(row["consequents"]))
            return singleton_support.get(a, 0.0) >= 0.05 and singleton_support.get(c, 0.0) >= 0.05

        pair_rules["both_popular"] = pair_rules.apply(both_popular, axis=1)
        shelf_convenience = pair_rules[pair_rules["both_popular"]].sort_values(
            ["lift", "confidence"], ascending=False
        )[["antecedents_str", "consequents_str", "support", "confidence", "lift"]].head(10)

        shelf_stimulation = pair_rules[~pair_rules["both_popular"]].sort_values(
            ["lift", "confidence"], ascending=False
        )[["antecedents_str", "consequents_str", "support", "confidence", "lift"]].head(10)
    else:
        shelf_convenience = pd.DataFrame(columns=["antecedents_str", "consequents_str", "support", "confidence", "lift"])
        shelf_stimulation = pd.DataFrame(columns=["antecedents_str", "consequents_str", "support", "confidence", "lift"])

    summary = {
        "n_transactions": len(transactions),
        "n_unique_items": int(one_hot.shape[1]),
        "n_frequent_itemsets": int(len(frequent_itemsets)),
        "n_rules": int(len(prepared_rules)),
        "min_support": min_support,
        "min_confidence": min_confidence,
        "min_lift": min_lift,
        "max_len": max_len,
    }

    return AnalysisResult(
        summary=summary,
        frequent_itemsets=frequent_itemsets.sort_values("support", ascending=False).copy(),
        raw_rules=rules.copy(),
        rules=prepared_rules,
        cross_sell=cross_sell,
        bundles=bundles,
        shelf_convenience=shelf_convenience,
        shelf_stimulation=shelf_stimulation,
    )
