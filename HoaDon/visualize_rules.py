from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
import pandas as pd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Visualize association rules export file.")
    parser.add_argument("--input", default="outputs/rules/association_rules_products.csv", help="Rules CSV file")
    parser.add_argument("--orders", default="HD-csv.csv", help="Orders CSV file")
    parser.add_argument("--output-dir", default="outputs/visualize-main/rules", help="Output directory")
    parser.add_argument("--top", type=int, default=20, help="Top rules to display")
    return parser.parse_args()


def setup_style() -> None:
    plt.style.use("seaborn-v0_8-whitegrid")
    plt.rcParams["font.family"] = "DejaVu Sans"
    plt.rcParams["axes.unicode_minus"] = False


def _fmt_pct(v: float) -> str:
    return f"{v * 100:.2f}%"


def _cleanup_old_png(output_dir: Path) -> None:
    for p in output_dir.glob("*.png"):
        p.unlink(missing_ok=True)


def _norm(text: str) -> str:
    return " ".join(str(text).strip().lower().split())


def _load_transactions(orders_path: Path) -> list[set[str]]:
    if not orders_path.exists():
        return []

    df_orders = pd.read_csv(orders_path, encoding="utf-8-sig")
    if "productNames" not in df_orders.columns:
        return []

    txs: list[set[str]] = []
    for raw in df_orders["productNames"].fillna(""):
        items = {_norm(x) for x in str(raw).split("|") if _norm(x)}
        if items:
            txs.append(items)
    return txs


def _support_in_transactions(transactions: Iterable[set[str]], item: str) -> float:
    txs = list(transactions)
    if not txs:
        return 0.0
    item_n = _norm(item)
    cnt = sum(1 for t in txs if item_n in t)
    return cnt / len(txs)


def main() -> None:
    args = parse_args()
    setup_style()

    input_path = Path(args.input)
    orders_path = Path(args.orders)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    _cleanup_old_png(output_dir)

    if not input_path.exists():
        raise FileNotFoundError(f"Rules file not found: {input_path}")

    df = pd.read_csv(input_path, encoding="utf-8-sig")
    required = ["antecedent", "consequent", "support", "confidence", "lift", "strategy_tag"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in rules file: {missing}")

    transactions = _load_transactions(orders_path)

    df = df.copy()
    df["rule_label"] = df["antecedent"].astype(str) + " -> " + df["consequent"].astype(str)
    df_top = df.sort_values(["lift", "confidence"], ascending=False).head(args.top)
    top_12_lift = df_top.head(12).copy().reset_index(drop=True)
    top_12_lift["hang"] = top_12_lift.index + 1

    # 01: Dashboard tong quan (bar + pie + line)
    fig, axes = plt.subplots(1, 3, figsize=(21, 7.6), constrained_layout=True)
    fig.patch.set_facecolor("#f4f6f8")

    ax1 = axes[0]
    ax1.bar(top_12_lift["hang"].astype(str), top_12_lift["lift"], color="#2c7fb8")
    ax1.set_title("1. Top luật theo Lift", fontweight="bold")
    ax1.set_xlabel("Hạng luật")
    ax1.set_ylabel("Lift")
    ax1.grid(axis="y", alpha=0.25)

    ax2 = axes[1]
    tag_counts = df["strategy_tag"].value_counts()
    ax2.pie(tag_counts.values, labels=tag_counts.index, autopct="%1.1f%%", startangle=90)
    ax2.set_title("2. Tỷ trọng nhóm chiến lược luật", fontweight="bold")
    ax2.axis("equal")

    ax3 = axes[2]
    ax3.plot(top_12_lift["hang"], top_12_lift["lift"], marker="o", linewidth=2.2, label="Lift")
    ax3.plot(top_12_lift["hang"], top_12_lift["confidence"], marker="s", linewidth=2.2, label="Confidence")
    ax3.plot(top_12_lift["hang"], top_12_lift["support"], marker="^", linewidth=2.2, label="Support")
    ax3.set_title("3. Xu hướng chỉ số top luật", fontweight="bold")
    ax3.set_xlabel("Hạng luật")
    ax3.set_ylabel("Giá trị chỉ số")
    ax3.grid(axis="y", alpha=0.25)
    ax3.legend(loc="best")

    fig.suptitle("Dashboard Tổng quan Khai phá Luật Kết hợp", fontsize=15, fontweight="bold", y=1.02)
    fig.savefig(output_dir / "01_dashboard_tong_quan_luat_ket_hop.png", dpi=180, bbox_inches="tight")
    plt.close(fig)

    # 02: Top rules by lift (bar ngang)
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.barh(df_top["rule_label"], df_top["lift"], color="#2c7fb8")
    ax.invert_yaxis()
    ax.set_title(f"Top {args.top} Luật kết hợp theo Lift", fontweight="bold")
    ax.set_xlabel("Lift")
    ax.set_ylabel("Luật")
    ax.grid(axis="x", alpha=0.25)
    plt.tight_layout()
    fig.savefig(output_dir / "02_top_luat_theo_lift_cot_ngang.png", dpi=180, bbox_inches="tight")
    plt.close(fig)

    # 03: Top rules by confidence (bar ngang)
    df_top_conf = df.sort_values(["confidence", "lift"], ascending=False).head(args.top)
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.barh(df_top_conf["rule_label"], df_top_conf["confidence"], color="#238b45")
    ax.invert_yaxis()
    ax.set_title(f"Top {args.top} Luật kết hợp theo Confidence", fontweight="bold")
    ax.set_xlabel("Confidence")
    ax.set_ylabel("Luật")
    ax.grid(axis="x", alpha=0.25)
    plt.tight_layout()
    fig.savefig(output_dir / "03_top_luat_theo_confidence_cot_ngang.png", dpi=180, bbox_inches="tight")
    plt.close(fig)

    # 04: Ty trong strategy tag (pie)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.pie(tag_counts.values, labels=tag_counts.index, autopct="%1.1f%%", startangle=90)
    ax.set_title("Tỷ trọng luật theo nhóm chiến lược", fontweight="bold")
    ax.axis("equal")
    plt.tight_layout()
    fig.savefig(output_dir / "04_ty_trong_luat_theo_nhom_chien_luoc_tron.png", dpi=180, bbox_inches="tight")
    plt.close(fig)

    # 05: Ty trong san pham o ve trai (pie)
    fig, ax = plt.subplots(figsize=(10, 8))
    ant_counts = df["antecedent"].value_counts().head(12)
    ax.pie(ant_counts.values, labels=ant_counts.index, autopct="%1.1f%%", startangle=90, textprops={"fontsize": 9})
    ax.set_title("Tỷ trọng top sản phẩm ở vế trái luật", fontweight="bold")
    ax.axis("equal")
    plt.tight_layout()
    fig.savefig(output_dir / "05_ty_trong_top_san_pham_ve_trai_tron.png", dpi=180, bbox_inches="tight")
    plt.close(fig)

    # 06: Duong xu huong chi so top luat (line)
    df_rank = df.sort_values("lift", ascending=False).head(args.top).reset_index(drop=True)
    df_rank["rank"] = df_rank.index + 1
    fig, ax = plt.subplots(figsize=(13, 7))
    ax.plot(df_rank["rank"], df_rank["lift"], marker="o", linewidth=2, label="Lift")
    ax.plot(df_rank["rank"], df_rank["confidence"], marker="s", linewidth=2, label="Confidence")
    ax.plot(df_rank["rank"], df_rank["support"], marker="^", linewidth=2, label="Support")
    ax.set_title("Xu hướng Lift/Confidence/Support của top luật", fontweight="bold")
    ax.set_xlabel("Hạng luật (sắp theo Lift)")
    ax.set_ylabel("Giá trị chỉ số")
    ax.grid(alpha=0.25)
    ax.legend(loc="best")
    plt.tight_layout()
    fig.savefig(output_dir / "06_xu_huong_chi_so_top_luat_duong.png", dpi=180, bbox_inches="tight")
    plt.close(fig)

    # 07: Ty le support thuc te cua top antecedent trong don hang (bar)
    if transactions:
        top_ant = df["antecedent"].value_counts().head(args.top)
        ant_support = pd.Series({k: _support_in_transactions(transactions, k) for k in top_ant.index})
        ant_support = ant_support.sort_values(ascending=False).head(args.top)

        fig, ax = plt.subplots(figsize=(14, 8))
        ax.barh(ant_support.index, ant_support.values, color="#6a51a3")
        ax.invert_yaxis()
        ax.set_title("Tỷ lệ support thực tế của top sản phẩm vế trái", fontweight="bold")
        ax.set_xlabel("Tỷ lệ support trên đơn hàng thật")
        ax.set_ylabel("Sản phẩm vế trái")
        ax.grid(axis="x", alpha=0.25)
        plt.tight_layout()
        fig.savefig(output_dir / "07_support_thuc_te_top_san_pham_ve_trai_cot_ngang.png", dpi=180, bbox_inches="tight")
        plt.close(fig)

    # 08: Top cap san pham mua cung theo support (bar)
    top_pair = (
        df.groupby(["antecedent", "consequent"], as_index=False)
        .agg(support=("support", "max"), confidence=("confidence", "max"), lift=("lift", "max"))
        .sort_values(["support", "lift"], ascending=False)
        .head(args.top)
    )
    top_pair["cap_label"] = top_pair["antecedent"].astype(str) + " + " + top_pair["consequent"].astype(str)
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.barh(top_pair["cap_label"], top_pair["support"], color="#dd6b20")
    ax.invert_yaxis()
    ax.set_title("Top cặp sản phẩm mua cùng theo Support", fontweight="bold")
    ax.set_xlabel("Support")
    ax.set_ylabel("Cặp sản phẩm")
    ax.grid(axis="x", alpha=0.25)
    plt.tight_layout()
    fig.savefig(output_dir / "08_top_cap_san_pham_mua_cung_theo_support_cot_ngang.png", dpi=180, bbox_inches="tight")
    plt.close(fig)

    summary = df[["support", "confidence", "lift"]].describe().T
    summary.to_csv(output_dir / "rules_summary_stats.csv", encoding="utf-8-sig")

    print(f"Visualized rules file: {input_path}")
    print(f"Output folder: {output_dir.resolve()}")
    for p in sorted(output_dir.glob("*.png")):
        print(f"- {p.name}")
    print("- rules_summary_stats.csv")


if __name__ == "__main__":
    main()
