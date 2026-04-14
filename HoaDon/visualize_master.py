import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def parse_args():
    parser = argparse.ArgumentParser(description="Visualize thống kê file master data sản phẩm theo sàn.")
    parser.add_argument(
        "--input",
        default="Final_Master_Data_v3.backup_before_augment.csv",
        help="Đường dẫn file master data cần thống kê.",
    )
    parser.add_argument(
        "--output-dir",
        default="visuals_master_backup",
        help="Thư mục lưu biểu đồ.",
    )
    parser.add_argument(
        "--top-brands",
        type=int,
        default=12,
        help="Số thương hiệu top hiển thị trong biểu đồ cột ngang.",
    )
    return parser.parse_args()


def setup_style():
    plt.style.use("seaborn-v0_8-whitegrid")
    plt.rcParams["font.family"] = "DejaVu Sans"
    plt.rcParams["axes.unicode_minus"] = False


def normalize_platform(raw):
    p = str(raw or "").strip().lower()
    if "bhx" in p:
        return "BHX"
    if "tiki" in p:
        return "Tiki"
    if "lotte" in p:
        return "Lotte Mart"
    if "winmart" in p:
        return "WinMart"
    return "Khác"


def add_note(ax, text):
    ax.text(
        0.02,
        0.02,
        text,
        transform=ax.transAxes,
        va="bottom",
        ha="left",
        fontsize=8,
        bbox=dict(boxstyle="round,pad=0.3", facecolor="#f8f9fa", edgecolor="#d0d7de"),
    )


def load_master(input_path):
    df = pd.read_csv(input_path, encoding="utf-8-sig")
    required = ["product_id", "platform", "brand", "sale_price", "category"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Thiếu cột bắt buộc trong master: {missing}")

    df["platform_group"] = df["platform"].apply(normalize_platform)
    df["brand_clean"] = df["brand"].fillna("unknown").astype(str).str.strip().str.lower()
    df["sale_price"] = pd.to_numeric(df["sale_price"], errors="coerce")
    df = df.dropna(subset=["product_id", "platform_group", "sale_price"])
    df = df[df["platform_group"] != "Khác"].copy()
    return df


def build_summary(df):
    base = (
        df.groupby("platform_group", as_index=False)
        .agg(
            so_mat_hang=("product_id", "nunique"),
            gia_ban_trung_binh=("sale_price", "mean"),
            gia_ban_trung_vi=("sale_price", "median"),
        )
        .sort_values("so_mat_hang", ascending=False)
    )
    return base


def plot_dashboard(summary, df, out_file):
    fig, axes = plt.subplots(1, 3, figsize=(21, 7.5), constrained_layout=True)
    fig.patch.set_facecolor("#f4f6f8")

    # 1) Cột dọc: số lượng mặt hàng theo sàn
    ax1 = axes[0]
    bars = ax1.bar(summary["platform_group"], summary["so_mat_hang"], color=["#2ca25f", "#3182bd", "#d94873", "#7b6fd0"])
    ax1.set_title("1. Số lượng mặt hàng theo sàn", fontweight="bold")
    ax1.set_xlabel("Sàn")
    ax1.set_ylabel("Số mặt hàng khác nhau")
    ax1.set_ylim(0, summary["so_mat_hang"].max() * 1.2)
    ax1.grid(axis="y", alpha=0.25)
    for b in bars:
        h = b.get_height()
        ax1.text(b.get_x() + b.get_width()/2, h, f"{int(h):,}".replace(",", "."), ha="center", va="bottom", fontsize=9)
    add_note(ax1, "Đơn vị: mặt hàng")

    # 2) Cột ngang: top thương hiệu
    ax2 = axes[1]
    top_brand = (
        df.groupby("brand_clean", as_index=False)
        .agg(so_mat_hang=("product_id", "nunique"))
        .sort_values("so_mat_hang", ascending=False)
        .head(10)
    )
    ax2.barh(top_brand["brand_clean"], top_brand["so_mat_hang"], color="#f28e2b")
    ax2.set_title("2. Top 10 thương hiệu theo số mặt hàng", fontweight="bold")
    ax2.set_xlabel("Số mặt hàng")
    ax2.set_ylabel("Thương hiệu")
    ax2.invert_yaxis()
    ax2.grid(axis="x", alpha=0.25)
    add_note(ax2, "Đơn vị: mặt hàng")

    # 3) Đường: mức giá trung bình và trung vị
    ax3 = axes[2]
    s = summary.set_index("platform_group")
    ax3.plot(s.index, s["gia_ban_trung_binh"], marker="o", linewidth=2.2, label="Giá bán trung bình")
    ax3.plot(s.index, s["gia_ban_trung_vi"], marker="s", linewidth=2.2, label="Giá bán trung vị")
    ax3.set_title("3. So sánh mức giá theo sàn", fontweight="bold")
    ax3.set_xlabel("Sàn")
    ax3.set_ylabel("Giá (VND)")
    ax3.grid(axis="y", alpha=0.25)
    ax3.legend(loc="upper left")
    add_note(ax3, "Đơn vị: VND")

    fig.suptitle("Dashboard Tổng quan Master Data (File Backup)", fontsize=15, fontweight="bold", y=1.02)
    fig.savefig(out_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def plot_pie(summary, out_file):
    fig, ax = plt.subplots(figsize=(10.5, 8.5))
    fig.patch.set_facecolor("#f4f6f8")

    wedges, _, _ = ax.pie(
        summary["so_mat_hang"],
        labels=summary["platform_group"],
        autopct="%1.1f%%",
        startangle=90,
        textprops={"fontsize": 10},
    )
    ax.set_title("Tỷ trọng số lượng mặt hàng theo sàn", fontweight="bold")
    ax.axis("equal")
    ax.legend(
        wedges,
        [f"{p}: {int(v):,} mặt hàng".replace(",", ".") for p, v in zip(summary["platform_group"], summary["so_mat_hang"])],
        title="Chú thích",
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
    )
    add_note(ax, "Đơn vị: % và mặt hàng")
    fig.savefig(out_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def write_summary(summary, output_dir):
    summary_out = summary.rename(
        columns={
            "platform_group": "san",
            "so_mat_hang": "so_mat_hang_khac_nhau",
            "gia_ban_trung_binh": "gia_ban_trung_binh_vnd",
            "gia_ban_trung_vi": "gia_ban_trung_vi_vnd",
        }
    )
    summary_out.to_csv(output_dir / "tong_hop_master_data_theo_san.csv", index=False, encoding="utf-8-sig")


def main():
    args = parse_args()
    setup_style()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {input_path}")

    df = load_master(input_path)
    summary = build_summary(df)

    write_summary(summary, output_dir)
    plot_dashboard(summary, df, output_dir / "01_dashboard_master_data_backup.png")
    plot_pie(summary, output_dir / "02_ty_trong_mat_hang_theo_san_bieu_do_tron.png")

    print("Đã visualize xong file master backup.")
    print(f"Input: {input_path}")
    print(f"Output: {output_dir.resolve()}")
    for p in sorted(output_dir.glob("*.png")):
        print(f"- {p.name}")
    print("- tong_hop_master_data_theo_san.csv")


if __name__ == "__main__":
    main()
