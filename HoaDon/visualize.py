import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def parse_args():
    parser = argparse.ArgumentParser(
        description="Vẽ dashboard thống kê hóa đơn chuyên nghiệp theo phong cách báo cáo."
    )
    parser.add_argument(
        "--input",
        default="HoaDon_DaMaHoa.csv",
        help="File hóa đơn đã mã hóa (mặc định: HoaDon_DaMaHoa.csv).",
    )
    parser.add_argument(
        "--output-dir",
        default="visuals_chuyen_nghiep",
        help="Thư mục lưu ảnh biểu đồ.",
    )
    parser.add_argument(
        "--top-brands",
        type=int,
        default=8,
        help="Số thương hiệu top hiển thị trong biểu đồ cột ngang.",
    )
    return parser.parse_args()


def setup_style():
    plt.style.use("seaborn-v0_8-whitegrid")
    plt.rcParams["font.family"] = "DejaVu Sans"
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["axes.titlesize"] = 13
    plt.rcParams["axes.labelsize"] = 11
    plt.rcParams["xtick.labelsize"] = 9
    plt.rcParams["ytick.labelsize"] = 9


def normalize_platform_label(raw):
    p = str(raw).strip().lower()
    if "bhx" in p or "bach hoa xanh" in p:
        return "BHX"
    if "lotte" in p:
        return "Lotte"
    if "tiki" in p:
        return "Tiki"
    if "winmart" in p:
        return "WinMart"
    return str(raw)


def load_data(input_path):
    df = pd.read_csv(input_path, encoding="utf-8-sig")

    required = ["receiptId", "receiptDateTime", "finalAmount"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Thiếu cột bắt buộc: {missing}")

    if "platform_text" in df.columns:
        df["display_platform"] = df["platform_text"].astype(str).map(normalize_platform_label)
    elif "platform" in df.columns:
        df["display_platform"] = df["platform"].astype(str).map(normalize_platform_label)
    else:
        raise ValueError("Không tìm thấy cột platform/platform_text")

    if "brand_text" in df.columns:
        df["display_brand"] = df["brand_text"].fillna("unknown").astype(str)
    elif "brand" in df.columns:
        df["display_brand"] = df["brand"].fillna("unknown").astype(str)
    elif "brands" in df.columns:
        df["display_brand"] = df["brands"].fillna("unknown").astype(str)
    else:
        df["display_brand"] = "unknown"

    if "quantity" not in df.columns:
        if "quantities" in df.columns:
            def sum_quantities(raw):
                parts = [p.strip() for p in str(raw).split("|") if p.strip()]
                total = 0
                for p in parts:
                    try:
                        total += int(float(p))
                    except Exception:
                        continue
                return total

            df["quantity"] = df["quantities"].apply(sum_quantities)
        elif "itemCount" in df.columns:
            df["quantity"] = pd.to_numeric(df["itemCount"], errors="coerce")
        else:
            df["quantity"] = 1

    for c in ["quantity", "finalAmount", "itemCount"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df["receiptDateTime"] = pd.to_datetime(
        df["receiptDateTime"],
        format="%d/%m/%Y %H:%M",
        errors="coerce",
    )

    df = df.dropna(subset=["receiptId", "display_platform", "finalAmount", "quantity"])
    return df


def build_invoice_level(df):
    use_item_count = "itemCount" in df.columns and df["itemCount"].notna().any()
    if use_item_count:
        invoice_level = (
            df.sort_values("receiptDateTime")
            .groupby("receiptId", as_index=False)
            .agg(
                display_platform=("display_platform", "first"),
                receiptDateTime=("receiptDateTime", "first"),
                finalAmount=("finalAmount", "first"),
                so_luong_mat_hang=("itemCount", "first"),
                tong_so_luong=("quantity", "first"),
            )
        )
        return invoice_level

    invoice_level = (
        df.sort_values("receiptDateTime")
        .groupby("receiptId", as_index=False)
        .agg(
            display_platform=("display_platform", "first"),
            receiptDateTime=("receiptDateTime", "first"),
            finalAmount=("finalAmount", "first"),
            so_luong_mat_hang=("receiptId", "size"),
            tong_so_luong=("quantity", "sum"),
        )
    )
    return invoice_level


def add_note(ax, text, x=0.02, y=0.02, fontsize=8):
    ax.text(
        x,
        y,
        text,
        transform=ax.transAxes,
        va="bottom",
        ha="left",
        fontsize=fontsize,
        bbox=dict(boxstyle="round,pad=0.28", facecolor="#f8f9fa", edgecolor="#d0d7de", alpha=0.9),
    )


def annotate_bars(ax, fontsize=8):
    for bar in ax.patches:
        height = bar.get_height()
        ax.annotate(
            f"{int(round(height)):,}".replace(",", "."),
            (bar.get_x() + bar.get_width() / 2, height),
            ha="center",
            va="bottom",
            fontsize=fontsize,
            xytext=(0, 3),
            textcoords="offset points",
        )


def create_dashboard_tong_quan(invoice_level, output_file):
    fig, axes = plt.subplots(1, 3, figsize=(21, 7.8), constrained_layout=True)
    fig.patch.set_facecolor("#f4f6f8")

    platform_order = ["BHX", "WinMart", "Tiki", "Lotte", "WinMart (Merged)"]

    # 1) Quota hóa đơn
    quota = invoice_level.groupby("display_platform").size().reindex(platform_order).dropna()
    ax1 = axes[0]
    colors = ["#2ca25f", "#de2d26", "#3182bd", "#d94873", "#7b6fd0"][: len(quota)]
    quota.plot(kind="bar", ax=ax1, color=colors)
    ax1.set_title("1. Phân bổ Quota Hóa đơn (11k đơn)", fontweight="bold")
    ax1.set_xlabel("Nền tảng")
    ax1.set_ylabel("Số lượng hóa đơn")
    ax1.grid(axis="y", alpha=0.25)
    ax1.tick_params(axis="x", rotation=10)
    ax1.set_ylim(0, quota.max() * 1.2)
    annotate_bars(ax1, fontsize=9)
    add_note(ax1, "Đơn vị: hóa đơn", y=0.03)

    # 2) Phân bố kích thước giỏ hàng
    basket = invoice_level["so_luong_mat_hang"].value_counts().sort_index()
    ax2 = axes[1]
    ax2.bar(basket.index.astype(str), basket.values, color="#5e3c99")
    ax2.set_title("2. Phân bố Kích thước Giỏ hàng", fontweight="bold")
    ax2.set_xlabel("Số lượng mặt hàng/Hóa đơn")
    ax2.set_ylabel("Tần suất xuất hiện")
    ax2.grid(axis="y", alpha=0.25)
    add_note(ax2, "Đơn vị trục X: mặt hàng | trục Y: số hóa đơn", y=0.03)

    # 3) Phân bố giá trị thanh toán
    ax3 = axes[2]
    values = invoice_level["finalAmount"]
    bin_count = 28
    categories, bins = pd.cut(values, bins=bin_count, retbins=True)
    counts = categories.value_counts().sort_index().values
    centers = (bins[:-1] + bins[1:]) / 2
    widths = bins[1:] - bins[:-1]

    ax3.bar(centers, counts, width=widths * 0.92, color="#a6cee3", edgecolor="#5aa0af", alpha=0.8)
    smooth = pd.Series(counts).rolling(3, min_periods=1, center=True).mean()
    ax3.plot(centers, smooth, color="#2c7fb8", linewidth=2.0)
    ax3.set_title("3. Phân bố Giá trị Thanh toán (VND)", fontweight="bold")
    ax3.set_xlabel("Số tiền (Final Amount)")
    ax3.set_ylabel("Mật độ đơn hàng")
    ax3.grid(axis="y", alpha=0.25)
    ax3.tick_params(axis="x", rotation=25)
    add_note(ax3, "Đơn vị trục X: VND | trục Y: số đơn theo khoảng giá", y=0.03)

    fig.suptitle("Dashboard Tổng quan Chất lượng Dữ liệu Hóa đơn", fontsize=15, fontweight="bold", y=1.03)
    fig.savefig(output_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def create_platform_distribution(invoice_level, output_file):
    summary = invoice_level.groupby("display_platform").size().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(13.5, 7.2))
    fig.patch.set_facecolor("#f4f6f8")
    ax.bar(summary.index, summary.values, color=["#58508d", "#3e7c8e", "#3b9c7d", "#7fbc5b", "#e45756"][: len(summary)])
    ax.set_title("1. Phân phối Nguồn dữ liệu (Platform Distribution)", fontweight="bold", fontsize=13)
    ax.set_xlabel("Nền tảng")
    ax.set_ylabel("Số lượng hóa đơn")
    ax.tick_params(axis="x", rotation=12)
    ax.set_ylim(0, summary.max() * 1.18)
    ax.grid(axis="y", alpha=0.25)
    annotate_bars(ax, fontsize=9)
    add_note(ax, "Đơn vị: hóa đơn | So sánh mức đóng góp dữ liệu theo sàn", y=0.03)
    plt.tight_layout()
    fig.savefig(output_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def create_top_brand_barh(df, output_file, top_n):
    if "brands" in df.columns:
        brand_tokens = (
            df["brands"]
            .fillna("")
            .astype(str)
            .str.split("|", regex=False)
            .explode()
            .astype(str)
            .str.strip()
        )
        brand_tokens = brand_tokens[brand_tokens != ""]
        top_brand = (
            brand_tokens.to_frame(name="display_brand")
            .assign(so_dong_item=1)
            .groupby("display_brand", as_index=False)["so_dong_item"]
            .sum()
            .sort_values("so_dong_item", ascending=False)
            .head(top_n)
        )
    else:
        top_brand = (
            df.groupby("display_brand", as_index=False)
            .agg(so_dong_item=("receiptId", "size"))
            .sort_values("so_dong_item", ascending=False)
            .head(top_n)
        )

    fig, ax = plt.subplots(figsize=(12.5, 7.2))
    fig.patch.set_facecolor("#f4f6f8")
    ax.barh(top_brand["display_brand"], top_brand["so_dong_item"], color="#f28e2b")
    ax.invert_yaxis()
    ax.set_title(f"Top {top_n} Thương hiệu xuất hiện nhiều nhất", fontweight="bold")
    ax.set_xlabel("Số dòng item")
    ax.set_ylabel("Thương hiệu")
    ax.grid(axis="x", alpha=0.25)
    ax.set_xlim(0, top_brand["so_dong_item"].max() * 1.22)

    for bar in ax.patches:
        w = bar.get_width()
        ax.text(w + max(1, w * 0.005), bar.get_y() + bar.get_height() / 2, f"{int(w):,}".replace(",", "."), va="center", fontsize=9)

    add_note(ax, "Đơn vị: dòng item | Dùng để phát hiện thương hiệu áp đảo", y=0.03)
    plt.tight_layout()
    fig.savefig(output_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def create_platform_pie(invoice_level, output_file):
    summary = invoice_level.groupby("display_platform").size().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(10.5, 8.8))
    fig.patch.set_facecolor("#f4f6f8")
    wedges, texts, autotexts = ax.pie(
        summary.values,
        labels=summary.index,
        autopct="%1.1f%%",
        startangle=90,
        pctdistance=0.8,
        textprops={"fontsize": 10},
    )
    ax.set_title("Tỷ trọng số lượng hóa đơn theo sàn", fontweight="bold")
    ax.axis("equal")
    ax.legend(
        wedges,
        [f"{lab}: {int(val):,} hóa đơn".replace(",", ".") for lab, val in zip(summary.index, summary.values)],
        title="Chú thích",
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
    )
    add_note(ax, f"Đơn vị: % và hóa đơn | Tổng: {int(summary.sum()):,} hóa đơn".replace(",", "."), x=-0.02, y=-0.06)
    plt.tight_layout()
    fig.savefig(output_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def create_monthly_line(invoice_level, output_file):
    trend = invoice_level.dropna(subset=["receiptDateTime"]).copy()
    trend["thang"] = trend["receiptDateTime"].dt.to_period("M").astype(str)
    pivot = trend.pivot_table(
        index="thang",
        columns="display_platform",
        values="receiptId",
        aggfunc="count",
        fill_value=0,
    )

    fig, ax = plt.subplots(figsize=(13.5, 6.6))
    fig.patch.set_facecolor("#f4f6f8")
    pivot.plot(ax=ax, marker="o", linewidth=2)
    ax.set_title("Xu hướng số lượng hóa đơn theo tháng", fontweight="bold")
    ax.set_xlabel("Tháng")
    ax.set_ylabel("Số hóa đơn")
    ax.grid(axis="y", alpha=0.25)
    plt.xticks(rotation=45, ha="right")
    add_note(ax, "Đơn vị: hóa đơn | Mỗi đường biểu diễn một nền tảng", y=0.03)
    ax.legend(title="Nền tảng", loc="upper left", bbox_to_anchor=(1.01, 1.0))
    plt.tight_layout()
    fig.savefig(output_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def write_summary(invoice_level, output_dir):
    summary = invoice_level.groupby("display_platform", as_index=False).agg(
        so_hoa_don=("receiptId", "size"),
        gia_tri_hoa_don_trung_binh=("finalAmount", "mean"),
        tong_thanh_toan=("finalAmount", "sum"),
        so_mat_hang_trung_binh=("so_luong_mat_hang", "mean"),
        tong_so_luong_trung_binh=("tong_so_luong", "mean"),
    )
    summary = summary.rename(columns={"display_platform": "nen_tang"}).sort_values("so_hoa_don", ascending=False)
    summary.to_csv(output_dir / "tổng_hợp_thống_kê_hóa_đơn_theo_sàn.csv", index=False, encoding="utf-8-sig")


def main():
    args = parse_args()
    setup_style()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file đầu vào: {input_path}")

    df = load_data(input_path)
    invoice_level = build_invoice_level(df)
    write_summary(invoice_level, output_dir)

    create_dashboard_tong_quan(invoice_level, output_dir / "01_dashboard_tổng_quan_hóa_đơn.png")
    create_platform_distribution(invoice_level, output_dir / "02_phân_phối_nguồn_dữ_liệu_theo_sàn.png")
    create_monthly_line(invoice_level, output_dir / "03_xu_hướng_hóa_đơn_theo_tháng_biểu_đồ_đường.png")
    create_platform_pie(invoice_level, output_dir / "04_tỷ_trọng_hóa_đơn_theo_sàn_biểu_đồ_tròn.png")
    create_top_brand_barh(df, output_dir / "05_top_thương_hiệu_nhiều_nhất_cột_ngang.png", args.top_brands)

    print("Đã tạo xong bộ biểu đồ chuyên nghiệp.")
    print(f"Nguồn dữ liệu: {input_path}")
    print(f"Thư mục output: {output_dir.resolve()}")
    print("Các file đã tạo:")
    for file in sorted(output_dir.glob("*.png")):
        print(f"- {file.name}")
    print("- tổng_hợp_thống_kê_hóa_đơn_theo_sàn.csv")


if __name__ == "__main__":
    main()
