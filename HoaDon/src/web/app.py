from __future__ import annotations

from pathlib import Path
from typing import Any

from flask import Flask, redirect, render_template, request, url_for

from analysis.recommender_web_data import DEFAULT_DATA_FILE, run_full_recommendation_pipeline


BASE_DIR = Path(__file__).resolve().parents[2]
app = Flask(
    __name__,
    template_folder=str(Path(__file__).resolve().parent / "templates"),
    static_folder=str(Path(__file__).resolve().parent / "static"),
)
RULES_EXPORT_PATH = BASE_DIR / "outputs" / "rules" / "association_rules_products.csv"


APP_STATE: dict[str, Any] = {
    "params": {
        "data_file": DEFAULT_DATA_FILE,
        "min_support": 0.01,
        "min_confidence": 0.30,
        "min_lift": 1.20,
        "max_len": 3,
        "max_rules": 50,
    },
    "catalog": None,
    "analysis": None,
    "recommendation_index": {},
    "rules_export": None,
    "error": None,
}


def _safe_float(value: str, default: float) -> float:
    try:
        return float(value)
    except Exception:
        return default


def _safe_int(value: str, default: int) -> int:
    try:
        return int(value)
    except Exception:
        return default


def _table_html(df, max_rows: int = 30) -> str:
    if df is None or df.empty:
        return "<p class='empty'>No data to display.</p>"
    return (
        df.head(max_rows)
        .round(4)
        .to_html(index=False, classes="result-table", border=0)
    )


def _refresh_state(params: dict[str, Any]) -> None:
    csv_path = Path(params["data_file"])
    if not csv_path.exists():
        raise FileNotFoundError(f"Data file not found: {csv_path}")

    result = run_full_recommendation_pipeline(
        csv_path=csv_path,
        min_support=params["min_support"],
        min_confidence=params["min_confidence"],
        min_lift=params["min_lift"],
        max_len=params["max_len"],
    )

    APP_STATE["params"] = params
    APP_STATE["catalog"] = result["catalog"]
    APP_STATE["analysis"] = result["analysis"]
    APP_STATE["recommendation_index"] = result["recommendation_index"]
    APP_STATE["rules_export"] = result["rules_export"]
    APP_STATE["error"] = None

    if APP_STATE["rules_export"] is not None:
        RULES_EXPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
        APP_STATE["rules_export"].to_csv(RULES_EXPORT_PATH, index=False, encoding="utf-8-sig")


@app.route("/", methods=["GET", "POST"])
def index():
    params = dict(APP_STATE["params"])

    context = {
        "params": params,
        "summary": None,
        "rules_table": None,
        "cross_sell_table": None,
        "bundle_table": None,
        "shelf_convenience_table": None,
        "shelf_stimulation_table": None,
        "error": APP_STATE.get("error"),
    }

    if request.method == "POST":
        params["data_file"] = request.form.get("data_file", DEFAULT_DATA_FILE).strip() or DEFAULT_DATA_FILE
        params["min_support"] = _safe_float(request.form.get("min_support", "0.01"), 0.01)
        params["min_confidence"] = _safe_float(request.form.get("min_confidence", "0.30"), 0.30)
        params["min_lift"] = _safe_float(request.form.get("min_lift", "1.20"), 1.20)
        params["max_len"] = _safe_int(request.form.get("max_len", "3"), 3)
        params["max_rules"] = _safe_int(request.form.get("max_rules", "50"), 50)

        try:
            _refresh_state(params)
            return redirect(url_for("products"))
        except Exception as exc:
            APP_STATE["error"] = str(exc)
            context["error"] = str(exc)

    analysis = APP_STATE.get("analysis")
    if analysis is not None:
        context["summary"] = analysis.summary
        context["rules_table"] = _table_html(analysis.rules, params["max_rules"])
        context["cross_sell_table"] = _table_html(analysis.cross_sell, 10)
        context["bundle_table"] = _table_html(analysis.bundles, 10)
        context["shelf_convenience_table"] = _table_html(analysis.shelf_convenience, 10)
        context["shelf_stimulation_table"] = _table_html(analysis.shelf_stimulation, 10)

    context["params"] = params
    return render_template("index.html", **context)


@app.route("/products", methods=["GET"])
def products():
    if APP_STATE.get("catalog") is None:
        try:
            _refresh_state(dict(APP_STATE["params"]))
        except Exception as exc:
            APP_STATE["error"] = str(exc)

    catalog = APP_STATE.get("catalog")
    error = APP_STATE.get("error")
    q = request.args.get("q", "").strip().lower()

    products_data: list[dict[str, Any]] = []
    if catalog is not None and not catalog.empty:
        view = catalog.copy()
        if q:
            view = view[view["product_name"].astype(str).str.lower().str.contains(q, na=False)]
        view = view.head(200)
        products_data = view.to_dict(orient="records")

    return render_template(
        "products.html",
        products=products_data,
        q=q,
        error=error,
        rules_export_file=str(RULES_EXPORT_PATH),
    )


def _build_product_detail_context(catalog, product_row, error):
    product = product_row.to_dict()
    rec_idx = APP_STATE.get("recommendation_index", {})
    rec_items = rec_idx.get(product["product_key"], [])[:10]

    recommendations: list[dict[str, Any]] = []
    for rec in rec_items:
        rec_key = rec["product_key"]
        rec_match = catalog[catalog["product_key"] == rec_key]
        if rec_match.empty:
            continue
        rec_product = rec_match.iloc[0].to_dict()
        rec_product.update(
            {
                "confidence": rec["confidence"],
                "lift": rec["lift"],
                "support": rec["support"],
                "consequent_support": rec.get("consequent_support", 0.0),
                "score": rec["score"],
                "source": rec.get("source", "association_rule"),
            }
        )
        recommendations.append(rec_product)

    return render_template(
        "product_detail.html",
        product=product,
        recommendations=recommendations,
        error=error,
    )


@app.route("/product/id/<int:product_id>", methods=["GET"])
def product_detail_by_id(product_id: int):
    if APP_STATE.get("catalog") is None:
        try:
            _refresh_state(dict(APP_STATE["params"]))
        except Exception as exc:
            APP_STATE["error"] = str(exc)

    catalog = APP_STATE.get("catalog")
    error = APP_STATE.get("error")
    if catalog is None or catalog.empty:
        return render_template("product_detail.html", product=None, recommendations=[], error=error)

    match = catalog[catalog["product_id"] == product_id]
    if match.empty:
        return render_template("product_detail.html", product=None, recommendations=[], error="Không tìm thấy sản phẩm")

    return _build_product_detail_context(catalog, match.iloc[0], error)


@app.route("/product/<path:product_key>", methods=["GET"])
def product_detail(product_key: str):
    if APP_STATE.get("catalog") is None:
        try:
            _refresh_state(dict(APP_STATE["params"]))
        except Exception as exc:
            APP_STATE["error"] = str(exc)

    catalog = APP_STATE.get("catalog")
    error = APP_STATE.get("error")
    if catalog is None or catalog.empty:
        return render_template("product_detail.html", product=None, recommendations=[], error=error)

    match = catalog[catalog["product_key"] == product_key]
    if match.empty:
        return render_template("product_detail.html", product=None, recommendations=[], error="Không tìm thấy sản phẩm")

    return _build_product_detail_context(catalog, match.iloc[0], error)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
