"""Lab 1: multivariate linear regression for gas-turbine CO emissions.

Run from the project root after installing requirements.txt.  The five original
UCI CSV files must be present in data/raw.  All model selection uses 2012 and
2013; 2014-2015 are held out until the selected model has been refitted.
"""

from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
import joblib
from scipy.stats import jarque_bera
from sklearn.compose import TransformedTargetRegressor
from sklearn.dummy import DummyRegressor
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import PolynomialFeatures, StandardScaler
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.outliers_influence import variance_inflation_factor


ROOT = Path(__file__).resolve().parent
RAW = ROOT / "data" / "raw"
OUT = ROOT / "outputs"
FIG = OUT / "figures"
TARGET = "CO"
FEATURES = ["AT", "AP", "AH", "AFDP", "GTEP", "TIT", "TAT", "TEY", "CDP"]
AMBIENT = ["AT", "AP", "AH"]
YEARS = range(2011, 2016)


def load_data() -> pd.DataFrame:
    frames = []
    for year in YEARS:
        path = RAW / f"gt_{year}.csv"
        if not path.exists():
            raise FileNotFoundError(f"Missing {path}; download UCI dataset 551")
        frame = pd.read_csv(path)
        expected = set(FEATURES + ["CO", "NOX"])
        if set(frame.columns) != expected:
            raise ValueError(f"Unexpected columns in {path}: {list(frame.columns)}")
        frame["year"] = year
        frames.append(frame)
    return pd.concat(frames, ignore_index=True)


def candidate_specs():
    specs = [
        ("mean_baseline", FEATURES, lambda: DummyRegressor(strategy="mean")),
        ("ambient_ols", AMBIENT, lambda: make_pipeline(StandardScaler(), LinearRegression())),
        ("full_ols", FEATURES, lambda: make_pipeline(StandardScaler(), LinearRegression())),
        (
            "log_target_ols",
            FEATURES,
            lambda: TransformedTargetRegressor(
                regressor=make_pipeline(StandardScaler(), LinearRegression()),
                func=np.log1p,
                inverse_func=np.expm1,
                check_inverse=False,
            ),
        ),
        ("full_ridge_100", FEATURES, lambda: make_pipeline(StandardScaler(), Ridge(alpha=100))),
        (
            "quadratic_ols",
            FEATURES,
            lambda: make_pipeline(
                StandardScaler(),
                PolynomialFeatures(degree=2, include_bias=False),
                StandardScaler(),
                LinearRegression(),
            ),
        ),
    ]
    for alpha in (1, 10, 100, 1000, 10000):
        specs.append(
            (
                f"quadratic_ridge_{alpha}",
                FEATURES,
                lambda alpha=alpha: make_pipeline(
                    StandardScaler(),
                    PolynomialFeatures(degree=2, include_bias=False),
                    StandardScaler(),
                    Ridge(alpha=alpha),
                ),
            )
        )
    return specs


def metrics(y_true, y_pred):
    y_true = np.asarray(y_true)
    y_pred = np.asarray(y_pred)
    return {
        "MAE_mg_m3": float(mean_absolute_error(y_true, y_pred)),
        "RMSE_mg_m3": float(np.sqrt(mean_squared_error(y_true, y_pred))),
        "R2": float(r2_score(y_true, y_pred)),
        "bias_pred_minus_actual_mg_m3": float(np.mean(y_pred - y_true)),
        "negative_predictions": int(np.sum(y_pred < 0)),
    }


def evaluate_candidates(data):
    results = []
    for name, columns, factory in candidate_specs():
        folds = []
        for val_year in (2012, 2013):
            train = data[data.year < val_year]
            valid = data[data.year == val_year]
            model = factory()
            model.fit(train[columns], train[TARGET])
            folds.append({"validation_year": val_year, **metrics(valid[TARGET], model.predict(valid[columns]))})
        results.append(
            {
                "model": name,
                "features": columns,
                "folds": folds,
                "mean_validation_MAE_mg_m3": float(np.mean([f["MAE_mg_m3"] for f in folds])),
                "mean_validation_RMSE_mg_m3": float(np.mean([f["RMSE_mg_m3"] for f in folds])),
            }
        )
    results.sort(key=lambda x: x["mean_validation_MAE_mg_m3"])
    return results


def ols_diagnostics(train):
    x = sm.add_constant(train[FEATURES], has_constant="add")
    fit = sm.OLS(train[TARGET], x).fit()
    residuals = fit.resid.to_numpy()
    bp_lm, bp_lm_p, bp_f, bp_f_p = het_breuschpagan(residuals, x)
    standardized = StandardScaler().fit_transform(train[FEATURES])
    standardized = np.column_stack([np.ones(len(standardized)), standardized])
    vifs = {
        name: float(variance_inflation_factor(standardized, i + 1))
        for i, name in enumerate(FEATURES)
    }
    jb = jarque_bera(residuals)
    return {
        "matrix_rank": int(np.linalg.matrix_rank(x)),
        "parameter_count": int(x.shape[1]),
        "vif": vifs,
        "breusch_pagan_LM": float(bp_lm),
        "breusch_pagan_p": float(bp_lm_p),
        "durbin_watson": float(durbin_watson(residuals)),
        "jarque_bera_p_optional": float(jb.pvalue),
        "ols_training_R2": float(fit.rsquared),
        "ols_parameters": {key: float(value) for key, value in fit.params.items()},
    }


def save_figures(data, train, test, predictions, diagnostics):
    FIG.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.size": 10, "figure.facecolor": "white"})

    fig, axes = plt.subplots(1, 2, figsize=(11.5, 4.2), constrained_layout=True)
    axes[0].hist(data[TARGET], bins=np.linspace(0, 45, 65), color="#1b6776", edgecolor="white")
    axes[0].set(xlabel="CO (мг/м³)", ylabel="Число наблюдений", title="Распределение CO")
    axes[1].boxplot([data.loc[data.year == y, TARGET] for y in YEARS], tick_labels=list(YEARS), showfliers=False)
    axes[1].set(xlabel="Год", ylabel="CO (мг/м³)", title="CO по годам (выбросы скрыты)")
    fig.savefig(FIG / "target_distribution.png", dpi=180)
    plt.close(fig)

    rng = np.random.default_rng(42)
    take = rng.choice(len(test), min(3500, len(test)), replace=False)
    actual = test[TARGET].to_numpy()[take]
    pred = np.asarray(predictions)[take]
    fig, axes = plt.subplots(1, 2, figsize=(11.5, 4.2), constrained_layout=True)
    axes[0].scatter(actual, pred, s=7, alpha=0.3, color="#1b6776")
    axes[0].plot([0, 45], [0, 45], color="#c34e36", lw=1.5)
    axes[0].set(xlabel="Фактический CO (мг/м³)", ylabel="Прогноз CO (мг/м³)", title="Отложенные 2014–2015 годы")
    axes[0].set_xlim(0, 45)
    axes[0].set_ylim(-1, 45)
    axes[1].scatter(pred, actual - pred, s=7, alpha=0.3, color="#1b6776")
    axes[1].axhline(0, color="#c34e36", lw=1.5)
    axes[1].set(xlabel="Прогноз CO (мг/м³)", ylabel="Факт − прогноз (мг/м³)", title="Ошибки прогноза")
    fig.savefig(FIG / "test_fit_residuals.png", dpi=180)
    plt.close(fig)

    vif = diagnostics["vif"]
    fig, ax = plt.subplots(figsize=(8.8, 4.4), constrained_layout=True)
    names = sorted(vif, key=vif.get)
    ax.barh(names, [vif[n] for n in names], color="#1b6776")
    ax.axvline(10, color="#c34e36", ls="--", lw=1.5, label="VIF = 10")
    ax.set(xlabel="Коэффициент VIF (логарифмическая шкала)", title="Связь признаков в модели OLS")
    ax.set_xscale("log")
    ax.legend()
    fig.savefig(FIG / "vif.png", dpi=180)
    plt.close(fig)


def main():
    OUT.mkdir(exist_ok=True)
    data = load_data()
    if data[FEATURES + [TARGET]].isna().any().any():
        raise ValueError("Unexpected missing values; inspect the original UCI files")
    train = data[data.year <= 2013]
    test = data[data.year >= 2014]
    candidate_results = evaluate_candidates(data)
    best_name = candidate_results[0]["model"]
    name, columns, factory = next(s for s in candidate_specs() if s[0] == best_name)
    final_model = factory()
    final_model.fit(train[columns], train[TARGET])
    predictions = final_model.predict(test[columns])
    joblib.dump(final_model, OUT / "model.joblib")
    if best_name.startswith("quadratic_ridge_"):
        terms = final_model.named_steps["polynomialfeatures"].get_feature_names_out(columns)
        pd.DataFrame(
            {"term": terms, "coefficient_after_scaling": final_model.named_steps["ridge"].coef_}
        ).to_csv(OUT / "model_coefficients.csv", index=False)
    full_ols = make_pipeline(StandardScaler(), LinearRegression()).fit(train[FEATURES], train[TARGET])
    baseline = DummyRegressor(strategy="mean").fit(train[FEATURES], train[TARGET])
    diagnostics = ols_diagnostics(train)

    by_year = {}
    for year in YEARS:
        part = data[data.year == year]
        by_year[str(year)] = {
            "rows": int(len(part)),
            "CO_mean_mg_m3": float(part[TARGET].mean()),
            "CO_median_mg_m3": float(part[TARGET].median()),
            "CO_std_mg_m3": float(part[TARGET].std()),
            "CO_p99_mg_m3": float(part[TARGET].quantile(0.99)),
        }
    results = {
        "source": "https://archive.ics.uci.edu/dataset/551/gas%2Bturbine%2Bco%2Band%2Bnox%2Bemission%2Bdata%2Bset",
        "source_doi": "10.24432/C5WC95",
        "target": TARGET,
        "predictors": FEATURES,
        "split": {"selection_train": "all years before each validation year", "validation_years": [2012, 2013], "final_train_years": [2011, 2012, 2013], "holdout_years": [2014, 2015]},
        "data_quality": {
            "rows": int(len(data)),
            "columns_in_original_files": 11,
            "total_missing_values": int(data.isna().sum().sum()),
            "exact_duplicate_measurements_with_year": int(data.duplicated().sum()),
            "CO_min_mg_m3": float(data[TARGET].min()),
            "CO_max_mg_m3": float(data[TARGET].max()),
            "CO_median_mg_m3": float(data[TARGET].median()),
            "CO_p99_mg_m3": float(data[TARGET].quantile(0.99)),
            "rows_by_year": by_year,
        },
        "candidate_models": candidate_results,
        "selected_model": best_name,
        "holdout": {
            "selected_overall": metrics(test[TARGET], predictions),
            "selected_2014": metrics(test.loc[test.year == 2014, TARGET], predictions[test.year.to_numpy() == 2014]),
            "selected_2015": metrics(test.loc[test.year == 2015, TARGET], predictions[test.year.to_numpy() == 2015]),
            "full_ols_overall": metrics(test[TARGET], full_ols.predict(test[FEATURES])),
            "training_mean_baseline_overall": metrics(test[TARGET], baseline.predict(test[FEATURES])),
        },
        "gauss_markov_diagnostics_for_full_ols": diagnostics,
        "notes": [
            "VIF and residual tests diagnose problems but cannot prove exogeneity.",
            "Durbin-Watson uses the source row order as an approximate hourly sequence; timestamps are absent.",
            "The final ridge model is linear in expanded features but ridge estimates are biased, so the Gauss-Markov BLUE theorem does not apply to it.",
            "The 2014-2015 holdout was excluded from candidate and alpha selection.",
        ],
    }
    save_figures(data, train, test, predictions, diagnostics)
    (OUT / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Selected:", best_name)
    print("Validation MAE:", round(candidate_results[0]["mean_validation_MAE_mg_m3"], 3))
    print("Holdout:", results["holdout"]["selected_overall"])
    print("OLS diagnostics:", {k: v for k, v in diagnostics.items() if k not in ("vif", "ols_parameters")})


if __name__ == "__main__":
    main()
