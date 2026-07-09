#Sensitivity Analysis
\
import os
import warnings
import numpy as np
import pandas as pd
from scipy.optimize import minimize
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import rcParams

warnings.filterwarnings("ignore")

rcParams["font.family"]      = "serif"
rcParams["font.serif"]       = ["Times New Roman", "Liberation Serif", "DejaVu Serif", "serif"]
rcParams["font.size"]        = 10
rcParams["mathtext.fontset"] = "stix"

DATA_DIR     = "."
RI_CSV       = "RI_values_used_full.csv"
OUTPUT_PNG   = "Fig_HighRiskShare_fine.png"
OUTPUT_CSV   = "HighRiskShare_fine.csv"
RI_THRESHOLD = 0.4
ALPHAS       = [round(a, 2) for a in np.arange(0.0, 1.0001, 0.05)]

FILES = {
    "Maize":      dict(file="Imp_Maize.xlsx",     kind="import"),
    "Beet sugar": dict(file="Imp_Beetsugar.xlsx", kind="import"),
    "Canola":     dict(file="Exp_Canola.xlsx",    kind="export"),
    "Wheat":      dict(file="Exp_Wheat.xlsx",     kind="export"),
    "Soybean":    dict(file="Exp_Soybean.xlsx",   kind="export"),
    "Colza":      dict(file="Exp_Colza.xlsx",     kind="export"),
}
SHEETS = {
    "import": ("imp-mass", "imp_value", "Maximum_export"),
    "export": ("exp-mass", "exp_value", "Maximum_import"),
}

PLOT_COMMODITIES = ["Canola", "Soybean", "Wheat", "Colza", "Maize"]
SWEEP_COMMODITIES = ["Maize", "Beet sugar", "Canola", "Wheat", "Soybean", "Colza"]
COLORS  = {"Canola":"#3a4cb4","Soybean":"#3a9b3a","Wheat":"#e67e22",
           "Colza":"#7b2cbf","Maize":"#c0392b"}
MARKERS = {"Canola":"o","Soybean":"^","Wheat":"s","Colza":"D","Maize":"v"}

FS_AXIS_LABEL  = 11
FS_TICK_LABEL  = 11
FS_RIGHT_LABEL = 12

def fuzzy_find(name, candidates):
    name = name.lower().strip()
    for c in candidates:
        if c.lower() == name:
            return c
    for c in candidates:
        if len(name) >= 5 and name[:5] in c.lower():
            return c
        if len(c) >= 5 and c.lower()[:5] in name:
            return c
    return None

class Engine:
    \
\
    def __init__(self, fn, mass_sheet, value_sheet, cap_sheet, ri_map, clean=True):
        xl = pd.ExcelFile(fn)
        dm = pd.read_excel(xl, sheet_name=mass_sheet,  index_col=0)
        dv = pd.read_excel(xl, sheet_name=value_sheet, index_col=0)
        dm = dm[dm.index.notna()]; dm.index = dm.index.map(str).str.strip()
        dv = dv[dv.index.notna()]; dv.index = dv.index.map(str).str.strip()
        common = sorted(set(dm.index) & set(dv.index))
        dm = dm.loc[common].fillna(0); dv = dv.loc[common].fillna(0)
        dm = dm[~dm.index.duplicated()]; dv = dv[~dv.index.duplicated()]

        years = sorted([int(c) for c in dm.columns
                        if str(c).strip().isdigit() and 1900 < int(c) < 2100])
        self.TEST  = 2024 if 2024 in years else max(years)
        self.TRAIN = [y for y in years if y < self.TEST]

        M = dm[years].apply(pd.to_numeric, errors="coerce").fillna(0)
        V = dv[years].apply(pd.to_numeric, errors="coerce").fillna(0)

        train = M[self.TRAIN]
        c1 = ((train > 0).sum(axis=1) >= 3) & (M[self.TEST] > 0)
        c2 = (train > 0).sum(axis=1) >= 30
        pool = M.index[(c1 | c2)].tolist()
        self.countries = pool; self.N = len(pool)
        M = M.loc[pool]; V = V.loc[pool]

        if clean:
            price = V.replace(0, np.nan) / M.replace(0, np.nan)
        else:
            price = V / M.replace(0, np.nan)
        tp = price[self.TRAIN]
        self.valid_price_years = ((~tp.isna()) & (tp > 0)).sum(axis=1).values

        mp = np.zeros(self.N)
        for i, c in enumerate(pool):
            pr = tp.loc[c].values
            v  = pr[(~np.isnan(pr)) & (pr > 0)]
            mp[i] = v.mean() if len(v) > 0 else 0
        med = np.median(mp[mp > 0])
        mp[mp == 0] = med
        self.mu = mp

        cov = np.zeros((self.N, self.N))
        TP = tp.values
        for i in range(self.N):
            for j in range(i, self.N):
                pi, pj = TP[i], TP[j]
                ok = (~np.isnan(pi)) & (~np.isnan(pj)) & (pi > 0) & (pj > 0)
                if ok.sum() >= 2:
                    cv_ = np.cov(pi[ok], pj[ok], ddof=1)[0, 1]
                    cov[i, j] = cv_; cov[j, i] = cv_
                elif i == j:
                    pv = pi[(~np.isnan(pi)) & (pi > 0)]
                    cov[i, i] = np.var(pv, ddof=1) if len(pv) >= 2 else 0.001
        ev, evec = np.linalg.eigh(cov)
        ev = np.maximum(ev, 1e-10)
        cov = evec @ np.diag(ev) @ evec.T
        self.cov = (cov + cov.T) / 2

        m24 = M[self.TEST].values.astype(float)
        self.D = m24.sum()
        self.w_emp   = m24 / self.D
        self.emp_ret = float(self.w_emp @ self.mu)

        last_val = np.zeros(self.N)
        for i, c in enumerate(pool):
            s  = M.loc[c]
            nz = [y for y in years if s[y] > 0]
            last_val[i] = s[max(nz)] if nz else 0
        cons = np.where(m24 > 0, m24, last_val)
        exp1 = cons.copy()
        try:
            mx = pd.read_excel(xl, sheet_name=cap_sheet, index_col=0)
            mx.index = mx.index.astype(str).str.strip()
            mx = mx[~mx.index.duplicated(keep="first")]
            for i, c in enumerate(pool):
                fv = None
                tgt = c if c in mx.index else fuzzy_find(c, mx.index.tolist())
                if tgt is not None and tgt in mx.index:
                    row = mx.loc[tgt]
                    if isinstance(row, pd.Series):
                        pref = [c for c in row.index if str(c).strip() in (str(self.TEST), f"{self.TEST}.0")]
                        for col in pref + [c for c in row.index if c not in pref]:
                            try:
                                if pd.notna(row[col]) and float(row[col]) > 0:
                                    fv = float(row[col]); break
                            except Exception:
                                pass
                if fv is not None and fv > 0:
                    exp1[i] = fv
            exp1 = np.maximum(exp1, cons)
        except Exception:
            pass
        self.cap = exp1

        ri = []
        for c in pool:
            r = ri_map.get(c)
            if r is None or (isinstance(r, float) and np.isnan(r)):
                k = fuzzy_find(c, list(ri_map))
                r = ri_map.get(k) if k else np.nan
            ri.append(r)
        ri = np.array(ri, dtype=float)
        if np.isnan(ri).any():
            ri = np.where(np.isnan(ri), np.nanmedian(ri), ri)
        self.RI = ri

    def opt(self, alpha, min_return=None, drop_no_price=True):
        active = [i for i, _ in enumerate(self.countries)
                  if (not drop_no_price) or self.valid_price_years[i] >= 2]
        idx = np.array(active); n = len(idx)
        mu = self.mu[idx]; cov = self.cov[np.ix_(idx, idx)]
        cap = self.cap[idx]; ri = self.RI[idx]
        k = (1 + alpha * ri) if alpha > 0 else np.ones(n)
        A = np.diag(k) @ cov @ np.diag(k)
        ev, evec = np.linalg.eigh(A); ev = np.maximum(ev, 1e-10)
        A = evec @ np.diag(ev) @ evec.T
        bounds = [(0, max(min(1.0, cap[i] / self.D) if cap[i] > 0 else 0.0, 1e-10))
                  for i in range(n)]
        if sum(b[1] for b in bounds) < 0.999:
            return None
        cons = [{"type": "eq", "fun": lambda w: w.sum() - 1}]
        if min_return is not None:
            cons.append({"type": "ineq", "fun": lambda w: w @ mu - min_return})
        w0 = np.array([b[1] for b in bounds]); w0 = w0 / w0.sum()
        r = minimize(lambda w: w @ A @ w, w0,
                     jac=lambda w: 2 * A @ w,
                     method="SLSQP", bounds=bounds, constraints=cons,
                     options={"maxiter": 1000, "ftol": 1e-12})
        w = np.maximum(r.x, 0); w = w / w.sum()
        return dict(weights=dict(zip([self.countries[i] for i in idx], w)))


def high_risk_share(weights, ri_map):
    s = 0.0
    for partner, w in weights.items():
        r = ri_map.get(partner)
        if r is None or (isinstance(r, float) and np.isnan(r)):
            k = fuzzy_find(partner, list(ri_map))
            r = ri_map.get(k) if k else None
        if r is not None and not (isinstance(r, float) and np.isnan(r)) and r > RI_THRESHOLD:
            s += w
    return 100.0 * s


def main():
    ri_df = pd.read_csv(os.path.join(DATA_DIR, RI_CSV))
    ri_map = {str(k).strip(): float(v)
              for k, v in zip(ri_df.iloc[:, 0],
                              pd.to_numeric(ri_df.iloc[:, 1], errors="coerce"))
              if pd.notna(v)}

    print(f"Running fine α sweep (Δα = 0.05, {len(ALPHAS)} values)...")

    results, var_results, wheat_rows = {}, {}, []
    for com in SWEEP_COMMODITIES:
        cfg = FILES[com]
        fp = os.path.join(DATA_DIR, cfg["file"])
        if not os.path.exists(fp):
            print(f"  [skip] {com}: missing input file {fp}")
            continue
        ms, vs, cs = SHEETS[cfg["kind"]]
        eng = Engine(fp, ms, vs, cs, ri_map, clean=True)
        emp_var = float(eng.w_emp @ eng.cov @ eng.w_emp)
        series, vseries = [], []
        for a in ALPHAS:
            r = eng.opt(a, min_return=eng.emp_ret, drop_no_price=True)
            wd = r["weights"]
            series.append(high_risk_share(wd, ri_map))
            wvec = np.array([wd.get(c, 0.0) for c in eng.countries])
            var = float(wvec @ eng.cov @ wvec)
            vseries.append((var - emp_var) / emp_var * 100.0)
            if com == "Wheat" and a in (0.0, 0.5, 1.0):
                for c, w in wd.items():
                    if w > 1e-6:
                        wheat_rows.append(dict(alpha=a, country=c, weight_pct=round(100*w, 2)))
        results[com], var_results[com] = series, vseries
        print(f"  {com:<11} HRS a=0:{series[0]:5.1f} a=1:{series[-1]:5.1f} | "
              f"dVar a=0:{vseries[0]:+.1f}% a=0.5:{vseries[len(ALPHAS)//2]:+.1f}% a=1:{vseries[-1]:+.1f}%")

    pd.DataFrame([dict(commodity=c, alpha=a, var_change_pct=round(v, 2))
                  for c, vs_ in var_results.items() for a, v in zip(ALPHAS, vs_)]
                 ).to_csv("Sensitivity_varchange_fine.csv", index=False)
    pd.DataFrame(wheat_rows).to_csv("Wheat_weights_by_alpha.csv", index=False)

    rows = []
    for com, series in results.items():
        for a, s in zip(ALPHAS, series):
            rows.append(dict(commodity=com, alpha=a, high_risk_share_pct=round(s, 2)))
    pd.DataFrame(rows).to_csv(OUTPUT_CSV, index=False)


    fig, ax = plt.subplots(figsize=(9, 5.5))
    for com in PLOT_COMMODITIES:
        if com not in results:
            continue
        series = results[com]
        ax.plot(ALPHAS, series,
                color=COLORS[com], marker=MARKERS[com], markersize=5,
                linewidth=1.8, label=com,
                markeredgecolor="white", markeredgewidth=0.5)
        ax.annotate(com, xy=(1.0, series[-1]), xytext=(10, 0),
                    textcoords="offset points",
                    color=COLORS[com],
                    fontsize=FS_RIGHT_LABEL, fontweight="bold", va="center")

    ax.set_xlabel("Risk Aversion Parameter (α)", fontsize=FS_AXIS_LABEL)
    ax.set_ylabel("High-Risk Partner Share (%)\n(Partners with RI > 0.4)",
                  fontsize=FS_AXIS_LABEL)
    ax.set_xticks([0.0, 0.25, 0.5, 0.75, 1.0])
    ax.tick_params(axis="both", labelsize=FS_TICK_LABEL)
    ax.grid(alpha=0.3)
    ax.set_xlim(-0.02, 1.14)
    ax.set_ylim(15, 75)
    plt.tight_layout()
    plt.savefig(OUTPUT_PNG, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"\nSaved {OUTPUT_PNG}, {OUTPUT_CSV}, Sensitivity_varchange_fine.csv, Wheat_weights_by_alpha.csv")

if __name__ == "__main__":
    main()
