#Export Portfolio 

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import minimize
import warnings
import os
import sys
warnings.filterwarnings('ignore')

GUROBI_AVAILABLE = False
try:
    from gurobipy import Model, GRB
    GUROBI_AVAILABLE = True
except ImportError:
    pass
\

COMMODITY = "Cattle"
INPUT_FILE = f"Exp_{COMMODITY}.xlsx"

SHEET_MASS = 'exp-mass'
SHEET_VALUE = 'exp_value'
SHEET_IMPORT_CAP = 'Maximum_import'

ALPHA = 0.5
MIN_TRAINING_YEARS = 3
MIN_HISTORICAL_YEARS = 30
TEST_YEAR = 2024

W_PS, W_CV, W_LPI, W_WSI = 0.299, 0.329, 0.328, 0.044

PS_FILE = "Political_Stability_Indicators.csv"
CV_FILE = "Climate_Vulnerability_indicator.csv"
LPI_FILE = "Logistic_Performance_Indicator.csv"
WSI_FILE = "Water_Stress_Indicator.csv"

OUTPUT_DIR = "."
SAVE_FIGURES = True
FIGURE_DPI = 300

\
\
COUNTRY_NAME_MAP = {
    'Belgium & Luxembourg': 'Belgium',
    'United Kingdom of Great Britain and Northern Ireland': 'United Kingdom',
    'United States of America': 'United States',
    'Venezuela (Bolivarian Republic of)': 'Venezuela',
    'Iran (Islamic Republic of)': 'Iran',
    'Republic of Korea': 'South Korea',
    'Türkiye': 'Turkey',
    'Viet Nam': 'Vietnam',
}

def get_mapped_name(name):
    return COUNTRY_NAME_MAP.get(name.strip(), name.strip())

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

def safe_load_csv(filepath, name, index_col=None):
    if not os.path.exists(filepath):
        \
        return None
    try:
        if index_col is not None:
            df = pd.read_csv(filepath, index_col=index_col, encoding='latin-1')
            df.index = df.index.astype(str).str.strip()
        else:
            df = pd.read_csv(filepath, encoding='latin-1')
            for col in ['Country Name', 'Country']:
                if col in df.columns:
                    df['Country'] = df[col].astype(str).str.strip()
                    break
        \
        return df
    except Exception as e:
        \
        return None

REGIONS = {
    'Asia-Pacific': ['China', 'Japan', 'Korea', 'India', 'Bangladesh', 'Pakistan',
                     'Indonesia', 'Malaysia', 'Thailand', 'Vietnam', 'Philippines',
                     'Nepal', 'Hong Kong', 'Taiwan', 'Singapore'],
    'Europe': ['Germany', 'Netherlands', 'France', 'Belgium', 'Luxembourg', 'Italy',
               'Spain', 'Portugal', 'United Kingdom', 'Ireland', 'Poland', 'Denmark',
               'Bulgaria', 'Romania', 'Sweden', 'Norway', 'Finland'],
    'Middle East': ['United Arab Emirates', 'Saudi Arabia', 'Iran', 'Turkey',
                    'Israel', 'Jordan', 'Kuwait'],
    'Americas': ['Mexico', 'Brazil', 'Argentina', 'Chile', 'Colombia', 'Peru',
                 'Ecuador', 'Costa Rica', 'Guatemala', 'Venezuela', 'United States']
}

def get_region(country):
    for r, kw in REGIONS.items():
        if any(k.lower() in country.lower() for k in kw):
            return r
    return 'Other'

if not os.path.exists(INPUT_FILE):
    \
    sys.exit(1)

xlsx = pd.ExcelFile(INPUT_FILE)
\
\

def find_sheet(xlsx, options):
    for opt in options:
        if opt in xlsx.sheet_names:
            return opt
    return None

mass_sheet = find_sheet(xlsx, [SHEET_MASS, 'exp_mass', 'exp-mass'])
value_sheet = find_sheet(xlsx, [SHEET_VALUE, 'exp-value'])
import_sheet = find_sheet(xlsx, [SHEET_IMPORT_CAP, 'maximum_import'])

if mass_sheet is None or value_sheet is None:
    \
    sys.exit(1)

exp_mass_raw = pd.read_excel(xlsx, sheet_name=mass_sheet, index_col=0)
exp_mass_raw.index = exp_mass_raw.index.astype(str).str.strip()

exp_value_raw = pd.read_excel(xlsx, sheet_name=value_sheet, index_col=0)
exp_value_raw.index = exp_value_raw.index.astype(str).str.strip()

\
\

common_countries = sorted(list(set(exp_mass_raw.index) & set(exp_value_raw.index)))
exp_mass_aligned = exp_mass_raw.loc[common_countries].fillna(0)
exp_value_aligned = exp_value_raw.loc[common_countries].fillna(0)
\

def get_year_columns(df):
    years = []
    for col in df.columns:
        try:
            y = int(col)
            if 1900 < y < 2100:
                years.append(y)
        except:
            pass
    return sorted(years)

all_years = get_year_columns(exp_mass_aligned)
TRAIN_YEARS = [y for y in all_years if y < TEST_YEAR]

if TEST_YEAR not in all_years:
    TEST_YEAR = max(all_years)
    TRAIN_YEARS = [y for y in all_years if y < TEST_YEAR]

\

exp_mass = exp_mass_aligned[all_years].copy()
exp_value = exp_value_aligned[all_years].copy()

has_import_cap = import_sheet is not None
if has_import_cap:
    max_import = pd.read_excel(xlsx, sheet_name=import_sheet, index_col=0)
    max_import.index = max_import.index.astype(str).str.strip()
    max_import = max_import[~max_import.index.duplicated(keep='first')]
else:
    max_import = None
\

train_mass = exp_mass[TRAIN_YEARS]

years_with_trade_train = (train_mass > 0).sum(axis=1)
has_training = years_with_trade_train >= MIN_TRAINING_YEARS
has_2024 = exp_mass[TEST_YEAR] > 0
criterion1 = has_training & has_2024
n_criterion1 = criterion1.sum()

historical_years_traded = (train_mass > 0).sum(axis=1)
criterion2 = historical_years_traded >= MIN_HISTORICAL_YEARS
n_criterion2_only = (criterion2 & ~criterion1).sum()

valid_mask = criterion1 | criterion2
valid_countries = exp_mass.index[valid_mask].tolist()

\
\
\
\

new_countries = exp_mass.index[criterion2 & ~criterion1].tolist()
if new_countries:
    \
    for c in new_countries[:10]:
        yrs = historical_years_traded[c]
    if len(new_countries) > 10:
        pass
\

exp_mass = exp_mass.loc[valid_countries].copy()
exp_value = exp_value.loc[valid_countries].copy()

all_countries = valid_countries
N = len(all_countries)

exp_unit_price = exp_value / exp_mass.replace(0, np.nan)
train_prices = exp_unit_price[TRAIN_YEARS]

mean_prices = np.zeros(N)
for i, country in enumerate(all_countries):
    prices = train_prices.loc[country].values
    valid = prices[~np.isnan(prices) & (prices > 0)]
    mean_prices[i] = np.mean(valid) if len(valid) > 0 else 0

median_price = np.median(mean_prices[mean_prices > 0])
mean_prices[mean_prices == 0] = median_price

\

cov_matrix = np.zeros((N, N))
for i in range(N):
    for j in range(i, N):
        pi = train_prices.iloc[i].values
        pj = train_prices.iloc[j].values
        valid = (~np.isnan(pi)) & (~np.isnan(pj)) & (pi > 0) & (pj > 0)
        if valid.sum() >= 2:
            cov_val = np.cov(pi[valid], pj[valid], ddof=1)[0, 1]
            cov_matrix[i, j] = cov_val
            cov_matrix[j, i] = cov_val
        elif i == j:
            pi_valid = pi[~np.isnan(pi) & (pi > 0)]
            cov_matrix[i, i] = np.var(pi_valid, ddof=1) if len(pi_valid) >= 2 else 0.001

eigvals, eigvecs = np.linalg.eigh(cov_matrix)
eigvals = np.maximum(eigvals, 1e-10)
cov_matrix = eigvecs @ np.diag(eigvals) @ eigvecs.T
cov_matrix = (cov_matrix + cov_matrix.T) / 2

\

exp_mass_2024 = exp_mass[TEST_YEAR].values
total_2024 = exp_mass_2024.sum()
D = total_2024

empirical_weights = exp_mass_2024 / total_2024

emp_variance = empirical_weights @ cov_matrix @ empirical_weights
emp_volatility = np.sqrt(emp_variance)
emp_return = empirical_weights @ mean_prices
emp_hhi = np.sum(empirical_weights ** 2)

\
\

\
sorted_idx = np.argsort(empirical_weights)[::-1]
for rank, i in enumerate(sorted_idx[:10]):
    pass
\

max_historical_trade = exp_mass[TRAIN_YEARS + [TEST_YEAR]].max(axis=1).values
last_trade_year = {}
last_trade_value = {}

for i, country in enumerate(all_countries):
    trade_series = exp_mass.loc[country]
    nonzero_years = [y for y in all_years if trade_series[y] > 0]
    if nonzero_years:
        last_yr = max(nonzero_years)
        last_trade_year[country] = last_yr
        last_trade_value[country] = trade_series[last_yr]
    else:
        last_trade_year[country] = None
        last_trade_value[country] = 0

capacity_conservative = np.zeros(N)
for i, country in enumerate(all_countries):
    if exp_mass_2024[i] > 0:
        capacity_conservative[i] = exp_mass_2024[i]
    else:

        capacity_conservative[i] = last_trade_value.get(country, 0)

\
\

capacity_conservative_alt = max_historical_trade.copy()
\
\

capacity_expansion1 = np.zeros(N)
if has_import_cap and max_import is not None:
    for i, country in enumerate(all_countries):
        found_val = None

        if country in max_import.index:
            try:
                row = max_import.loc[country]
                if isinstance(row, pd.Series):
                    for col in row.index:
                        try:
                            if pd.notna(row[col]) and float(row[col]) > 0:
                                found_val = float(row[col])
                                break
                        except:
                            pass
                else:
                    found_val = float(row)
            except:
                pass

        if found_val is None:
            match = fuzzy_find(country, max_import.index.tolist())
            if match:
                try:
                    row = max_import.loc[match]
                    if isinstance(row, pd.Series):
                        for col in row.index:
                            try:
                                if pd.notna(row[col]) and float(row[col]) > 0:
                                    found_val = float(row[col])
                                    break
                            except:
                                pass
                    else:
                        found_val = float(row)
                except:
                    pass

        if found_val is not None and found_val > 0:
            capacity_expansion1[i] = found_val
        else:

            capacity_expansion1[i] = capacity_conservative[i]

    capacity_expansion1 = np.maximum(capacity_expansion1, capacity_conservative)
else:
    capacity_expansion1 = capacity_conservative.copy()
\

capacity_expansion2 = max_historical_trade.copy()
capacity_expansion2 = np.maximum(capacity_expansion2, capacity_conservative)
\

ps_df = safe_load_csv(PS_FILE, "PS")
cv_df = safe_load_csv(CV_FILE, "CV", index_col=0)
lpi_df = safe_load_csv(LPI_FILE, "LPI")
wsi_df = safe_load_csv(WSI_FILE, "WSI")

risk_data = []
for country in all_countries:
    mapped = get_mapped_name(country)
    ps, cv, lpi, wsi = None, None, None, None

    if ps_df is not None and 'Country' in ps_df.columns:
        match = fuzzy_find(mapped, ps_df['Country'].tolist())
        if match:
            row = ps_df[ps_df['Country'] == match]
            if len(row) > 0:
                for col in ['2023', '2022', '2021']:
                    if col in row.columns:
                        v = row[col].values[0]
                        if str(v) != '..' and pd.notna(v):
                            try: ps = float(v)
                            except: pass
                            break

    if cv_df is not None:
        match = fuzzy_find(mapped, cv_df.index.tolist())
        if match:
            for col in ['2023', '2022', '2021']:
                if col in cv_df.columns:
                    try: cv = float(cv_df.loc[match, col]); break
                    except: pass

    if lpi_df is not None and 'Country' in lpi_df.columns:
        match = fuzzy_find(mapped, lpi_df['Country'].tolist())
        if match and 'average' in lpi_df.columns:
            row = lpi_df[lpi_df['Country'] == match]
            if len(row) > 0:
                lpi = float(row['average'].values[0])

    if wsi_df is not None and 'Country' in wsi_df.columns:
        match = fuzzy_find(mapped, wsi_df['Country'].tolist())
        if match and 'WSI_risk' in wsi_df.columns:
            row = wsi_df[wsi_df['Country'] == match]
            if len(row) > 0 and pd.notna(row['WSI_risk'].values[0]):
                wsi = float(row['WSI_risk'].values[0])

    risk_data.append({'Country': country, 'PS_raw': ps, 'CV_raw': cv, 'LPI_raw': lpi, 'WSI_raw': wsi})

risk_df = pd.DataFrame(risk_data)

risk_df['PS_risk'] = risk_df['PS_raw'].apply(lambda x: (2.5-x)/5.0 if pd.notna(x) else None)
risk_df['CV_risk'] = risk_df['CV_raw']
risk_df['LPI_risk'] = risk_df['LPI_raw'].apply(lambda x: (5-x)/4.0 if pd.notna(x) else None)
risk_df['WSI_risk'] = risk_df['WSI_raw']

for col in ['PS_risk', 'CV_risk', 'LPI_risk', 'WSI_risk']:
    med = risk_df[col].median()
    risk_df[col] = risk_df[col].fillna(med if pd.notna(med) else 0.5)

risk_df['RI'] = W_PS*risk_df['PS_risk'] + W_CV*risk_df['CV_risk'] + W_LPI*risk_df['LPI_risk'] + W_WSI*risk_df['WSI_risk']
risk_df['k_factor'] = 1 + ALPHA * risk_df['RI']

RI = risk_df['RI'].values
k_factors = risk_df['k_factor'].values

\

top_idx = np.argmax(empirical_weights)
TOP_PARTNER = all_countries[top_idx]
\

region_shares = {}
for r in list(REGIONS.keys()) + ['Other']:
    s = sum(empirical_weights[i] for i, c in enumerate(all_countries) if get_region(c) == r)
    if s > 0.001:
        region_shares[r] = s

TOP_REGION = max(region_shares, key=region_shares.get)
TOP_REGION_COUNTRIES = [c for c in all_countries if get_region(c) == TOP_REGION]
\

US_COUNTRY = next((c for c in all_countries if 'United States' in c), None)
if US_COUNTRY:
    us_idx = all_countries.index(US_COUNTRY)
\

PR_COUNTRY = next((c for c in all_countries if 'Puerto Rico' in c), None)
if PR_COUNTRY:
    pr_idx = all_countries.index(PR_COUNTRY)
\

def optimize_gurobi(alpha, capacity, exclusions=None, min_return=None):
    \
    if not GUROBI_AVAILABLE:
        return {'status': 'Gurobi not available'}

    if exclusions is None:
        exclusions = []

    active = np.array([c not in exclusions for c in all_countries])
    idx = np.where(active)[0]
    n = len(idx)

    if n == 0:
        return {'status': 'No countries'}

    mu = mean_prices[idx]
    cov = cov_matrix[np.ix_(idx, idx)]
    cap = capacity[idx]
    ri = RI[idx]

    k = (1 + alpha * ri) if alpha > 0 else np.ones(n)
    K = np.diag(k)
    cov_adj = K @ cov @ K

    ev, evec = np.linalg.eigh(cov_adj)
    ev = np.maximum(ev, 1e-10)
    cov_adj = evec @ np.diag(ev) @ evec.T

    try:
        m = Model('portfolio')
        m.setParam('OutputFlag', 0)
        m.setParam('NumericFocus', 3)

        w = m.addMVar(n, lb=0, ub=1)
        m.addConstr(w.sum() == 1)

        for i in range(n):
            if cap[i] > 0:
                m.addConstr(D * w[i] <= cap[i])

        if min_return is not None:
            m.addConstr(w @ mu >= min_return)

        m.setObjective(w @ cov_adj @ w, GRB.MINIMIZE)
        m.optimize()

        if m.status == GRB.OPTIMAL:
            wopt = w.X
            var = wopt @ cov @ wopt

            weights = np.zeros(N)
            for i_loc, j in enumerate(idx):
                weights[j] = wopt[i_loc]

            return {
                'weights': weights,
                'variance': var,
                'volatility': np.sqrt(max(var, 0)),
                'return': wopt @ mu,
                'hhi': np.sum(wopt**2),
                'weighted_risk': np.sum(wopt * ri),
                'n_active': int(np.sum(wopt > 0.001)),
                'status': 'Optimal',
                'solver': 'Gurobi'
            }
        elif m.status == GRB.INFEASIBLE:
            return {'status': 'Infeasible', 'solver': 'Gurobi'}
        else:
            return {'status': f'Gurobi status {m.status}', 'solver': 'Gurobi'}
    except Exception as e:
        return {'status': f'Gurobi error: {e}', 'solver': 'Gurobi'}

def optimize_scipy(alpha, capacity, exclusions=None, min_return=None):
    \
    if exclusions is None:
        exclusions = []

    active = np.array([c not in exclusions for c in all_countries])
    idx = np.where(active)[0]
    n = len(idx)

    if n == 0:
        return {'status': 'No countries'}

    mu = mean_prices[idx]
    cov = cov_matrix[np.ix_(idx, idx)]
    cap = capacity[idx]
    ri = RI[idx]

    k = (1 + alpha * ri) if alpha > 0 else np.ones(n)
    K = np.diag(k)
    cov_adj = K @ cov @ K

    ev, evec = np.linalg.eigh(cov_adj)
    ev = np.maximum(ev, 1e-10)
    cov_adj = evec @ np.diag(ev) @ evec.T

    def objective(w):
        return w @ cov_adj @ w

    def gradient(w):
        return 2 * cov_adj @ w

    constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]

    if min_return is not None:
        constraints.append({'type': 'ineq', 'fun': lambda w, mu=mu, mr=min_return: w @ mu - mr})

    bounds = []
    for i in range(n):
        ub = min(1.0, cap[i] / D) if cap[i] > 0 else 0.0
        bounds.append((0, max(ub, 1e-10)))

    max_possible = sum(b[1] for b in bounds)
    if max_possible < 0.999:
        return {'status': 'Infeasible', 'solver': 'scipy'}

    w0 = np.array([b[1] for b in bounds])
    w0 = w0 / w0.sum() if w0.sum() > 0 else np.ones(n) / n

    try:
        result = minimize(objective, w0, method='SLSQP', jac=gradient,
                          bounds=bounds, constraints=constraints,
                          options={'maxiter': 1000, 'ftol': 1e-12})

        if result.success or result.fun < 1e10:
            wopt = np.maximum(result.x, 0)
            wopt = wopt / wopt.sum() if wopt.sum() > 0 else w0

            var = wopt @ cov @ wopt

            weights = np.zeros(N)
            for i_loc, j in enumerate(idx):
                weights[j] = wopt[i_loc]

            return {
                'weights': weights,
                'variance': var,
                'volatility': np.sqrt(max(var, 0)),
                'return': wopt @ mu,
                'hhi': np.sum(wopt**2),
                'weighted_risk': np.sum(wopt * ri),
                'n_active': int(np.sum(wopt > 0.001)),
                'status': 'Optimal',
                'solver': 'scipy'
            }
        else:
            return {'status': f'scipy: {result.message}', 'solver': 'scipy'}
    except Exception as e:
        return {'status': f'scipy error: {e}', 'solver': 'scipy'}

def optimize_dual(alpha, capacity, exclusions=None, min_return=None):
    \
    result_scipy = optimize_scipy(alpha, capacity, exclusions, min_return)

    if GUROBI_AVAILABLE:
        result_gurobi = optimize_gurobi(alpha, capacity, exclusions, min_return)
    else:
        result_gurobi = {'status': 'Gurobi not available', 'solver': 'Gurobi'}

    if result_gurobi.get('status') == 'Optimal':
        primary = result_gurobi
    else:
        primary = result_scipy

    return primary, result_gurobi, result_scipy

emp_result = {
    'weights': empirical_weights,
    'variance': emp_variance,
    'volatility': emp_volatility,
    'return': emp_return,
    'hhi': emp_hhi,
    'weighted_risk': np.sum(empirical_weights * RI),
    'n_active': int(np.sum(empirical_weights > 0.001)),
    'status': 'Empirical',
    'solver': 'N/A'
}

scenarios = [
    ('Baseline', 0, []),
    ('Extended', ALPHA, []),
    ('SingleShock', ALPHA, [TOP_PARTNER]),
    ('Regional', ALPHA, TOP_REGION_COUNTRIES),
    ('USExclude', ALPHA, [c for c in [US_COUNTRY, PR_COUNTRY] if c]),
]

capacity_configs = [
    ('Conservative', capacity_conservative),
    ('Conservative_Alt', capacity_conservative_alt),
    ('Expansion1', capacity_expansion1),
    ('Expansion2', capacity_expansion2),
]

all_results = {}
solver_comparison = []

for cap_name, capacity in capacity_configs:
    pass
\
\
\

    all_results[cap_name] = {'optimum': {}, 'optimal': {}}

    all_results[cap_name]['optimum']['Empirical'] = emp_result
    all_results[cap_name]['optimal']['Empirical'] = emp_result

    is_infeasible = False

    \
    for name, alpha, excl in scenarios:
        primary, grb, scp = optimize_dual(alpha, capacity, excl, None)

        if primary.get('status') == 'Optimal':
            all_results[cap_name]['optimum'][name] = primary
            vc = (primary['variance'] - emp_variance) / emp_variance * 100
            rc = (primary['return'] - emp_return) / emp_return * 100
\

            if grb.get('status') == 'Optimal' and scp.get('status') == 'Optimal':
                var_diff = abs(grb['variance'] - scp['variance'])
                solver_comparison.append({
                    'Capacity': cap_name, 'Type': 'Optimum', 'Scenario': name,
                    'Gurobi_Var': grb['variance'], 'Scipy_Var': scp['variance'],
                    'Difference': var_diff, 'Match': 'Yes' if var_diff < 1e-6 else 'No'
                })
        else:
            \
            if cap_name == 'Conservative' and name == 'Baseline':
                is_infeasible = True

    \
    for name, alpha, excl in scenarios:
        primary, grb, scp = optimize_dual(alpha, capacity, excl, emp_return)

        if primary.get('status') == 'Optimal':
            all_results[cap_name]['optimal'][name] = primary
            vc = (primary['variance'] - emp_variance) / emp_variance * 100
\

            if grb.get('status') == 'Optimal' and scp.get('status') == 'Optimal':
                var_diff = abs(grb['variance'] - scp['variance'])
                solver_comparison.append({
                    'Capacity': cap_name, 'Type': 'Optimal', 'Scenario': name,
                    'Gurobi_Var': grb['variance'], 'Scipy_Var': scp['variance'],
                    'Difference': var_diff, 'Match': 'Yes' if var_diff < 1e-6 else 'No'
                })
        else:
            pass
\

    if cap_name == 'Conservative' and is_infeasible:
        pass
\

output_file = os.path.join(OUTPUT_DIR, f"{COMMODITY}_Results_Comprehensive.xlsx")

def make_summary(results, ptype):
    rows = []
    for name in ['Empirical', 'Baseline', 'Extended', 'SingleShock', 'Regional', 'USExclude']:
        if name in results and results[name].get('status') in ['Optimal', 'Empirical']:
            r = results[name]
            rows.append({
                'Scenario': name, 'Type': ptype, 'N_Active': r['n_active'],
                'Variance': r['variance'], 'Var_Change_%': (r['variance']-emp_variance)/emp_variance*100,
                'Volatility_%': r['volatility']*100, 'Return': r['return'],
                'Return_Change_%': (r['return']-emp_return)/emp_return*100,
                'HHI': r['hhi'], 'Weighted_Risk': r['weighted_risk'],
                'Solver': r.get('solver', 'N/A')
            })
    return pd.DataFrame(rows)

with pd.ExcelWriter(output_file, engine='openpyxl') as w:

    for cap_name in ['Conservative', 'Conservative_Alt', 'Expansion1', 'Expansion2']:
        if cap_name in all_results:
            df_optimum = make_summary(all_results[cap_name]['optimum'], 'Optimum')
            df_optimal = make_summary(all_results[cap_name]['optimal'], 'Optimal')

            sheet_optimum = f'{COMMODITY}_{cap_name}_Optimum'[:31]
            sheet_optimal = f'{COMMODITY}_{cap_name}_Optimal'[:31]
            sheet_alloc = f'{COMMODITY}_Alloc_{cap_name}'[:31]

            if not df_optimum.empty:
                df_optimum.to_excel(w, sheet_name=sheet_optimum, index=False)
            if not df_optimal.empty:
                df_optimal.to_excel(w, sheet_name=sheet_optimal, index=False)

            alloc = {'Country': all_countries, 'RI': RI, 'k_factor': k_factors,
                     'Mean_Price': mean_prices,
                     'Cap_Conservative': capacity_conservative,
                     'Cap_Expansion1': capacity_expansion1,
                     'Cap_Expansion2': capacity_expansion2}
            for name, r in all_results[cap_name]['optimum'].items():
                if r.get('status') in ['Optimal', 'Empirical']:
                    alloc[f'{name}_%'] = r['weights'] * 100
            alloc_df = pd.DataFrame(alloc)
            alloc_df = alloc_df.sort_values('Empirical_%', ascending=False)
            alloc_df.to_excel(w, sheet_name=sheet_alloc, index=False)

    risk_df.to_excel(w, sheet_name=f'{COMMODITY}_Risk_Indicators'[:31], index=False)

    if solver_comparison:
        pd.DataFrame(solver_comparison).to_excel(w, sheet_name=f'{COMMODITY}_Solver_Comparison'[:31], index=False)

    config_df = pd.DataFrame({
        'Parameter': ['Commodity', 'Input_File', 'Training_Years', 'Test_Year', 'Alpha',
                      'Min_Training_Years', 'Min_Historical_Years', 'N_Countries',
                      'Total_Export_D', 'Emp_Variance', 'Emp_Return', 'Emp_HHI',
                      'Top_Partner', 'Top_Region', 'Gurobi_Available'],
        'Value': [COMMODITY, INPUT_FILE, f'{min(TRAIN_YEARS)}-{max(TRAIN_YEARS)}', TEST_YEAR, ALPHA,
                  MIN_TRAINING_YEARS, MIN_HISTORICAL_YEARS, N, D, emp_variance, emp_return, emp_hhi,
                  TOP_PARTNER, TOP_REGION, GUROBI_AVAILABLE]
    })
    config_df.to_excel(w, sheet_name=f'{COMMODITY}_Config'[:31], index=False)

\

if SAVE_FIGURES:
    COLORS = {'Empirical':'#e74c3c', 'Baseline':'#3498db', 'Extended':'#2ecc71',
              'SingleShock':'#9b59b6', 'Regional':'#f39c12', 'USExclude':'#1abc9c'}
    SHORT = {'Empirical':'Emp', 'Baseline':'Base', 'Extended':'Ext',
             'SingleShock':'Shock', 'Regional':'Region', 'USExclude':'NoUS'}

    def create_performance_figure(results, cap_name, ptype):
        \
        scen = [s for s in ['Empirical','Baseline','Extended','SingleShock','Regional','USExclude']
                if s in results and results[s].get('status') in ['Optimal', 'Empirical']]

        if len(scen) <= 1:
            \
            return None

        x = np.arange(len(scen))
        colors = [COLORS[s] for s in scen]

        fig, ax = plt.subplots(2, 2, figsize=(12, 9))
        for (m, lab, a, mul) in [('volatility','Volatility %',ax[0,0],100),
                                  ('weighted_risk','Weighted Risk',ax[0,1],1),
                                  ('hhi','HHI',ax[1,0],1),
                                  ('return','Return',ax[1,1],1)]:
            vals = [results[s][m]*mul for s in scen]
            bars = a.bar(x, vals, color=colors, edgecolor='black')
            a.set_ylabel(lab)
            a.set_xticks(x)
            a.set_xticklabels([SHORT[s] for s in scen])
            a.grid(axis='y', alpha=0.3)
            for b, v in zip(bars, vals):
                a.text(b.get_x()+b.get_width()/2, b.get_height(), f'{v:.3f}',
                       ha='center', va='bottom', fontsize=8)

        plt.suptitle(f'{COMMODITY}: Performance - {cap_name} ({ptype})', fontweight='bold')
        plt.tight_layout()

        fig_name = f"{COMMODITY}_Performance_{cap_name}_{ptype}.png"
        plt.savefig(os.path.join(OUTPUT_DIR, fig_name), dpi=FIGURE_DPI)
        plt.close()
        return fig_name

    def create_allocation_figure(results, cap_name, ptype):
        \
        scen = [s for s in ['Empirical','Baseline','Extended','SingleShock','Regional','USExclude']
                if s in results and results[s].get('status') in ['Optimal', 'Empirical']]

        if len(scen) <= 1:
            return None

        n_scen = len(scen)
        n_cols = min(3, n_scen)
        n_rows = (n_scen + n_cols - 1) // n_cols

        fig, axes = plt.subplots(n_rows, n_cols, figsize=(5*n_cols, 5*n_rows))
        if n_rows == 1 and n_cols == 1:
            axes = np.array([[axes]])
        elif n_rows == 1:
            axes = axes.reshape(1, -1)
        elif n_cols == 1:
            axes = axes.reshape(-1, 1)
        axes = axes.flatten()

        for i, s in enumerate(scen):
            ax = axes[i]
            wt = results[s]['weights']
            sidx = np.argsort(wt)[::-1][:8]
            tw = wt[sidx]
            tn = [all_countries[j][:12] for j in sidx]

            oth = 1 - tw.sum()
            if oth > 0.005:
                tw = np.append(tw, oth)
                tn.append('Others')

            mask = tw > 0.001
            tw, tn = tw[mask], [n for n, m in zip(tn, mask) if m]

            if len(tw) > 0:
                pc = plt.cm.Set3(np.linspace(0, 1, len(tw)))
                wedges, _, _ = ax.pie(tw, autopct=lambda p: f'{p:.1f}%' if p>2 else '',
                                       colors=pc, startangle=90)
                ax.legend(wedges, tn, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=8)
            ax.set_title(f"{s} (HHI={results[s]['hhi']:.3f})")

        for i in range(len(scen), len(axes)):
            axes[i].axis('off')

        plt.suptitle(f'{COMMODITY}: Allocations - {cap_name} ({ptype})', fontweight='bold')
        plt.tight_layout()

        fig_name = f"{COMMODITY}_Allocations_{cap_name}_{ptype}.png"
        plt.savefig(os.path.join(OUTPUT_DIR, fig_name), dpi=FIGURE_DPI)
        plt.close()
        return fig_name

    figures_created = []

    for cap_name in ['Conservative', 'Conservative_Alt', 'Expansion1', 'Expansion2']:
        if cap_name not in all_results:
            continue

        for ptype in ['Optimum', 'Optimal']:
            results = all_results[cap_name][ptype.lower()]

            fig = create_performance_figure(results, cap_name, ptype)
            if fig:
                figures_created.append(fig)
\

            fig = create_allocation_figure(results, cap_name, ptype)
            if fig:
                figures_created.append(fig)
\

\

\
for cap_name in ['Conservative', 'Conservative_Alt', 'Expansion1', 'Expansion2']:
    if cap_name not in all_results:
        continue

    results = all_results[cap_name]['optimum']
    scen = [s for s in ['Empirical','Baseline','Extended','SingleShock','Regional','USExclude']
            if s in results and results[s].get('status') in ['Optimal', 'Empirical']]

    if len(scen) <= 1:
        \
        continue

    \
\
    for s in scen:
        r = results[s]
        vc = (r['variance']-emp_variance)/emp_variance*100 if s != 'Empirical' else 0
