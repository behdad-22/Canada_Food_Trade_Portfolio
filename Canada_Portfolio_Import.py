import pandas as pd
import numpy as np
from scipy.optimize import minimize
import warnings
import os
import sys
warnings.filterwarnings('ignore')




GUROBI_AVAILABLE = False
try:
    from gurobipy import Model, GRB
    GUROBI_AVAILABLE = True
    print("✓ Gurobi available")
except ImportError:
    print("⚠ Gurobi not available - using scipy only")






DATA_DIR = r"C:\Behdad\Research\Model\R\Paper 3\Import_Code"

COMMODITY = "beetsugar"
INPUT_FILE = os.path.join(DATA_DIR, f"Imp_{COMMODITY}.xlsx")

SHEET_MASS = 'imp-mass'
SHEET_VALUE = 'imp_value'
SHEET_EXPORT_CAP = 'Maximum_export'

ALPHA = 0.5
MIN_TRAINING_YEARS = 3
TEST_YEAR = 2024

W_PS, W_CV, W_LPI, W_WSI = 0.299, 0.329, 0.328, 0.044


PS_FILE = os.path.join(DATA_DIR, "Political_Stability_Indicators.csv")
CV_FILE = os.path.join(DATA_DIR, "Climate_Vulnerability_indicator.csv")
LPI_FILE = os.path.join(DATA_DIR, "Logistic_Performance_Indicator.csv")
WSI_FILE = os.path.join(DATA_DIR, "Water_Stress_Indicator.csv")

OUTPUT_DIR = DATA_DIR


CANADIAN_PROVINCES = ['Ontario', 'Quebec', 'Manitoba', 'Alberta', 'Saskatchewan',
                      'British Columbia', 'Nova Scotia', 'New Brunswick',
                      'Prince Edward Island', 'Newfoundland']


CANADA_RISK = {
    'PS_risk': 0.336,
    'CV_risk': 0.282,
    'LPI_risk': 0.280,
    'WSI_risk': 0.247,
}





print("="*70)
print(f"CANADA {COMMODITY.upper()} IMPORT PORTFOLIO OPTIMIZATION")
print("With Domestic Capacity Scenario")
print("="*70)





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
    name = str(name).lower().strip()
    for c in candidates:
        if str(c).lower() == name:
            return c
    for c in candidates:
        if len(name) >= 5 and name[:5] in str(c).lower():
            return c
        if len(str(c)) >= 5 and str(c).lower()[:5] in name:
            return c
    return None

def is_canadian_province(name):
    name_lower = str(name).lower()
    for prov in CANADIAN_PROVINCES:
        if prov.lower() in name_lower:
            return True
    return False

def safe_load_csv(filepath, name, index_col=None):
    if not os.path.exists(filepath):
        print(f"⚠ {name}: not found")
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
        print(f"✓ {name}: loaded")
        return df
    except Exception as e:
        print(f"⚠ {name}: {e}")
        return None

REGIONS = {
    'North America': ['United States', 'Mexico', 'Puerto Rico'],
    'South America': ['Brazil', 'Argentina', 'Chile', 'Colombia', 'Peru', 'Ecuador', 'Paraguay', 'Uruguay'],
    'Europe': ['Germany', 'Netherlands', 'France', 'Belgium', 'Luxembourg', 'Italy',
               'Spain', 'Portugal', 'United Kingdom', 'Ireland', 'Poland', 'Denmark',
               'Bulgaria', 'Romania', 'Sweden', 'Norway', 'Finland', 'Austria', 'Hungary'],
    'Asia-Pacific': ['China', 'Japan', 'Korea', 'India', 'Indonesia', 'Thailand', 'Vietnam'],
}

def get_region(country):
    for r, kw in REGIONS.items():
        if any(k.lower() in str(country).lower() for k in kw):
            return r
    return 'Other'




print("\n" + "-"*70)
print("STEP 1: Loading Input Data")
print("-"*70)
if not os.path.exists(INPUT_FILE):
    print(f"\n❌ ERROR: File not found: {INPUT_FILE}")
    sys.exit(1)
xlsx = pd.ExcelFile(INPUT_FILE)
print(f"✓ File: {INPUT_FILE}")
print(f"  Sheets: {xlsx.sheet_names}")
imp_mass_raw = pd.read_excel(xlsx, sheet_name=SHEET_MASS, index_col=0)
imp_mass_raw = imp_mass_raw[imp_mass_raw.index.notna()]
imp_mass_raw.index = imp_mass_raw.index.astype(str).str.strip()
imp_value_raw = pd.read_excel(xlsx, sheet_name=SHEET_VALUE, index_col=0)
imp_value_raw = imp_value_raw[imp_value_raw.index.notna()]
imp_value_raw.index = imp_value_raw.index.astype(str).str.strip()




imp_mass_raw = imp_mass_raw.replace('..', np.nan)
imp_mass_raw = imp_mass_raw.replace(['', '-', 'x', 'X', '...'], np.nan)
imp_mass_raw = imp_mass_raw.apply(pd.to_numeric, errors='coerce')
imp_mass_raw = imp_mass_raw.fillna(0)


imp_value_raw = imp_value_raw.replace('..', np.nan)
imp_value_raw = imp_value_raw.replace(['', '-', 'x', 'X', '...'], np.nan)
imp_value_raw = imp_value_raw.apply(pd.to_numeric, errors='coerce')
imp_value_raw = imp_value_raw.fillna(0)


print(f"✓ {SHEET_MASS}: {imp_mass_raw.shape[0]} sources")
print(f"✓ {SHEET_VALUE}: {imp_value_raw.shape[0]} sources")

max_export = None
if SHEET_EXPORT_CAP in xlsx.sheet_names:
    max_export = pd.read_excel(xlsx, sheet_name=SHEET_EXPORT_CAP, index_col=0)
    max_export = max_export[max_export.index.notna()]
    max_export.index = max_export.index.astype(str).str.strip()

    max_export = max_export.replace('..', np.nan)
    max_export = max_export.replace(['', '-', 'x', 'X', '...'], np.nan)
    max_export = max_export.apply(pd.to_numeric, errors='coerce')
    max_export = max_export.fillna(0)

    print(f"✓ {SHEET_EXPORT_CAP}: {max_export.shape[0]} sources")

def get_year_columns(df):
    years = []
    for col in df.columns:
        try:
            y = int(str(col).strip())
            if 1900 < y < 2100:
                years.append(y)
        except:
            pass
    return sorted(years)
all_years = get_year_columns(imp_mass_raw)
TRAIN_YEARS = [y for y in all_years if y < TEST_YEAR]
if TEST_YEAR not in all_years:
    TEST_YEAR = max(all_years)
    TRAIN_YEARS = [y for y in all_years if y < TEST_YEAR]
print(f"✓ Years: {min(all_years)}-{max(all_years)}, Training: {len(TRAIN_YEARS)} years")





print("\n" + "-"*70)
print("STEP 2: Separating Foreign Sources and Canadian Provinces")
print("-"*70)


common_sources = sorted(list(set(imp_mass_raw.index) & set(imp_value_raw.index)))
imp_mass_aligned = imp_mass_raw.loc[common_sources].fillna(0)
imp_value_aligned = imp_value_raw.loc[common_sources].fillna(0)


foreign_sources = [s for s in common_sources if not is_canadian_province(s) and s != 'Canada']
domestic_sources = [s for s in common_sources if is_canadian_province(s)]

print(f"✓ Total sources: {len(common_sources)}")
print(f"✓ Foreign sources: {len(foreign_sources)}")
print(f"✓ Canadian provinces: {len(domestic_sources)}")
print(f"  Provinces: {domestic_sources}")


imp_mass_foreign = imp_mass_aligned.loc[foreign_sources][all_years].copy()
imp_value_foreign = imp_value_aligned.loc[foreign_sources][all_years].copy()

imp_mass_domestic = imp_mass_aligned.loc[domestic_sources][all_years].copy()
imp_value_domestic = imp_value_aligned.loc[domestic_sources][all_years].copy()





print("\n" + "-"*70)
print("STEP 3: Calculating Prices and Covariance (Foreign Sources)")
print("-"*70)


years_with_trade = (imp_mass_foreign[TRAIN_YEARS] > 0).sum(axis=1)
has_training = years_with_trade >= MIN_TRAINING_YEARS
has_test_year = imp_mass_foreign[TEST_YEAR] > 0
valid_mask = has_training & has_test_year

valid_foreign = imp_mass_foreign.index[valid_mask].tolist()
print(f"✓ Valid foreign sources (≥{MIN_TRAINING_YEARS} training years + {TEST_YEAR} trade): {len(valid_foreign)}")

imp_mass_f = imp_mass_foreign.loc[valid_foreign].copy()
imp_value_f = imp_value_foreign.loc[valid_foreign].copy()

N_foreign = len(valid_foreign)


imp_unit_price_f = imp_value_f / imp_mass_f.replace(0, np.nan)
train_prices_f = imp_unit_price_f[TRAIN_YEARS]

mean_prices_f = np.zeros(N_foreign)
for i, source in enumerate(valid_foreign):
    prices = train_prices_f.loc[source].values
    valid_p = prices[~np.isnan(prices) & (prices > 0)]
    mean_prices_f[i] = np.mean(valid_p) if len(valid_p) > 0 else 0


median_price = np.median(mean_prices_f[mean_prices_f > 0]) if any(mean_prices_f > 0) else 1.0
mean_prices_f[mean_prices_f == 0] = median_price

print(f"✓ Mean price range (foreign): {mean_prices_f.min():.4f} to {mean_prices_f.max():.4f}")


cov_matrix_f = np.zeros((N_foreign, N_foreign))
for i in range(N_foreign):
    for j in range(i, N_foreign):
        pi = train_prices_f.iloc[i].values
        pj = train_prices_f.iloc[j].values
        valid = (~np.isnan(pi)) & (~np.isnan(pj)) & (pi > 0) & (pj > 0)
        if valid.sum() >= 2:
            cov_val = np.cov(pi[valid], pj[valid], ddof=1)[0, 1]
            cov_matrix_f[i, j] = cov_val
            cov_matrix_f[j, i] = cov_val
        elif i == j:
            pi_valid = pi[~np.isnan(pi) & (pi > 0)]
            cov_matrix_f[i, i] = np.var(pi_valid, ddof=1) if len(pi_valid) >= 2 else 0.001


eigvals, eigvecs = np.linalg.eigh(cov_matrix_f)
eigvals = np.maximum(eigvals, 1e-10)
cov_matrix_f = eigvecs @ np.diag(eigvals) @ eigvecs.T
cov_matrix_f = (cov_matrix_f + cov_matrix_f.T) / 2

print(f"✓ Covariance matrix (foreign): {N_foreign}×{N_foreign}")





print("\n" + "-"*70)
print(f"STEP 4: Empirical {TEST_YEAR} Allocation (Foreign Only)")
print("-"*70)

imp_mass_2024_f = imp_mass_f[TEST_YEAR].values
total_imports_2024 = imp_mass_2024_f.sum()
D_foreign = total_imports_2024

empirical_weights_f = imp_mass_2024_f / total_imports_2024

emp_variance_f = empirical_weights_f @ cov_matrix_f @ empirical_weights_f
emp_volatility_f = np.sqrt(emp_variance_f)
emp_return_f = empirical_weights_f @ mean_prices_f
emp_hhi_f = np.sum(empirical_weights_f ** 2)

print(f"✓ Total foreign imports D: {D_foreign:,.0f} tonnes")
print(f"✓ Empirical (foreign): Var={emp_variance_f:.6f}, Vol={emp_volatility_f*100:.2f}%, HHI={emp_hhi_f:.4f}")


print(f"\nTop 10 foreign sources by {TEST_YEAR}:")
sorted_idx_f = np.argsort(empirical_weights_f)[::-1]
for rank, i in enumerate(sorted_idx_f[:10]):
    if empirical_weights_f[i] > 0.001:
        print(f"  {rank+1}. {valid_foreign[i][:25]}: {empirical_weights_f[i]*100:.1f}%")





print("\n" + "-"*70)
print("STEP 5: Domestic Production (Canadian Provinces)")
print("-"*70)


domestic_production_2024 = {}
domestic_unit_price_2024 = {}

for prov in domestic_sources:
    mass_2024 = imp_mass_domestic.loc[prov, TEST_YEAR] if TEST_YEAR in imp_mass_domestic.columns else 0
    value_2024 = imp_value_domestic.loc[prov, TEST_YEAR] if TEST_YEAR in imp_value_domestic.columns else 0

    if pd.notna(mass_2024) and mass_2024 > 0:
        domestic_production_2024[prov] = mass_2024
        if pd.notna(value_2024) and value_2024 > 0:

            domestic_unit_price_2024[prov] = value_2024 / mass_2024
        else:
            domestic_unit_price_2024[prov] = median_price
        print(f"  {prov}: {mass_2024:,.0f} tonnes, ${domestic_unit_price_2024[prov]:.2f}/tonne")

total_domestic_2024 = sum(domestic_production_2024.values())
total_consumption_2024 = total_imports_2024 + total_domestic_2024

print(f"\n✓ Total domestic production: {total_domestic_2024:,.0f} tonnes")
print(f"✓ Total foreign imports: {total_imports_2024:,.0f} tonnes")
print(f"✓ TOTAL CONSUMPTION (D): {total_consumption_2024:,.0f} tonnes")


domestic_share = total_domestic_2024 / total_consumption_2024 * 100
import_share = total_imports_2024 / total_consumption_2024 * 100
print(f"\n✓ Current mix: {domestic_share:.1f}% domestic, {import_share:.1f}% imports")





print("\n" + "-"*70)
print("STEP 6: Loading Risk Indicators")
print("-"*70)

ps_df = safe_load_csv(PS_FILE, "PS")
cv_df = safe_load_csv(CV_FILE, "CV", index_col=0)
lpi_df = safe_load_csv(LPI_FILE, "LPI")
wsi_df = safe_load_csv(WSI_FILE, "WSI")


risk_data_f = []
for source in valid_foreign:
    mapped = get_mapped_name(source)
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
            for col in ['2022', '2021', '2020']:
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

    risk_data_f.append({'Source': source, 'PS_raw': ps, 'CV_raw': cv, 'LPI_raw': lpi, 'WSI_raw': wsi})

risk_df_f = pd.DataFrame(risk_data_f)
risk_df_f.set_index('Source', inplace=True)

risk_df_f['PS_risk'] = risk_df_f['PS_raw'].apply(lambda x: (2.5-x)/5.0 if pd.notna(x) else None)
risk_df_f['CV_risk'] = risk_df_f['CV_raw']
risk_df_f['LPI_risk'] = risk_df_f['LPI_raw'].apply(lambda x: (5-x)/4.0 if pd.notna(x) else None)
risk_df_f['WSI_risk'] = risk_df_f['WSI_raw']

for col in ['PS_risk', 'CV_risk', 'LPI_risk', 'WSI_risk']:
    med = risk_df_f[col].median()
    risk_df_f[col] = risk_df_f[col].fillna(med if pd.notna(med) else 0.5)

risk_df_f['RI'] = W_PS*risk_df_f['PS_risk'] + W_CV*risk_df_f['CV_risk'] + W_LPI*risk_df_f['LPI_risk'] + W_WSI*risk_df_f['WSI_risk']
risk_df_f['k_factor'] = 1 + ALPHA * risk_df_f['RI']

RI_f = risk_df_f['RI'].values
k_factors_f = risk_df_f['k_factor'].values

print(f"✓ RI range (foreign): {RI_f.min():.3f} to {RI_f.max():.3f}")


RI_canada = W_PS*CANADA_RISK['PS_risk'] + W_CV*CANADA_RISK['CV_risk'] + W_LPI*CANADA_RISK['LPI_risk'] + W_WSI*CANADA_RISK['WSI_risk']
print(f"✓ Canada RI: {RI_canada:.3f}")





print("\n" + "-"*70)
print("STEP 7: Defining Scenarios")
print("-"*70)


top_idx_f = np.argmax(empirical_weights_f)
TOP_SOURCE = valid_foreign[top_idx_f]
print(f"✓ Top foreign source: {TOP_SOURCE} ({empirical_weights_f[top_idx_f]*100:.1f}%)")


region_shares = {}
for r in list(REGIONS.keys()) + ['Other']:
    s = sum(empirical_weights_f[i] for i, c in enumerate(valid_foreign) if get_region(c) == r)
    if s > 0.001:
        region_shares[r] = s

if region_shares:
    TOP_REGION = max(region_shares, key=region_shares.get)
    TOP_REGION_COUNTRIES = [c for c in valid_foreign if get_region(c) == TOP_REGION]
    print(f"✓ Top region: {TOP_REGION} ({len(TOP_REGION_COUNTRIES)} sources, {region_shares[TOP_REGION]*100:.1f}%)")
else:
    TOP_REGION = None
    TOP_REGION_COUNTRIES = []


US_SOURCE = next((c for c in valid_foreign if 'United States' in c), None)
PR_SOURCE = next((c for c in valid_foreign if 'Puerto Rico' in c), None)
US_EXCLUSION_LIST = [c for c in [US_SOURCE, PR_SOURCE] if c]

if US_SOURCE:
    us_idx = valid_foreign.index(US_SOURCE)
    print(f"✓ US: {US_SOURCE} ({empirical_weights_f[us_idx]*100:.1f}%)")
if PR_SOURCE:
    pr_idx = valid_foreign.index(PR_SOURCE)
    print(f"✓ Puerto Rico: {PR_SOURCE} ({empirical_weights_f[pr_idx]*100:.1f}%)")





print("\n" + "-"*70)
print("STEP 8: Setting Up Capacity")
print("-"*70)


capacity_foreign = imp_mass_f[TRAIN_YEARS + [TEST_YEAR]].max(axis=1).values


if max_export is not None:
    for i, source in enumerate(valid_foreign):
        if source in max_export.index:
            try:
                cap_val = float(max_export.loc[source, 'Value'])
                if pd.notna(cap_val) and cap_val > 0:
                    capacity_foreign[i] = max(capacity_foreign[i], cap_val)
            except:
                pass

print(f"✓ Foreign capacity range: {capacity_foreign.min():,.0f} to {capacity_foreign.max():,.0f} tonnes")





def optimize_scipy(alpha, capacity, sources, mean_prices, cov_matrix, RI, D, exclusions=None, min_return=None):
    if exclusions is None:
        exclusions = []

    active = np.array([s not in exclusions for s in sources])
    idx = np.where(active)[0]
    n = len(idx)

    if n == 0:
        return {'status': 'No sources'}

    mu = mean_prices[idx]
    cov = cov_matrix[np.ix_(idx, idx)]
    cap = capacity[idx]
    ri = RI[idx]

    k = (1 + alpha * ri) if alpha > 0 else np.ones(n)
    K = np.diag(k.astype(float))
    cov_adj = K @ cov.astype(float) @ K

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

            weights = np.zeros(len(sources))
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





print("\n" + "-"*70)
print("STEP 9: Running Scenarios 1-5 (Foreign Sources Only)")
print("-"*70)

emp_result_f = {
    'weights': empirical_weights_f,
    'variance': emp_variance_f,
    'volatility': emp_volatility_f,
    'return': emp_return_f,
    'hhi': emp_hhi_f,
    'weighted_risk': np.sum(empirical_weights_f * RI_f),
    'n_active': int(np.sum(empirical_weights_f > 0.001)),
    'status': 'Empirical',
    'solver': 'N/A'
}

scenarios_foreign = [
    ('Baseline', 0, []),
    ('Extended', ALPHA, []),
    ('SingleShock', ALPHA, [TOP_SOURCE]),
    ('Regional', ALPHA, TOP_REGION_COUNTRIES),
    ('USExclude', ALPHA, US_EXCLUSION_LIST),
]

results_foreign = {'optimum': {'Empirical': emp_result_f}, 'optimal': {'Empirical': emp_result_f}}

print(f"\n--- OPTIMUM (No Return Constraint) ---")
for name, alpha, excl in scenarios_foreign:
    result = optimize_scipy(alpha, capacity_foreign, valid_foreign, mean_prices_f,
                           cov_matrix_f, RI_f, D_foreign, excl, None)

    if result.get('status') == 'Optimal':
        results_foreign['optimum'][name] = result
        vc = (result['variance'] - emp_variance_f) / emp_variance_f * 100
        print(f"  {name:<12}: Var={result['variance']:.6f}({vc:+.1f}%) [{result['solver']}]")
    else:
        print(f"  {name:<12}: {result.get('status')}")

print(f"\n--- OPTIMAL (Return ≥ {emp_return_f:.4f}) ---")
for name, alpha, excl in scenarios_foreign:
    result = optimize_scipy(alpha, capacity_foreign, valid_foreign, mean_prices_f,
                           cov_matrix_f, RI_f, D_foreign, excl, emp_return_f)

    if result.get('status') == 'Optimal':
        results_foreign['optimal'][name] = result
        vc = (result['variance'] - emp_variance_f) / emp_variance_f * 100
        print(f"  {name:<12}: Var={result['variance']:.6f}({vc:+.1f}%) [{result['solver']}]")
    else:
        print(f"  {name:<12}: {result.get('status')}")





print("\n" + "-"*70)
print("STEP 10: Scenario 6 - DomesticCapacity (Including Canadian Provinces)")
print("-"*70)


valid_provinces = list(domestic_production_2024.keys())
all_sources = valid_foreign + valid_provinces
N_total = len(all_sources)

print(f"✓ Total sources (foreign + domestic): {N_total}")
print(f"  Foreign: {len(valid_foreign)}, Domestic: {len(valid_provinces)}")



capacity_total = np.zeros(N_total)
capacity_total[:N_foreign] = capacity_foreign
for i, prov in enumerate(valid_provinces):
    capacity_total[N_foreign + i] = domestic_production_2024[prov]


mean_prices_total = np.zeros(N_total)
mean_prices_total[:N_foreign] = mean_prices_f


if US_SOURCE and US_SOURCE in valid_foreign:
    us_idx_price = valid_foreign.index(US_SOURCE)
    us_price = mean_prices_f[us_idx_price]
else:
    us_price = median_price

for i, prov in enumerate(valid_provinces):

    mean_prices_total[N_foreign + i] = domestic_unit_price_2024.get(prov, us_price)


RI_total = np.zeros(N_total)
RI_total[:N_foreign] = RI_f
for i in range(len(valid_provinces)):
    RI_total[N_foreign + i] = RI_canada



cov_matrix_total = np.zeros((N_total, N_total))
cov_matrix_total[:N_foreign, :N_foreign] = cov_matrix_f


if US_SOURCE and US_SOURCE in valid_foreign:
    us_idx_cov = valid_foreign.index(US_SOURCE)
    us_var = cov_matrix_f[us_idx_cov, us_idx_cov]
else:
    us_var = np.mean(np.diag(cov_matrix_f))

for i in range(len(valid_provinces)):
    prov_idx = N_foreign + i

    cov_matrix_total[prov_idx, prov_idx] = us_var

    for j in range(N_foreign):
        if valid_foreign[j] == US_SOURCE:
            corr = 0.95
        else:
            corr = 0.5
        cov_matrix_total[prov_idx, j] = corr * np.sqrt(us_var * cov_matrix_f[j, j])
        cov_matrix_total[j, prov_idx] = cov_matrix_total[prov_idx, j]

    for j in range(len(valid_provinces)):
        if i != j:
            other_idx = N_foreign + j
            cov_matrix_total[prov_idx, other_idx] = 0.95 * us_var
            cov_matrix_total[other_idx, prov_idx] = cov_matrix_total[prov_idx, other_idx]


eigvals, eigvecs = np.linalg.eigh(cov_matrix_total)
eigvals = np.maximum(eigvals, 1e-10)
cov_matrix_total = eigvecs @ np.diag(eigvals) @ eigvecs.T
cov_matrix_total = (cov_matrix_total + cov_matrix_total.T) / 2


D_total = total_consumption_2024


empirical_mass_total = np.zeros(N_total)
empirical_mass_total[:N_foreign] = imp_mass_2024_f
for i, prov in enumerate(valid_provinces):
    empirical_mass_total[N_foreign + i] = domestic_production_2024[prov]

empirical_weights_total = empirical_mass_total / D_total

emp_variance_total = empirical_weights_total @ cov_matrix_total @ empirical_weights_total
emp_return_total = empirical_weights_total @ mean_prices_total
emp_hhi_total = np.sum(empirical_weights_total ** 2)
emp_weighted_risk_total = np.sum(empirical_weights_total * RI_total)

print(f"\n✓ Total D (consumption): {D_total:,.0f} tonnes")
print(f"✓ Empirical (combined): Var={emp_variance_total:.6f}, HHI={emp_hhi_total:.4f}")

emp_result_total = {
    'weights': empirical_weights_total,
    'variance': emp_variance_total,
    'volatility': np.sqrt(emp_variance_total),
    'return': emp_return_total,
    'hhi': emp_hhi_total,
    'weighted_risk': emp_weighted_risk_total,
    'n_active': int(np.sum(empirical_weights_total > 0.001)),
    'status': 'Empirical',
    'solver': 'N/A'
}


results_domestic = {'optimum': {'Empirical': emp_result_total}, 'optimal': {'Empirical': emp_result_total}}

print(f"\n--- OPTIMUM (DomesticCapacity) ---")
result = optimize_scipy(ALPHA, capacity_total, all_sources, mean_prices_total,
                       cov_matrix_total, RI_total, D_total, [], None)

if result.get('status') == 'Optimal':
    results_domestic['optimum']['DomesticCapacity'] = result
    vc = (result['variance'] - emp_variance_total) / emp_variance_total * 100
    print(f"  DomesticCapacity: Var={result['variance']:.6f}({vc:+.1f}%)")


    print(f"\n  Allocation:")
    wt = result['weights']

    foreign_share = np.sum(wt[:N_foreign]) * 100
    domestic_share = np.sum(wt[N_foreign:]) * 100
    print(f"    Foreign: {foreign_share:.1f}%")
    print(f"    Domestic: {domestic_share:.1f}%")


    sorted_idx = np.argsort(wt)[::-1]
    print(f"\n  Top 10 allocations:")
    for rank, i in enumerate(sorted_idx[:10]):
        if wt[i] > 0.001:
            source = all_sources[i]
            is_domestic = "🍁" if i >= N_foreign else ""
            print(f"    {rank+1}. {source[:25]}: {wt[i]*100:.1f}% {is_domestic}")
else:
    print(f"  DomesticCapacity: {result.get('status')}")

print(f"\n--- OPTIMAL (DomesticCapacity, Return ≥ {emp_return_total:.4f}) ---")
result = optimize_scipy(ALPHA, capacity_total, all_sources, mean_prices_total,
                       cov_matrix_total, RI_total, D_total, [], emp_return_total)

if result.get('status') == 'Optimal':
    results_domestic['optimal']['DomesticCapacity'] = result
    vc = (result['variance'] - emp_variance_total) / emp_variance_total * 100
    print(f"  DomesticCapacity: Var={result['variance']:.6f}({vc:+.1f}%)")
else:
    print(f"  DomesticCapacity: {result.get('status')}")





print("\n" + "-"*70)
print("STEP 11: Opportunity Cost Analysis")
print("-"*70)

print("""
When Canadian provinces supply corn to other regions of Canada instead of
using it locally, there may be an OPPORTUNITY COST - the benefit foregone.

This analysis calculates:
1. Current allocation (empirical)
2. Optimized allocation (DomesticCapacity)
3. Change in domestic vs foreign sourcing
4. Implied opportunity cost if provinces redirect production
""")

if 'DomesticCapacity' in results_domestic['optimum']:
    opt_result = results_domestic['optimum']['DomesticCapacity']
    wt_opt = opt_result['weights']


    emp_foreign_share = np.sum(empirical_weights_total[:N_foreign]) * 100
    emp_domestic_share = np.sum(empirical_weights_total[N_foreign:]) * 100

    opt_foreign_share = np.sum(wt_opt[:N_foreign]) * 100
    opt_domestic_share = np.sum(wt_opt[N_foreign:]) * 100

    print(f"\n{'Source Type':<20} {'Empirical':>12} {'Optimized':>12} {'Change':>12}")
    print("-"*56)
    print(f"{'Foreign':<20} {emp_foreign_share:>11.1f}% {opt_foreign_share:>11.1f}% {opt_foreign_share-emp_foreign_share:>+11.1f}%")
    print(f"{'Domestic (CA)':<20} {emp_domestic_share:>11.1f}% {opt_domestic_share:>11.1f}% {opt_domestic_share-emp_domestic_share:>+11.1f}%")


    print(f"\n{'Province':<20} {'Empirical':>15} {'Optimized':>15} {'Change':>15}")
    print("-"*65)
    for i, prov in enumerate(valid_provinces):
        emp_prov = empirical_weights_total[N_foreign + i] * D_total
        opt_prov = wt_opt[N_foreign + i] * D_total
        change = opt_prov - emp_prov
        print(f"{prov:<20} {emp_prov:>14,.0f} {opt_prov:>14,.0f} {change:>+14,.0f}")


    print(f"\n" + "="*70)
    print("OPPORTUNITY COST CALCULATION")
    print("="*70)




    for i, prov in enumerate(valid_provinces):
        emp_vol = empirical_weights_total[N_foreign + i] * D_total
        opt_vol = wt_opt[N_foreign + i] * D_total

        if opt_vol > emp_vol:
            redirected = opt_vol - emp_vol
            local_price = domestic_unit_price_2024.get(prov, us_price)



            opp_cost = redirected * local_price / 1000

            print(f"\n{prov}:")
            print(f"  Redirected volume: {redirected:,.0f} tonnes")
            print(f"  Local price: ${local_price:.2f}/tonne")
            print(f"  Opportunity cost: ${opp_cost:,.0f} thousand")
            print(f"  (This is the value foregone if {prov} redirects corn to other provinces)")





print("\n" + "-"*70)
print("STEP 12: Saving Results")
print("-"*70)

output_file = os.path.join(OUTPUT_DIR, f"{COMMODITY}_Import_Results.xlsx")

def make_summary(results, ptype, sources, emp_var, emp_ret):
    rows = []
    for name in ['Empirical', 'Baseline', 'Extended', 'SingleShock', 'Regional', 'USExclude', 'DomesticCapacity']:
        if name in results and results[name].get('status') in ['Optimal', 'Empirical']:
            r = results[name]
            rows.append({
                'Scenario': name, 'Type': ptype, 'N_Active': r['n_active'],
                'Variance': r['variance'], 'Var_Change_%': (r['variance']-emp_var)/emp_var*100,
                'Volatility_%': r['volatility']*100, 'Return': r['return'],
                'Return_Change_%': (r['return']-emp_ret)/emp_ret*100 if emp_ret > 0 else 0,
                'HHI': r['hhi'], 'Weighted_Risk': r['weighted_risk'],
                'Solver': r.get('solver', 'N/A')
            })
    return pd.DataFrame(rows)

with pd.ExcelWriter(output_file, engine='openpyxl') as w:

    df_optimum_f = make_summary(results_foreign['optimum'], 'Optimum', valid_foreign, emp_variance_f, emp_return_f)
    df_optimal_f = make_summary(results_foreign['optimal'], 'Optimal', valid_foreign, emp_variance_f, emp_return_f)

    df_optimum_f.to_excel(w, sheet_name=f'{COMMODITY}_Foreign_Optimum', index=False)
    df_optimal_f.to_excel(w, sheet_name=f'{COMMODITY}_Foreign_Optimal', index=False)


    df_optimum_d = make_summary(results_domestic['optimum'], 'Optimum', all_sources, emp_variance_total, emp_return_total)
    df_optimal_d = make_summary(results_domestic['optimal'], 'Optimal', all_sources, emp_variance_total, emp_return_total)

    df_optimum_d.to_excel(w, sheet_name=f'{COMMODITY}_Domestic_Optimum', index=False)
    df_optimal_d.to_excel(w, sheet_name=f'{COMMODITY}_Domestic_Optimal', index=False)


    alloc_f = {'Source': valid_foreign, 'RI': RI_f, 'Mean_Price': mean_prices_f, 'Capacity': capacity_foreign}
    for name, r in results_foreign['optimum'].items():
        if r.get('status') in ['Optimal', 'Empirical']:
            alloc_f[f'{name}_%'] = r['weights'] * 100
    alloc_df_f = pd.DataFrame(alloc_f)
    alloc_df_f = alloc_df_f.sort_values('Empirical_%', ascending=False)
    alloc_df_f.to_excel(w, sheet_name=f'{COMMODITY}_Alloc_Foreign', index=False)


    if 'DomesticCapacity' in results_domestic['optimum']:
        alloc_d = {
            'Source': all_sources,
            'Type': ['Foreign']*N_foreign + ['Domestic']*len(valid_provinces),
            'RI': RI_total,
            'Mean_Price': mean_prices_total,
            'Capacity': capacity_total,
            'Empirical_%': empirical_weights_total * 100,
            'DomesticCapacity_%': results_domestic['optimum']['DomesticCapacity']['weights'] * 100
        }
        alloc_df_d = pd.DataFrame(alloc_d)
        alloc_df_d = alloc_df_d.sort_values('Empirical_%', ascending=False)
        alloc_df_d.to_excel(w, sheet_name=f'{COMMODITY}_Alloc_Domestic', index=False)


    risk_df_f.to_excel(w, sheet_name=f'{COMMODITY}_Risk_Indicators')


    config_df = pd.DataFrame({
        'Parameter': ['Commodity', 'Input_File', 'Training_Years', 'Test_Year', 'Alpha',
                      'N_Foreign_Sources', 'N_Domestic_Provinces', 'Total_Imports',
                      'Total_Domestic', 'Total_Consumption', 'Canada_RI',
                      'W_PS', 'W_CV', 'W_LPI', 'W_WSI'],
        'Value': [COMMODITY, INPUT_FILE, f'{min(TRAIN_YEARS)}-{max(TRAIN_YEARS)}', TEST_YEAR, ALPHA,
                  N_foreign, len(valid_provinces), total_imports_2024,
                  total_domestic_2024, total_consumption_2024, RI_canada,
                  W_PS, W_CV, W_LPI, W_WSI]
    })
    config_df.to_excel(w, sheet_name=f'{COMMODITY}_Config', index=False)

print(f"✓ Saved: {output_file}")





print("\n" + "="*70)
print(f"{COMMODITY.upper()} IMPORT ANALYSIS COMPLETE")
print("="*70)

print(f"\n--- SCENARIOS 1-5 (Foreign Only) ---")
print(f"{'Scenario':<12} {'Variance':>10} {'Var%':>8} {'HHI':>8}")
print("-"*45)
for s in ['Empirical', 'Baseline', 'Extended', 'SingleShock', 'Regional', 'USExclude']:
    if s in results_foreign['optimum']:
        r = results_foreign['optimum'][s]
        vc = (r['variance']-emp_variance_f)/emp_variance_f*100 if s != 'Empirical' else 0
        print(f"{s:<12} {r['variance']:>10.6f} {vc:>+7.1f}% {r['hhi']:>8.4f}")

print(f"\n--- SCENARIO 6 (DomesticCapacity) ---")
if 'DomesticCapacity' in results_domestic['optimum']:
    r = results_domestic['optimum']['DomesticCapacity']
    vc = (r['variance']-emp_variance_total)/emp_variance_total*100
    print(f"Variance: {r['variance']:.6f} ({vc:+.1f}%)")
    print(f"Foreign share: {opt_foreign_share:.1f}%")
    print(f"Domestic share: {opt_domestic_share:.1f}%")

print(f"\n✓ Output file: {output_file}")
print("="*70)
print("DONE!")
print("="*70)
