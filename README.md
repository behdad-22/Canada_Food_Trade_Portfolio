# Agricultural Trade Portfolio Optimization under Multi-Dimensional Disruption Risks

Analysis code and data for the manuscript "Agricultural Trade Portfolio Optimization under Multi-Dimensional Disruption Risks: Enhancing Resilience of Canada's Food System" (Saed, Karakoc, Elshorbagy, Razavi), submitted to *Earth's Future*.

**Repository:** [https://github.com/behdad-22/Canada_Food_Trade_Portfolio]
**Archived version:** [(https://doi.org/10.5281/zenodo.21271392)]
**Corresponding author:** Behdad Saed, University of Saskatchewan.

## Contents

### Code (Python 3)

**1. `Canada_Import_Portfolio.py`**
Import portfolio optimization (maize, beet sugar). Set `Commodity` in the configuration block, then run. Writes `{commodity}_Import_Results.xlsx`.

**2. `Canada_Portfolio_Export.py`**
Export portfolio optimization (canola, wheat, soybean, colza, cattle). Same pattern, writes `{commodity}_Export_Results.xlsx`.

**3. `Sensitivity_analysis.py`**
Self-contained fine sensitivity sweep of the risk-aversion parameter (alpha from 0 to 1 in steps of 0.05, 21 values). Produces Figure 5 of the manuscript and the CSV files behind SI Table S5 and SI Figure S9.

### Data

**4. Trade workbooks.** `Imp_Maize.xlsx`, `Imp_Beetsugar.xlsx`, `Exp_Canola.xlsx`, `Exp_Wheat.xlsx`, `Exp_Soybean.xlsx`, `Exp_Colza.xlsx`, `Exp_Cattle.xlsx`. Gross bilateral flows (mass and value, 1986-2024) and partner capacity sheets.

**5. Risk indicators.** `Political_Stability_Indicators.csv`, `Climate_Vulnerability_indicator.csv`, `Logistic_Performance_Indicator.csv`, `Water_Stress_Indicator.csv`.

**6. `RI_values_used_full.csv`.** The composite Risk Index per partner as used in the analysis, an input to the sensitivity analysis code.

The data are compiled from FAO, UN Comtrade, the World Bank (WGI and LPI), ND-GAIN, and WRI Aqueduct. Full details are given in the manuscript's Methodology.

### Outputs

The result workbooks and CSV files listed above. The scripts also print summary results to the terminal as they run.

## Requirements

Python 3.10 or newer with `numpy`, `pandas`, `scipy`, `openpyxl`, and `matplotlib`.

A Gurobi (`gurobipy`) license is not required to run the scripts or to reproduce the results. If Gurobi is installed but its license has expired or is invalid, the export script may raise a license error, in that case uninstall `gurobipy` or disable the Gurobi branch and the script will use SciPy. See the solver note below.

## How to run

Place the relevant input files in the same folder as the script and run it from that folder.

```
python Canada_Import_Portfolio.py
python Canada_Portfolio_Export.py
python Sensitivity_analysis.py
```

For the two main scripts, set `Commodity` in the configuration block at the top to select the commodity. The sensitivity analysis script requires the six crop workbooks and `RI_values_used_full.csv` beside it.

## Which result is reported, "Optimal" not "Optimum"

Each main script writes two solutions.

The **Optimal** is the minimum-price-volatility portfolio subject to the constraint that the portfolio's expected cost (imports) or revenue (exports) is preserved at the empirical level. This keeps the optimized portfolio economically comparable with the one Canada actually held, so it is the realistic solution. All portfolios, reductions, and allocations reported in the manuscript are the Optimal.

The **Optimum** is the pure minimum-price-volatility portfolio with no economic target. It shows the lowest volatility technically attainable but may be unrealistic or infeasible in some cases, so it is a reference point only and is not reported. Reviewers reproducing the numbers should read the Optimal sheets.

## Solvers, and the Gurobi robustness check

The optimization is a quadratic program solved with SciPy's Sequential Least Squares Programming (SLSQP). Because SLSQP is a local solver, the export model was additionally solved with the Gurobi optimizer as an independent check on solution robustness, and the two produced identical optimal allocations in every run. Gurobi therefore only confirms that the SciPy solution is not a local artifact, it is not needed to reproduce the results. The Gurobi call runs automatically when `gurobipy` and a valid license are present and is skipped otherwise, in which case the SciPy result is used. Any valid Gurobi license, including a free academic or restricted license, is sufficient, and its start-up banner is informational and does not affect the output.

## Mapping of code to manuscript artifacts

Figures 1 to 4, Table 2, SI Tables S2 to S4, and SI Figures S1 to S7, from the import and export scripts' output workbooks. Figure 5, SI Table S5 (the five crop rows), SI Figure S8, and SI Figure S9, from the sensitivity script and its CSV outputs. Beet sugar's row in SI Table S5 is computed with the main import optimization, as stated in the table caption, the sensitivity script reproduces the main optimizations for the five plotted commodities. Table 1 lists the PCA-derived indicator weights, which are set as constants in each script.

## License and citation

Code released under the MIT License. If you use this repository, please cite the manuscript and the archived version above. Code is released under the MIT License. Data files are released under CC BY 4.0.
