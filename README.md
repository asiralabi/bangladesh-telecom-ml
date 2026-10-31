# Bangladesh Telecom Market — Data Science & AI Project

## Overview

End-to-end data science project on the Bangladesh mobile-telecom subscriber dataset (2008–2015). Covers exploratory data analysis, preprocessing, regression modelling, unsupervised learning, and a short machine-unlearning demonstration. The full workflow lives in a single Jupyter notebook.

**Course:** Data Science and AI Project (TIU, Spring 2026, B3E6020101)
**Instructor:** DUBEY Priya

## Dataset

- **File:** `GP Model.xlsx - Telecom Mkt. in Bangladesh.csv`
- **Source:** Internal modelling sheet on the Bangladeshi mobile market
- **Operators tracked:** Grameenphone, Banglalink, Robi, Airtel, Citycell, Teletalk, plus the Sector total
- **Coverage:** Monthly subscriber counts (in millions) from Jan-2008 through 2015, including some forward projections
- **Raw shape:** Wide layout, with months as columns and operators as rows — reshaped during preprocessing into a tidy long format (~600 rows after melting)

### Variables after reshaping

| Column | Description |
| --- | --- |
| `date` | Month/year of the observation |
| `year`, `month` | Calendar fields derived from `date` |
| `operator` | Telecom operator (categorical) |
| `subscribers_mn` | Subscriber count in millions (regression target) |
| `net_addition` | Month-on-month net subscriber additions |
| `growth_rate` | Subscriber growth rate (%) |
| `sector_total` | Total sector subscribers (millions) |
| `market_share` | Operator share of the sector total |
| `time_index` | Months since Jan-2008 (numeric trend feature) |

## Project Goals

1. **EDA** — understand subscriber trends, market dynamics, distributions, correlations and outliers
2. **Preprocessing** — reshape the wide CSV to tidy form, handle missing values, encode categoricals, scale numeric features
3. **Regression** — predict monthly subscriber counts using Linear, Ridge and Lasso regression
4. **Unsupervised learning** — cluster operator-month behaviour with KMeans; reduce dimensionality with PCA
5. **Machine unlearning** — frame an industry scenario (a telecom-analytics consultancy honouring an Airtel data-deletion request), train the production AI model on the full dataset, then make the model *forget* every Airtel row and verify both that the forget set is no longer well-predicted and that utility on the remaining operators is preserved

## Repository Structure

```
.
├── README.md                                       # This file
├── Telecom_BD_Project.ipynb                        # Single end-to-end notebook
├── GP Model.xlsx - Telecom Mkt. in Bangladesh.csv  # Raw dataset
└── Data science and AI project.pdf                 # Course syllabus
```

## Running the Notebook

Requirements: Python 3.9+ with the standard data-science stack.

```bash
pip install pandas numpy matplotlib seaborn scikit-learn jupyter
jupyter notebook Telecom_BD_Project.ipynb
```

Run the cells in order. The notebook reads the CSV from the same directory, so keep both files together.

## Notebook Outline

1. **Imports & configuration**
2. **Raw data loading** — read the wide CSV and locate the relevant rows
3. **Reshaping & preprocessing** — melt to long format, parse dates, type-cast, impute missing values
4. **Exploratory Data Analysis** — descriptive statistics, time-series plots, market-share evolution, correlation heatmap, outlier check via boxplots
5. **Feature engineering** — calendar features, lag features, operator one-hot encoding
6. **Regression modelling** — train/test split, scaling, Linear vs Ridge vs Lasso, metrics (R², MAE, RMSE), residual analysis, coefficient interpretation
7. **Unsupervised learning** — KMeans clustering of operator-month behaviour, elbow analysis, PCA projection for visualisation
8. **Machine unlearning** — define a "forget set", retrain a regression model on the remaining data, compare predictions and coefficients with the original model
9. **Findings & limitations**

## Key Findings (summary)

- Grameenphone has consistently held ~40–46% of the market; the sector roughly doubled in subscribers between 2008 and 2015.
- Ridge regression generalises slightly better than plain OLS once operator dummies are included, and Lasso usefully drives small coefficients to zero — a clean illustration of regularisation trade-offs.
- KMeans on standardised operator-month features cleanly separates dominant operators (Grameen) from challenger and niche operators (Citycell, Teletalk), matching domain intuition.
- The machine-unlearning experiment (Airtel data-deletion scenario) cleanly preserved accuracy on the other operators while destroying the model's ability to track Airtel — a textbook illustration of the "exact unlearning by retraining" baseline against which approximate methods (SISA, influence functions) are compared.

## Limitations

- The dataset is relatively small once aggregated, so model performance is reported as illustrative rather than production-grade.
- Some columns (net additions, growth rates) contain forward-looking assumption values; we treat only the historical block as ground truth.
- "Machine unlearning" is demonstrated via the simple retrain-from-scratch baseline rather than approximate-unlearning methods (e.g., influence functions or SISA), which are out of scope for this undergraduate project.
