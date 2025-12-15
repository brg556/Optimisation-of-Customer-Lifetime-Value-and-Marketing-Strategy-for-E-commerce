# Global Superstore: From Raw Data to Prescriptive Strategy

## Project Description

This project is a complete demonstration of a 'Business Analytics' cycle, implemented entirely in Excel without the use of Python or R.

The goal is not simply to produce graphs, but to simulate real-world business decision-making: from cleaning raw data to prescriptive analysis, including advanced statistical modeling.

This repository proves that Excel, pushed to its limits (Solver, Monte Carlo, Manual Regressions), remains a powerful Data Science tool.

-----

## Business Context (Scenario)

As a Business Analyst for an international retail chain (*Global Superstore*), management entrusted me with a 4-year sales history analysis with three strategic questions:

1. Segmentation : Who are our customers and how can we adapt our offering?

2. Profitability : What are the exact drivers of our margins and the causes of our returns?

3. Future : What are the forecasts for year N+1 and what are the financial risks?

-----

## 🛠️ Technical Stack & Skills

This project utilizes the full Excel analytical suite:
  **Data Preparation (ETL):** `XLOOKUP`, `VLOOKUP`, `Power Query`, `Text-to-Columns`.
  **Conditional Logic :** `Nested IFS`, `AND/OR` for order quality scoring.
  **Machine Learning (Excel) :** K-Means Clustering algorithm built manually (via Solver) on an RFM (Recency, Frequency, Monetary) analysis.
  **Statistical Modeling:**
  - Linear & Multiple Regression (Profit Factors).
  - Logistic Regression (Return Probability) calculated via Maximum Likelihood Estimation.
  **Risk Management:** Monte Carlo simulation (1000 iterations).
  **Optimization:** Linear Programming (Simplex LP) for budget allocation.

-----

## Analysis Structure

| Stage | Analysis | Key Tools | Business Objective |

| :--- | :--- | :--- | :--- |

| **01** | **Data Cleaning** | `XLOOKUP`, `IF`, `DATE` |Consolidate returns and clean heterogeneous date formats.|
| **02** | **Segmentation** | `Solver`, `Euclidean Distance` |Identify 3 customer clusters (VIP, Standard, At Risk) using K-Means.|
| **03** | **Keys Factors** | `Analysis Toolpak` | Measure the real impact of the discount on profit ($R^2$ et P-values). |
| **04** | **Predictions** | `Solver (GRG Nonlinear)` | Model the probability of return: $$P(Y=1) = \frac{1}{1 + e^{-(\beta_0 + \beta_1X)}}$$ |
| **05** | **Forecasting** | `Forecast Sheet` | Forecast sales for year N+1 while accounting for seasonality. |
| **06** | **Risks** | `NORM.INV`, `RAND` |  Monte Carlo simulation: Probability of financial loss. |
| **07** | **Optimization** | `Solver (Simplex)` | Maximize marketing ROI under budget constraints. |

-----

## Repository Architecture

```text
📂 Global-Superstore-Analytics
│
├── 📂 data
│   ├── raw_superstore.csv          # Raw Data (Source : Kaggle)
│   └── processed_superstore.xlsx   # Cleaned and Enriched Data (Master Table)
│
├── 📂 analysis
│   ├── 01_Data_Prep_Lookups.xlsx   # Data Cleaning and Joins
│   ├── 02_Clustering_RFM.xlsx      # Manual K-Mean Segmentation
│   ├── 03_Regressions.xlsx         # Linear Models and Logistics
│   ├── 04_Monte_Carlo_Sim.xlsx     # Risks Simulation
│   └── 05_Optimization.xlsx        # Budget Allocation (Solver)
│
├── 📂 reports
│   └── Executive_Summary.pdf       # Final Dashboard and Conclusions
│
└── README.md
```

-----

## Key Findings (Examples)

* **Insight 1:**

* **Insight 2:**

* **Insight 3:** 

-----
## Author

**Elise Berger**
*Étudiante passionnée par la Data Analysis et la résolution de problèmes business.*
[Lien vers votre LinkedIn]
