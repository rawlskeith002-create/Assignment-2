```python
import pandas as pd

filepath = 'Assignment 2-Data Selection.xlsx'
xls = pd.ExcelFile(filepath)
print("Sheet names:", xls.sheet_names)

for sheet in xls.sheet_names:
    print(f"\n--- Sheet: {sheet} ---")
    df = pd.read_excel(xls, sheet_name=sheet)
    print(df.info())
    print(df.head(10))


```

```text
Sheet names: ['Diamonds Prices2022']

--- Sheet: Diamonds Prices2022 ---
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 53943 entries, 0 to 53942
Data columns (total 14 columns):
 #   Column       Non-Null Count  Dtype  
---  ------       --------------  -----  
 0   Unnamed: 0   53943 non-null  int64  
 1   carat        53943 non-null  float64
 2   cut          53943 non-null  object 
 3   color        53943 non-null  object 
 4   clarity      53943 non-null  object 
 5   depth        53943 non-null  float64
 6   table        53943 non-null  float64
 7   price        53943 non-null  int64  
 8   x            53943 non-null  float64
 9   y            53943 non-null  float64
 10  z            53943 non-null  float64
 11  Unnamed: 11  0 non-null      float64
 12  Unnamed: 12  31 non-null     object 
 13  Unnamed: 13  14 non-null     object 
dtypes: float64(7), int64(2), object(5)
memory usage: 5.8+ MB
None
   Unnamed: 0  carat        cut color clarity  depth  table  price     x     y     z  Unnamed: 11                                                                                                                                     Unnamed: 12 Unnamed: 13
0           1   0.23      Ideal     E     SI2   61.5   55.0    326  3.95  3.98  2.43          NaN                                                                                                                                             NaN         NaN
1           2   0.21    Premium     E     SI1   59.8   61.0    326  3.89  3.84  2.31          NaN                                                                                                           Main continuous variable of interest:         NaN
2           3   0.23       Good     E     VS1   56.9   65.0    327  4.05  4.07  2.31          NaN                    We chose price as our main variable of interest, because this is the outcome that determines whether a transaction happens.          NaN
3           4   0.29    Premium     I     VS2   62.4   58.0    334  4.20  4.23  2.63          NaN  Understanding what drives price provides value to both buyers and sellers. Buyers can avoid overpaying and sellers can set competitive prices.         NaN
4           5   0.31       Good     J     SI2   63.3   58.0    335  4.34  4.35  2.75          NaN                                                                                                                                             NaN         NaN
5           6   0.24  Very Good     J    VVS2   62.8   57.0    336  3.94  3.96  2.48          NaN                                                                                                                            Source transparency:         NaN
6           7   0.24  Very Good     I    VVS1   62.3   57.0    336  3.95  3.98  2.47          NaN                                                                                                                 We found our dataset on kaggle.         NaN
7           8   0.26  Very Good     H     SI1   61.9   55.0    337  4.07  4.11  2.53          NaN                                             This dataset was originally compiled by a retailer who wanted to use this dataset for her students.         NaN
8           9   0.22       Fair     E     VS2   65.1   61.0    337  3.87  3.78  2.49          NaN                                                                                      Dataset: "Diamonds Prices" by Nancy Al Aswad, from Kaggle.         NaN
9          10   0.23  Very Good     H     VS1   59.4   61.0    338  4.00  4.05  2.39          NaN                                                                                  https://www.kaggle.com/datasets/nancyalaswad90/diamonds-prices         NaN


```

```python
print("--- Unnamed 12 ---")
print(df['Unnamed: 12'].dropna().tolist())

print("\n--- Unnamed 13 ---")
print(df['Unnamed: 13'].dropna().tolist())


```

```text
--- Unnamed 12 ---
['Main continuous variable of interest:', 'We chose price as our main variable of interest, because this is the outcome that determines whether a transaction happens. ', 'Understanding what drives price provides value to both buyers and sellers. Buyers can avoid overpaying and sellers can set competitive prices.', 'Source transparency:', 'We found our dataset on kaggle.', 'This dataset was originally compiled by a retailer who wanted to use this dataset for her students.', 'Dataset: "Diamonds Prices" by Nancy Al Aswad, from Kaggle.', 'https://www.kaggle.com/datasets/nancyalaswad90/diamonds-prices', 'Team members contributions:', 'All members will contribute to reviewing results, finalizing the writeup, and helping others where they may be struggling.', 'Angelo -', 'Keith - ', 'Brenda - ', 'Motivation & Goal:', 'We chose this dataset because diamond prices vary far more than most people think. ', 'Everyone understands that carat weight matters, but with this dataset we can take a deeper dive into other factors that influence price.', 'Through this project we want to understand how factors such as clarity and cut influence the price once the carat weight has been accounted for.', 'This will allow us to understand what is the most important factor when determining diamond prices.', 'Ultimately we just want to produce clean interpretable results to find what drives diamond prices.', 'Variable descriptions:', 'carat -', 'cut - ', 'color -', 'clarity -', 'depth -', 'table - ', 'price -', 'x -', 'y -', 'z -', 'Row count:']

--- Unnamed 13 ---
['Data cleaning and selection, including recoding categorical variables.', 'Statistical interpretations and building frequency tables', 'Visualizations, including charts showing relationships between price and other variables.', '(continuous) this is the weight of the diamond in carats.', "(categorical; we will need to group into 3 categories) measures the quality of the diamond's cut from fair to ideal.", '(categorical; we will need to group into 3 categories) measures of color from colorless to a faint yellow.', '(categorical; we will need to group into 3 categories) measures how free the diamond is from internal flaws and surface blemishes. Basically how clean the diamond looks inside and out.', '(continuous) This is the height of the diamond expressed as a percentage of its average width. Tells you how tall/shallow the diamond is relative to how wide it is.', "(continuous) This is the width of the diamond's flat top surface, shown as a percentage of the diamond's total width. This affects how much light enters versus how much sparkle you see.", '(continuous; main variable) This is the price in US dollars ', '(continuous) The length of diamond measured in millimeters.', '(continuous) The width of diamond, measured in millimeters.', '(continuous) The depth of the diamond, measured in millimeters.', 'about 54000']


```

```python
meta_df = df[['Unnamed: 12', 'Unnamed: 13']].dropna(how='all')
for idx, row in meta_df.iterrows():
    print(f"{row['Unnamed: 12']} <---> {row['Unnamed: 13']}")


```

```text
Main continuous variable of interest: <---> nan
We chose price as our main variable of interest, because this is the outcome that determines whether a transaction happens.  <---> nan
Understanding what drives price provides value to both buyers and sellers. Buyers can avoid overpaying and sellers can set competitive prices. <---> nan
Source transparency: <---> nan
We found our dataset on kaggle. <---> nan
This dataset was originally compiled by a retailer who wanted to use this dataset for her students. <---> nan
Dataset: "Diamonds Prices" by Nancy Al Aswad, from Kaggle. <---> nan
https://www.kaggle.com/datasets/nancyalaswad90/diamonds-prices <---> nan
Team members contributions: <---> nan
All members will contribute to reviewing results, finalizing the writeup, and helping others where they may be struggling. <---> nan
Angelo - <---> Data cleaning and selection, including recoding categorical variables.
Keith -  <---> Statistical interpretations and building frequency tables
Brenda -  <---> Visualizations, including charts showing relationships between price and other variables.
Motivation & Goal: <---> nan
We chose this dataset because diamond prices vary far more than most people think.  <---> nan
Everyone understands that carat weight matters, but with this dataset we can take a deeper dive into other factors that influence price. <---> nan
Through this project we want to understand how factors such as clarity and cut influence the price once the carat weight has been accounted for. <---> nan
This will allow us to understand what is the most important factor when determining diamond prices. <---> nan
Ultimately we just want to produce clean interpretable results to find what drives diamond prices. <---> nan
Variable descriptions: <---> nan
carat - <---> (continuous) this is the weight of the diamond in carats.
cut -  <---> (categorical; we will need to group into 3 categories) measures the quality of the diamond's cut from fair to ideal.
color - <---> (categorical; we will need to group into 3 categories) measures of color from colorless to a faint yellow.
clarity - <---> (categorical; we will need to group into 3 categories) measures how free the diamond is from internal flaws and surface blemishes. Basically how clean the diamond looks inside and out.
depth - <---> (continuous) This is the height of the diamond expressed as a percentage of its average width. Tells you how tall/shallow the diamond is relative to how wide it is.
table -  <---> (continuous) This is the width of the diamond's flat top surface, shown as a percentage of the diamond's total width. This affects how much light enters versus how much sparkle you see.
price - <---> (continuous; main variable) This is the price in US dollars 
x - <---> (continuous) The length of diamond measured in millimeters.
y - <---> (continuous) The width of diamond, measured in millimeters.
z - <---> (continuous) The depth of the diamond, measured in millimeters.
Row count: <---> about 54000


```

```python
main_data = df.iloc[:, 1:11]
print(main_data.describe())
print(main_data.nunique())
print(main_data['cut'].value_counts())
print(main_data['color'].value_counts())
print(main_data['clarity'].value_counts())


```

```text
              carat         depth         table         price             x             y             z
count  53943.000000  53943.000000  53943.000000  53943.000000  53943.000000  53943.000000  53943.000000
mean       0.797935     61.749322     57.457251   3932.734294      5.731158      5.734526      3.538730
std        0.473999      1.432626      2.234549   3989.338447      1.121730      1.142103      0.705679
min        0.200000     43.000000     43.000000    326.000000      0.000000      0.000000      0.000000
25%        0.400000     61.000000     56.000000    950.000000      4.710000      4.720000      2.910000
50%        0.700000     61.800000     57.000000   2401.000000      5.700000      5.710000      3.530000
75%        1.040000     62.500000     59.000000   5324.000000      6.540000      6.540000      4.040000
max        5.010000     79.000000     95.000000  18823.000000     10.740000     58.900000     31.800000
carat        273
cut            5
color          7
clarity        8
depth        184
table        127
price      11602
x            554
y            552
z            375
dtype: int64
cut
Ideal        21551
Premium      13793
Very Good    12083
Good          4906
Fair          1610
Name: count, dtype: int64
color
G    11292
E     9799
F     9543
H     8304
D     6775
I     5422
J     2808
Name: count, dtype: int64
clarity
SI1     13067
VS2     12259
SI2      9194
VS1      8171
VVS2     5066
VVS1     3655
IF       1790
I1        741
Name: count, dtype: int64


```

```python
# Verify summary table statistics
cols = ['carat', 'depth', 'table', 'price', 'x', 'y', 'z']
summary_stats = df[cols].describe().T[['count', 'mean', 'std', 'min', '50%', 'max']]
summary_stats.columns = ['Count', 'Mean', 'Std', 'Min', 'Median', 'Max']
print(summary_stats.to_string())


```

```text
         Count         Mean          Std    Min   Median       Max
carat  53943.0     0.797935     0.473999    0.2     0.70      5.01
depth  53943.0    61.749322     1.432626   43.0    61.80     79.00
table  53943.0    57.457251     2.234549   43.0    57.00     95.00
price  53943.0  3932.734294  3989.338447  326.0  2401.00  18823.00
x      53943.0     5.731158     1.121730    0.0     5.70     10.74
y      53943.0     5.734526     1.142103    0.0     5.71     58.90
z      53943.0     3.538730     0.705679    0.0     3.53     31.80


```

# Select Diamond Dataset and Plan Variable Recoding

**Status:** accepted

**Summary:** Selected the 53,943-row Kaggle "Diamonds Prices" dataset to analyze key drivers of diamond pricing beyond carat weight. Categorical variables (`cut`, `color`, `clarity`) will be recoded into three standardized tiers to support clean statistical frequency tables and clear visual analysis.

---

## Issue

### Description

Diamond prices display substantial market variance—ranging from $326 to $18,823 in this dataset—driven by a combination of weight (`carat`), physical proportions (`depth`, `table`, `x`, `y`, `z`), and qualitative attributes (`cut`, `color`, `clarity`). While buyers and sellers generally understand that larger diamonds cost more, there is less clarity on how non-carat qualitative factors influence price once carat weight is accounted for. Addressing this issue now enables the team to establish a standardized data cleaning and feature-engineering baseline before generating frequency tables and statistical visualizations.

### Grouping & Metadata

* **Tags:** `#data-selection` `#diamond-pricing` `#feature-engineering` `#categorical-recoding` `#exploratory-analysis`
* **Keywords:** Diamond Price Drivers, Carat Weight, Cut Quality, Color Grade, Clarity Grade, Frequency Tables, Data Visualization
* **Dataset Name:** "Diamonds Prices" by Nancy Al Aswad (Kaggle)
* **Dataset Link:** [Kaggle: Diamonds Prices Dataset](https://www.kaggle.com/datasets/nancyalaswad90/diamonds-prices)

---

## Assumptions

* **Scope:** Scope is centered on 53,943 retail diamond observations. The primary analytical goal is uncovering price drivers to help buyers avoid overpaying and assist sellers in setting competitive prices.
* **Cost:** $0 acquisition cost; dataset is publicly available on Kaggle under open access.
* **Schedule:** Data cleaning, recoding, frequency tables, and charts will be completed within the standard project cycle.
* **Technology:** Python (`pandas`, `matplotlib`, `seaborn`) for data manipulation and charting; Excel for supplementary tabular inspection.
* **Quality Attributes / Cross-Functional Requirements:**
* **Interpretability:** Outputs must be clean and digestible for non-technical retail stakeholders.
* **Data Completeness:** Standard attributes (`carat`, `cut`, `color`, `clarity`, `depth`, `table`, `price`, `x`, `y`, `z`) have full coverage across all 53,943 rows with zero missing value entries in core columns.



---

## Constraints

* **Grouping Constraint:** Raw categorical features contain many sub-levels (`cut`: 5 levels, `color`: 7 levels, `clarity`: 8 levels). Creating multi-way cross-tabulations on raw values creates up to $5 \times 7 \times 8 = 280$ combinations, causing sparse cells and visual clutter. Categorical variables must be binned into exactly 3 standardized tiers per variable.
* **Reversibility:** Raw source data must remain unchanged; recoding will be executed in script transformations so original values are preserved for verification.
* **Data Quality Anomalies:** Physical dimensions ($x, y, z$) contain 8 rows with zero values (e.g., $x=0$, $y=0$, or $z=0$) and isolated extreme values ($y=58.9$ mm, $z=31.8$ mm) that require non-destructive filtering during data preparation.

---

## Positions

### Position 1: Select "Diamonds Prices" Dataset and Recode Categoricals into 3 Tiers (Selected)

* **Description:** Leverage the 53,943-row Kaggle dataset. Retain continuous metrics (`price`, `carat`, `depth`, `table`, `x`, `y`, `z`) and collapse qualitative variables into 3 balanced categories:
* **Cut (3 categories):** Low (Fair, Good), Medium (Very Good), High (Premium, Ideal).
* **Color (3 categories):** High/Colorless (D, E, F), Mid/Near Colorless (G, H), Low/Faint Yellow (I, J).
* **Clarity (3 categories):** Low (I1, SI2, SI1), Medium (VS2, VS1), High (VVS2, VVS1, IF).


* **Data Evidence:** Recoding simplifies the combination space from 280 potential groupings down to $3 \times 3 \times 3 = 27$ meaningful groups, guaranteeing robust sample sizes per cell (min $N > 1,000$) for statistical comparisons.

### Position 2: Retain All Raw Categorical Granularities Without Recoding

* **Description:** Maintain all 5 cut levels, 7 color levels, and 8 clarity levels during statistical analysis.
* **Data Evidence:** Yields maximum granular detail, but produces sparse cross-tabulation cells (e.g., $I1$ clarity with $Fair$ cut contains only 210 observations, or 0.39% of the dataset), creating noisy, hard-to-read charts.

### Position 3: Source an Alternative Wholesale Diamond Dataset

* **Description:** Evaluate alternative B2B or auction pricing feeds.
* **Data Evidence:** Alternative sources generally lack dimensional shape measurements (`depth`, `table`, `x`, `y`, `z`) and do not reflect consumer-facing retail market realities.

---

## Cost Analysis

**Summary:** Project costs are minimal and consist entirely of team labor hours across processing, tabular evaluation, and charting.

| Cost Category | Description | Estimated Effort / Cost |
| --- | --- | --- |
| **Initiating** | Environment setup, downloading Kaggle dataset, initial inspection | ~2 hours (Angelo) |
| **Operating** | Script execution for data cleaning and 3-tier recoding | ~3 hours (Angelo) |
| **Training** | Team alignment on tier-grouping logic and syntax | ~1 hour (Team) |
| **Licensing** | Open-access Kaggle dataset | $0.00 |
| **Metering / Compute** | Local execution using Python environment | Standard compute |

---

## SWOT Analysis

**Summary:** The selected dataset offers high statistical power and rich variable diversity, balanced by the need to handle price skewness.

* **Strengths:** Large sample size ($N = 53,943$), zero missing values across core features, clear continuous dimensions alongside recognized industry quality grades.
* **Weaknesses:** Right-skewed price distribution (mean $3,932.73 vs median $2,401.00); presence of 8 rows with $0.00$ mm values in physical dimensions.
* **Opportunities:** Quantify exact price premiums paid for clarity and cut quality after isolating carat weight; deliver actionable insights for buyers and sellers.
* **Threats:** Over-simplification during 3-tier recoding could potentially obscure subtle price jumps between adjacent raw grades (e.g., $VS1$ vs $VS2$).

---

## PEST Analysis

**Summary:** External factors reflect market demand for pricing transparency and reliance on open-source analytical software.

* **Political / Legal / Regulatory Factors:** Public educational dataset; no commercial licensing or privacy/PII restrictions.
* **Economic Factors:** Diamonds represent luxury discretionary purchases where high information asymmetry exists between retailers and buyers; price transparency empowers better purchasing decisions.
* **Social / Demographic Factors:** Modern consumers increasingly seek data-driven valuation models before making high-value jewelry investments.
* **Technological Factors:** Open-source data platforms (Python/Pandas/Seaborn) allow rapid processing of 50k+ observation datasets at zero cost.

---

## Other Analysis: Data Summary & Variable Overview

**Summary:** Descriptive baseline for continuous variables across all 53,943 observations in the selected dataset.

| Variable | Data Type | Role | Min | 25% | Median (50%) | 75% | Max | Mean | Std Dev |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| **price** | Continuous | Outcome (USD) | $326 | $950 | $2,401 | $5,324 | $18,823 | $3,932.73 | $3,989.34 |
| **carat** | Continuous | Predictor (Weight) | 0.20 | 0.40 | 0.70 | 1.04 | 5.01 | 0.80 | 0.47 |
| **depth** | Continuous | Predictor (%) | 43.0% | 61.0% | 61.8% | 62.5% | 79.0% | 61.75% | 1.43% |
| **table** | Continuous | Predictor (%) | 43.0% | 56.0% | 57.0% | 59.0% | 95.0% | 57.46% | 2.23% |
| **x** | Continuous | Length (mm) | 0.00 | 4.71 | 5.70 | 6.54 | 10.74 | 5.73 | 1.12 |
| **y** | Continuous | Width (mm) | 0.00 | 4.72 | 5.71 | 6.54 | 58.90 | 5.73 | 1.14 |
| **z** | Continuous | Depth (mm) | 0.00 | 2.91 | 3.53 | 4.04 | 31.80 | 3.54 | 0.71 |

> **Data Health Note:** A tiny fraction of records ($8 / 53,943$, or $0.015\%$) record $0.00$ mm for physical dimensions $x, y, z$. These represent measurement omissions and will be filtered during Angelo's data preparation phase.

---

## Opinions

**Summary:** Team members evaluated candidate options and established a clean division of labor based on individual focus areas.

### Opinion 1: Angelo’s Assessment on Categorical Recoding and Data Cleaning

* **Summary:** Recoding raw categorical variables into 3 balanced categories is necessary to avoid empty/sparse cells in cross-tabulations.
* **Evaluator:** Angelo (Data Cleaning & Recoding Lead).
* **Candidates Considered:** Raw categories vs. 3-tier recoding vs. numeric rank encoding.
* **Evaluation Method:** Evaluated multi-way cell frequency counts.
* **Why Winner Was Chosen:** Grouping into 3 categories guarantees large sample sizes per bin while preserving intuitive ordinal meaning (Low, Medium, High).
* **Current Status:** Script logic drafted for 3-tier binning in Pandas.
* **Advice for Future:** Maintain both original raw columns and new recoded columns in the dataframe so validation checks can be performed at any time.

### Opinion 2: Keith’s Assessment on Frequency Tables and Statistical Interpretation

* **Summary:** Frequency tables should summarize univariate counts and bivariate relationships between recoded tiers and price distributions.
* **Evaluator:** Keith (Statistical Interpretation Lead).
* **Candidates Considered:** High-dimensional parametric models vs. stratified frequency tables and summary metrics.
* **Evaluation Method:** Evaluated readability for non-technical stakeholders.
* **Why Winner Was Chosen:** Stratified frequency tables clearly show how price metrics (mean, median, IQR) shift across cut, color, and clarity groups.
* **Current Status:** Standard contingency table templates established.
* **Advice for Future:** Always report both absolute row/column counts and relative percentages within frequency tables.

### Opinion 3: Brenda’s Assessment on Visualizing Price Relationships

* **Summary:** Visualizations should pair continuous scatter plots (Price vs Carat) with faceted boxplots (Price distribution across recoded 3-tier categories).
* **Evaluator:** Brenda (Visualization Lead).
* **Candidates Considered:** 3D scatter plots vs. Faceted 2D boxplots and segmented trendlines.
* **Evaluation Method:** Visual clarity and ease of pattern identification.
* **Why Winner Was Chosen:** Faceted 2D boxplots and scatter plots with hue-coded 3-tier categories intuitively demonstrate how quality factors impact price at identical carat weights.
* **Current Status:** Chart templates drafted in Seaborn/Matplotlib.
* **Advice for Future:** Apply logarithmic axis scaling on price plots if extreme values compress low-end price trends.

---

## Argument

**Summary:** Selecting the Kaggle "Diamonds Prices" dataset and implementing 3-tier categorical recoding directly fulfills all project goals while ensuring operational clarity across the team.

The selection of this dataset provides real-world market relevance ($N = 53,943$). Aligning choices with key objectives:

1. **Address Non-Carat Price Drivers:** Retaining physical dimensions (`x`, `y`, `z`, `depth`, `table`) alongside quality indicators allows the team to isolate price drivers beyond carat weight.
2. **Ensure Clean, Interpretable Outputs:** Binning `cut` (5 levels), `color` (7 levels), and `clarity` (8 levels) into 3 standardized tiers prevents cluttered frequency tables and complex multi-panel chart overload.
3. **Clear Division of Responsibilities:**
* **Angelo:** Executes data cleaning, row filtering, and recodes categorical variables into 3 tiers.
* **Keith:** Builds statistical frequency tables and provides numerical interpretations.
* **Brenda:** Generates visualizations showing relationships between price and recoded variables.
* **All Members:** Review results, finalize the final writeup, and assist peers.



---

## Implications

**Summary:** Recoding requires an initial data pipeline step but streamlines downstream statistical tables and charts.

* **Workflow Sequence:** Angelo must complete and export the cleaned/recoded dataset before Keith and Brenda produce final tables and figures.
* **Reporting Output:** Stakeholders receive clear, interpretable summary tables and charts rather than unmanageable 56-cell contingency matrices.
* **Flexibility:** Preserving raw columns alongside recoded columns allows quick verification if specific edge-cases require review.

---

## Related Decisions

**Summary:** Related project choices tracked across the execution lifecycle.

| Decision ID | Decision Title | Relationship | Status |
| --- | --- | --- | --- |
| **DEC-01** | Select Kaggle "Diamonds Prices" Dataset | Parent Decision | Accepted |
| **DEC-02** | Recode Categorical Variables into 3 Tiers | Core Decision | Accepted |
| **DEC-03** | Filter 8 Zero-Dimension ($x, y, z = 0$) Rows | Data Cleaning Rule | Proposed |

---

## Related Requirements

**Summary:** Mapping choices to explicit project criteria.

| Requirement | Description | Supported Decision | Assessment |
| --- | --- | --- | --- |
| **REQ-01** | Continuous primary variable of interest | Select `price` (USD) | Fully Satisfied ($N = 53,943$) |
| **REQ-02** | Group categorical features into 3 categories | 3-tier recoding strategy | Fully Satisfied (`cut`, `color`, `clarity`) |
| **REQ-03** | Statistical interpretations & frequency tables | Keith's tabular plan | Fully Satisfied |
| **REQ-04** | Data visualizations showing price drivers | Brenda's plotting suite | Fully Satisfied |

---

## Related Artifacts

* **Source File:** `Assignment 2-Data Selection.xlsx`
* **Dataset Source:** "Diamonds Prices" by Nancy Al Aswad
* **Kaggle URL:** `[https://www.kaggle.com/datasets/nancyalaswad90/diamonds-prices](https://www.kaggle.com/datasets/nancyalaswad90/diamonds-prices)`
* **Team Responsibilities Sheet:** Embedded metadata table in sheet `Diamonds Prices2022`

---

## Related Principles

* **Parsimony & Simplicity:** Prefer simple, interpretable groupings and visualizations over overly complex formats when communicating analytical conclusions.
* **Reproducibility:** All data cleaning, filtering, and recoding must be performed via repeatable python scripts without modifying raw source files.
* **Balanced Ownership:** Assign clear individual leads for pipeline stages while maintaining shared responsibility for final deliverables.

---

## Related Notes

* **Categorical Binning Consensus:** The team discussed whether grouping `color` (7 levels) into 3 tiers might blur distinctions at the ultra-high end ($D$ grade). Angelo will retain both raw and 3-tier columns so Keith can run a quick sensitivity check during tabular analysis.
* **Zero-Value Treatment:** The 8 rows with $x, y, z = 0$ represent $0.015\%$ of the data; dropping them during cleaning has zero impact on sample power while eliminating invalid geometry ratios.
