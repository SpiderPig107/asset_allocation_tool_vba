# VBA Investment Portfolio Performance Analyzer
Automated VBA portfolio analyzer with dynamic risk modeling and performance visualization.

To experience the full automation and dynamic charting, I recommend downloading the original macro-enabled workbook. The file name is ```Project_v.1.14.xlsm```.

## Project Overview
This project provides a robust Excel-based solution for quantitative portfolio analysis. It enables users to select specific assets from a pre-defined list and evaluate their combined performance using both standard and downside risk metrics. The tool automates the entire process from data processing to report generation via custom VBA macros.

### Core Features

**1. Dynamic Asset Selection:** The interface allows users to choose two companies from a set of 12 for comparative analysis.

**2. Custom Allocation Logic:** Supports user-defined weighting factors for individual assets. It must be positive to prevent short selling.

**3. Automated Charting:** Automatically generate dedicated chart sheets for the selected companies and the overall portfolio.

**4. Comprehensive Risk Metrics:** Calculates and displays key statistical indicators directly on the graphs, including:

      - Mean: To assess average expected return.
      
      - Standard Deviation: To measure total volatility.
      
      - Semi-Standard Deviation: To evaluate downside risk and potential losses.

**5. Portfolio Value Tracking:** Visualizes the change in portfolio value over time, assuming a baseline initial value of 100.

## Technical Highlights
- **Excel Object Model:** The macro manages the creation and formatting of separate sheets for charts and data summaries.
  
- **Statistical Computation:** Implements financial formulas in VBA to calculate downside volatility (semi-deviation), providing more nuanced risk assessment than standard variance alone.

- **User Experience (UX) Focused:** Designed with a clear workflow where initial data entry triggers automated visualization and reporting.

## How to Use
1. **Enter Data:** Ensure initial company data is populated in the spreadsheet.
2. **Select Assets:** Use the interface to choose the two companies you wish to analyze.
3. **Assign Weights:** Enter the asset allocations (weighting factors) for each company.
4. **Run Macro:** Execute the macro to generate the analysis charts and statistical reports on separate sheets.
