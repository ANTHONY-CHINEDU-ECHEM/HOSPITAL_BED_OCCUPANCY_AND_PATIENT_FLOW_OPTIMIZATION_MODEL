# Hospital Bed Occupancy and Patient Flow Optimization Model

An Excel decision model that forecasts inpatient bed occupancy 72 hours ahead by unit, flags units at breach risk, and quantifies the value of discharge and staffing policy changes for a 400 bed regional hospital. Built from a supplied 15,394 row ADT extract through a fully documented, eleven step cleaning procedure.

Excel Project 1 of 10 in a data analyst portfolio. Prepared by Anthony Chinedu Echem.

## Business problem

A 400 bed hospital loses an estimated 2.3 million dollars a year to bed block delays, ad hoc discharge planning and ED boarding. Leadership needs a living model that predicts occupancy 72 hours out and flags units at risk before they breach.

## Results at a glance

<table>
  <tr><th>Measure</th><th>Result</th></tr>
  <tr><td>Hospital occupancy at forecast start</td><td>71.8 percent</td></tr>
  <tr><td>Expected occupancy at 72 hours</td><td>69.3 percent</td></tr>
  <tr><td>Mean discharge lag (order to bed release)</td><td>5.6 hours</td></tr>
  <tr><td>Annual cost of bed unavailable time</td><td>2.3 million dollars (calibrated to the brief)</td></tr>
  <tr><td>Value of default policy levers</td><td>about 554 thousand dollars a year, equivalent to 5.1 staffed beds</td></tr>
  <tr><td>Forecast backtest, hospital weeks within 4 points</td><td>88.2 percent (mean absolute error 2.05 points)</td></tr>
</table>

## Success metrics

<table>
  <tr><th>Catalogue metric</th><th>Result</th><th>Status</th></tr>
  <tr><td>Forecast accuracy within 4 points of actual</td><td>88.2 percent of hospital weeks; 36.1 percent at unit week level, limited by about one occupancy reading per unit per day</td><td>Met at hospital level</td></tr>
  <tr><td>Full simulation runtime under 30 seconds</td><td>8 seconds, measured full recalculation of 333,179 formulas</td><td>Met</td></tr>
  <tr><td>Three validated historical surge scenarios</td><td>3 of 3 highest admission weeks replayed stably</td><td>Met</td></tr>
  <tr><td>Bed meeting preparation time cut by about 70 percent</td><td>Not measurable from data; heat map replaces manual census compilation</td><td>Estimate</td></tr>
</table>

## Workbook structure

<table>
  <tr><th>Sheet</th><th>Purpose</th></tr>
  <tr><td>01 Executive KPI Summary</td><td>One page KPIs, scorecard, live findings and recommendations</td></tr>
  <tr><td>02 Breach Risk Dashboard</td><td>Unit heat map, Monte Carlo fan chart, admissions trend and lag histogram, driven by unit, forecast date and policy selectors</td></tr>
  <tr><td>03 72h Occupancy Forecast</td><td>Rolling 72 hour expected census, variance and breach probability by unit</td></tr>
  <tr><td>04 Discharge Lag Fitting</td><td>Poisson, censored Normal, Gamma and hurdle Gamma fits with chi square tests by specialty</td></tr>
  <tr><td>05 Scenario Simulator</td><td>Seven day discharge, lag reduction, turnover, before noon, staffing and surge levers with financial impact</td></tr>
  <tr><td>06 Assumptions and Validation</td><td>Assumption register, validation log and limitations</td></tr>
  <tr><td>M1 to M4</td><td>Little's Law calibration, 1,000 trial Monte Carlo engine, weekly backtest, surge stress test</td></tr>
  <tr><td>C1 to C11</td><td>Data cleaning procedure, one step per sheet</td></tr>
  <tr><td>Clean Data, Raw Data, Data Dictionary</td><td>Analysis table with named ranges, and the supplied data unchanged</td></tr>
</table>

## Method

1. Clean the extract (see docs/data_cleaning_procedure.md). The recorded discharge date contradicts length of stay in 94 percent of records, so the derived discharge date (admission plus length of stay) is used throughout.
2. Allocate the 400 beds to units by share of patient days and calibrate daily admissions with Little's Law (census equals arrival rate times effective length of stay, where effective stay includes discharge lag and bed turnover).
3. Propagate expected census and its variance exactly for three days, with weekend discharge factors, and convert to breach probability.
4. Cross check with 1,000 Monte Carlo trials built from CRITBINOM draws; the simulated mean agrees with the exact mean within 2 percent.
5. Value bed unavailable hours at the cost per hour implied by the brief, and simulate policy levers.

## Key findings

* Beds stay blocked for about 6.3 hours after each discharge decision (lag plus turnover), which the brief values at 2.3 million dollars a year.
* Discharge lag is best described by a Normal distribution censored at zero; its dispersion index of about 2.7 means variation, not only the average, drives delay.
* Recorded occupancy does not move with admission volume (weekly correlation close to zero), so it is not reconciled to the ADT feed. Occupancy should come from a reconciled midnight census before unit level forecasting goes live.

## How to use

Open the workbook in Microsoft Excel 2010 or later; calculation is automatic. Change the teal bordered selectors on the dashboard (unit, forecast date, policy mode) and the yellow input cells on 05 and M1. Blue text on yellow marks inputs, black text marks formulas, green text marks links. Press F9 to redraw the Monte Carlo trials.

The selectors are in cell drop down lists (data validation), which behave like combo boxes and work in every Excel version. A Form Control combo box can be linked to the same cells if preferred.

## Reproducing the build

The workbook is generated by Python so every sheet can be rebuilt from the raw data.

```
pip install pandas numpy openpyxl scipy
python3 src/p1_build.py
```

Run from the repository root. The script reads data/raw, writes workbook/, and the workbook recalculates when opened.

## Repository structure

```
README.md
workbook/      finished Excel model
data/raw/      supplied dataset, unchanged
src/           cleaning engine, sheet writers and build script
docs/          data cleaning procedure and methodology
```

## Limitations

The extract is a sample of about 3,800 encounters a year, so rates are expressed as shares and calibrated to the 400 bed base. Clock times are absent, so discharge lag is taken from the numeric field after validation. Ambiguous slash dates (15.6 percent) were resolved as month first and flagged; excluding them moves the backtest from 88.2 to 85.8 percent.

## Author

Anthony Chinedu Echem
