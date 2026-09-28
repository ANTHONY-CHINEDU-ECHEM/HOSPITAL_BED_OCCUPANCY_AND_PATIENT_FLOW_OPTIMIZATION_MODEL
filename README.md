# Hospital Bed Occupancy and Patient Flow Optimization Model

An Excel decision model that forecasts inpatient bed occupancy 72 hours ahead by unit, flags units at risk of breaching safe capacity, and puts a financial value on discharge and staffing policy changes for a 400 bed regional hospital. The model is built from a supplied 15,394 row admission, discharge and transfer (ADT) extract through a fully documented, eleven step cleaning procedure, and it runs entirely in native Excel with no add ins, macros or external services.

## Project brief

Bed capacity is the constraint that quietly governs almost everything else in an acute hospital. When a ward cannot release a bed, the patient waiting for it stays on a trolley in the emergency department, elective admissions are cancelled or delayed, theatre lists slip, and staff spend their day negotiating space instead of delivering care. None of this appears on a single line of the budget, yet together it forms one of the largest avoidable costs a hospital carries. For the 400 bed regional hospital in this study, leadership estimates that bed block delays, ad hoc discharge planning and emergency department boarding cost around 2.3 million dollars every year, before counting the harder to price effects on patient experience, clinical risk and staff morale.

The root of the problem is that capacity is managed reactively. The daily bed meeting is typically built on a census compiled by hand from several systems, describes where the hospital stood that morning rather than where it is heading, and gives little warning before a unit tips into overflow. Discharge decisions are made during ward rounds, but the bed is not physically free until transport, medication, paperwork and cleaning are complete, and that gap between the decision and the empty bed is rarely measured. Weekend discharge rates fall away, so Monday mornings inherit a backlog that was entirely predictable on Friday. Leaders are left making escalation calls on instinct, with no shared view of which units are genuinely at risk and no evidence for which interventions would be worth their cost.

This project was commissioned to replace that reactive posture with a living, forward looking model. The brief asked for three things: a rolling 72 hour occupancy forecast by unit, automated escalation flags that identify breach risk before it happens, and a scenario simulator that shows leadership what policy changes such as seven day discharging, faster discharge processing or additional staffing would actually be worth in beds and in money. The model also had to be credible to a sceptical clinical and finance audience, which meant documenting every cleaning decision and assumption, validating the forecast against history, stress testing it against real surge weeks, and being honest about what the underlying data can and cannot support. Building it in native Excel was a deliberate choice, so that bed managers and analysts could open, inspect, adjust and trust it without specialist software.

## Objectives

* Deliver a rolling 72 hour expected census, variance and breach probability for every inpatient unit.
* Quantify the discharge lag (the time between a discharge order and the bed becoming available) and identify the statistical distribution that best describes it by specialty.
* Convert bed unavailable time into a defensible annual cost that reconciles to the brief.
* Let leadership test discharge, turnover, staffing and surge policies and see the impact in beds and dollars immediately.
* Validate the forecast against historical weeks and replay the three heaviest admission weeks as surge scenarios.
* Document every assumption, validation check and limitation in the workbook itself.

## What makes this model different

Most occupancy workbooks are static trackers that report yesterday's census. This model treats the hospital as a queue. It calibrates admissions, length of stay, discharge lag and bed turnover with Little's Law, propagates the expected census and its variance exactly across three days with weekend discharge effects, and then cross checks those results with a 1,000 trial Monte Carlo engine built from native worksheet functions. The outcome is not a single number but a probability that each unit breaches its capacity, which is the quantity a bed manager actually needs in order to decide whether to escalate.

## Data

The supplied ADT extract contains 15,394 rows describing inpatient encounters, with admission dates, recorded discharge dates, length of stay, specialty, unit, recorded occupancy and a numeric discharge lag field. Cleaning followed eleven documented steps, one per worksheet (C1 to C11), covering structure checks, date parsing, duplicate handling, field validation, logical consistency and flagging.

Two data issues shaped the entire analysis and are worth stating plainly:

* The recorded discharge date contradicts the recorded length of stay in 94 percent of records. The model therefore uses a derived discharge date (admission date plus length of stay) throughout, and treats the recorded field as unreliable.
* 15.6 percent of dates are written in an ambiguous slash format. These were resolved as month first and flagged, and their influence on the results was tested directly (see Assumptions and limitations).

## Results at a glance

<table>
  <tr><th>Measure</th><th>Result</th></tr>
  <tr><td>Hospital occupancy at forecast start</td><td>71.8 percent</td></tr>
  <tr><td>Expected occupancy at 72 hours</td><td>69.3 percent</td></tr>
  <tr><td>Mean discharge lag (order to bed release)</td><td>5.6 hours</td></tr>
  <tr><td>Total bed blocked time per discharge (lag plus turnover)</td><td>about 6.3 hours</td></tr>
  <tr><td>Annual cost of bed unavailable time</td><td>2.3 million dollars (calibrated to the brief)</td></tr>
  <tr><td>Value of default policy levers</td><td>about 554 thousand dollars a year, equivalent to 5.1 staffed beds</td></tr>
  <tr><td>Forecast backtest, hospital weeks within 4 points of actual</td><td>88.2 percent (mean absolute error 2.05 points)</td></tr>
  <tr><td>Monte Carlo agreement with exact expected census</td><td>within 2 percent</td></tr>
</table>

<img width="1066" height="360" alt="Screenshot 2026-09-28 at 22 56 09" src="https://github.com/user-attachments/assets/59b666a7-05b9-44c4-9476-e65a790d75de" />

## Success metrics

<table>
  <tr><th>Target</th><th>Result</th><th>Status</th></tr>
  <tr><td>Forecast accuracy within 4 points of actual occupancy</td><td>88.2 percent of hospital weeks; 36.1 percent at unit week level, limited by roughly one occupancy reading per unit per day</td><td>Met at hospital level</td></tr>
  <tr><td>Full simulation runtime under 30 seconds</td><td>8 seconds for a measured full recalculation of 333,179 formulas</td><td>Met</td></tr>
  <tr><td>Three validated historical surge scenarios</td><td>The 3 highest admission weeks replayed stably</td><td>Met</td></tr>
  <tr><td>Bed meeting preparation time cut by about 70 percent</td><td>Not measurable from the data; the heat map replaces manual census compilation</td><td>Estimate</td></tr>
</table>

<img width="1158" height="299" alt="Screenshot 2026-09-28 at 22 57 56" src="https://github.com/user-attachments/assets/6c00ac2f-e4eb-4795-a239-7cd26932eee3" />


## Findings in detail

### 1. The hospital's capacity problem is a flow problem, not a bed count problem

Occupancy at the start of the forecast window stands at 71.8 percent and is expected to ease to 69.3 percent over the following 72 hours. On headline figures alone, a 400 bed hospital running at around 70 percent does not look short of beds. Yet the brief describes 2.3 million dollars a year lost to bed block and boarding. The two facts are reconciled by flow: every discharge leaves a bed unusable for about 6.3 hours, made up of a mean discharge lag of 5.6 hours plus bed turnover. Across the volume of discharges the hospital handles, those hours add up to a substantial block of capacity that exists on paper but cannot be used when patients are waiting. The insight for leadership is that buying or opening beds would treat the symptom, while shortening the time between the discharge decision and the clean, available bed treats the cause.

### 2. Variability in discharge lag matters as much as its average

The discharge lag was fitted by specialty against four candidate distributions (Poisson, Normal censored at zero, Gamma and a hurdle Gamma), with chi square goodness of fit tests. The censored Normal gives the best description. Crucially, the dispersion index is about 2.7, meaning the variance is nearly three times what a Poisson process would produce. In operational terms, discharges do not simply run a little late on average; some run very late, and it is these long tails that create the unpredictable bed shortages that trigger escalation. Interventions that make discharge more consistent, such as standardised criteria led discharge, pharmacy turnaround targets and a protected discharge lounge, should therefore be valued alongside those that reduce the average.

<img width="767" height="363" alt="Screenshot 2026-09-28 at 22 59 15" src="https://github.com/user-attachments/assets/4f481588-4512-4cf6-903c-bac4f1990d41" />


### 3. Roughly a quarter of the annual loss is recoverable with realistic policy changes

The scenario simulator combines seven day discharging, discharge lag reduction, faster bed turnover, a before noon discharge target, staffing and surge levers. At their default settings, these levers recover about 554 thousand dollars a year, which is equivalent to 5.1 staffed beds and roughly 24 percent of the 2.3 million dollar loss. This framing matters for decision making: the hospital can gain the effective capacity of about five beds without capital spend or recruitment for new posts, simply by releasing beds sooner. Because every lever is an input, leadership can test more or less ambitious targets and see the result in beds and dollars instantly.

<img width="685" height="328" alt="Screenshot 2026-09-28 at 23 00 45" src="https://github.com/user-attachments/assets/52221c65-fa41-4b27-b7ac-98dd366227c3" />

### 4. Weekend discharge behaviour builds a predictable Monday backlog

The forecast applies weekend discharge factors when propagating the census, and the effect is visible in the 72 hour horizon: occupancy that looks comfortable on a Friday can tighten sharply by Monday morning because discharges slow while admissions continue. This is the pattern the seven day discharge lever is designed to address, and it is also why the forecast should be read ahead of the weekend rather than on the day.

<img width="689" height="330" alt="Screenshot 2026-09-28 at 23 01 33" src="https://github.com/user-attachments/assets/fffbff4a-3217-4e41-92d9-8ad78c8eedb4" />


### 5. The forecast is reliable at hospital level and honest about unit level limits

Over the weekly backtest, the hospital level forecast lands within 4 points of actual occupancy in 88.2 percent of weeks, with a mean absolute error of 2.05 points. At unit week level the figure falls to 36.1 percent. This gap is a data finding rather than a modelling failure: the extract contains roughly one occupancy reading per unit per day, which is too sparse and noisy to validate unit forecasts tightly. The model is therefore ready to support hospital wide capacity planning now, while unit level breach flags should be treated as directional until better census data is available.

<img width="697" height="306" alt="Screenshot 2026-09-28 at 23 02 21" src="https://github.com/user-attachments/assets/b9d98644-eced-4edd-a8af-b12111c10562" />

### 6. Recorded occupancy is not reconciled to patient movements

Weekly recorded occupancy shows a correlation close to zero with admission volume. In a functioning system, more admissions should push occupancy up, so this indicates that the occupancy field is not reconciled to the ADT feed. This is one of the most important findings for the organisation, because any capacity dashboard built on that field would be unreliable. The recommendation is to source occupancy from a reconciled midnight census before unit level forecasting goes live.

### 7. The model holds up under surge

The three highest admission weeks in the history were replayed through the model as stress tests, and all three ran stably, producing plausible census trajectories and breach probabilities. This gives confidence that the model will remain usable in exactly the conditions where it is most needed.

## Recommendations

1. Treat discharge lag and bed turnover as the primary capacity lever. Set a target for the time from discharge decision to bed release and report it daily by unit.
2. Target variability as well as the average. Focus on the long tail of delayed discharges through criteria led discharge, earlier medication and transport booking, and a discharge lounge.
3. Pilot seven day discharging on the units with the largest weekend backlog and use the simulator to track the recovered bed hours against the modelled value.
4. Establish a reconciled midnight census as the single source of occupancy truth before relying on unit level breach flags.
5. Use the breach risk dashboard as the standard pack for the daily bed meeting, replacing manual census compilation.
6. Add discharge order and bed release timestamps to the data feed so that lag can be measured directly rather than taken from a numeric field.

## Method

1. **Clean the extract.** Eleven documented steps (see docs/data_cleaning_procedure.md). Because the recorded discharge date contradicts length of stay in 94 percent of records, the derived discharge date (admission plus length of stay) is used throughout.
2. **Allocate and calibrate.** The 400 beds are allocated to units by their share of patient days. Daily admissions are calibrated using Little's Law, where census equals arrival rate multiplied by effective length of stay, and effective stay includes discharge lag and bed turnover.
3. **Forecast exactly.** Expected census and its variance are propagated exactly for three days with weekend discharge factors, then converted into a breach probability for each unit.
4. **Cross check by simulation.** A 1,000 trial Monte Carlo engine built from CRITBINOM draws reproduces the forecast; the simulated mean agrees with the exact mean within 2 percent.
5. **Value the problem and the solutions.** Bed unavailable hours are valued at the cost per hour implied by the brief, and policy levers are simulated to show their impact in beds and dollars.
6. **Validate.** A weekly backtest measures forecast accuracy, and the three heaviest admission weeks are replayed as surge stress tests.

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
  <tr><td>Clean Data, Raw Data, Data Dictionary</td><td>Analysis table with named ranges, the supplied data unchanged, and field definitions</td></tr>
</table>

## How to use

Open the workbook in Microsoft Excel 2010 or later; calculation is automatic.

* Change the teal bordered selectors on the dashboard (unit, forecast date and policy mode) to redraw the heat map, fan chart and trend views.
* Adjust the yellow input cells on sheet 05 (policy levers) and sheet M1 (calibration) to test scenarios.
* Formatting convention: blue text on yellow marks inputs, black text marks formulas, and green text marks links between sheets.
* Press F9 to redraw the Monte Carlo trials.

The selectors are in cell drop down lists (data validation), which behave like combo boxes and work in every Excel version. A Form Control combo box can be linked to the same cells if preferred.

## Reproducing the build

The workbook is generated by Python, so every sheet can be rebuilt from the raw data.

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

## Assumptions and limitations

* **Sample scale.** The extract represents about 3,800 encounters a year, fewer than a 400 bed hospital would handle, so rates are expressed as shares and calibrated to the 400 bed base.
* **No clock times.** Admission and discharge times are absent, so discharge lag is taken from the numeric field after validation rather than measured from timestamps.
* **Ambiguous dates.** 15.6 percent of slash dates were resolved as month first and flagged. Excluding them moves the backtest from 88.2 to 85.8 percent, so the conclusions are not sensitive to this choice.
* **Cost calibration.** The cost per bed unavailable hour is implied by the 2.3 million dollar figure in the brief rather than taken from the hospital's ledger.
* **Unit level validation.** Unit forecasts are limited by sparse and unreconciled occupancy readings and should be treated as directional until a reconciled midnight census is available.
* **Meeting time saving.** The 70 percent reduction in bed meeting preparation time is an estimate and would need to be measured after deployment.
