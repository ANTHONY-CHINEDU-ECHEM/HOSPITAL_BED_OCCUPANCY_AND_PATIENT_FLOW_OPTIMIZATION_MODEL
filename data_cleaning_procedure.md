# Data Cleaning Procedure, Project 1

This document mirrors the eleven cleaning sheets (C1 to C11) in the workbook. Every figure below was produced by the same Python engine that wrote those sheets, and each sheet repeats the key counts as live Excel formulas so the reconciliation can be checked inside the workbook.

## Guiding principles

1. No row is deleted unless it is completely empty or an exact copy of an earlier row.
2. No value is changed silently: every changed cell is traceable through a flag or audit column (key_repaired_flag, outlier_fields, imputed_fields, and the cross field flags).
3. Where two fields contradict each other, one is declared authoritative, the decision is written down, and the other is kept for audit.
4. Dates, identifiers, survey answers and legacy model outputs are never imputed.

## Step summary

<table>
  <tr><th>Step</th><th>Sheet</th><th>Action</th><th>Rows in</th><th>Rows out</th><th>Cells changed</th></tr>
  <tr><td>S1</td><td>C2 Structural Integrity</td><td>Removed fully blank rows</td><td>15,394</td><td>15,372</td><td>0</td></tr>
  <tr><td>S2</td><td>C2 Structural Integrity</td><td>Removed exact duplicate records</td><td>15,372</td><td>15,250</td><td>0</td></tr>
  <tr><td>S3</td><td>C3 Key Collision Repair</td><td>Repaired duplicate encounter_id values</td><td>15,250</td><td>15,250</td><td>46</td></tr>
  <tr><td>S4</td><td>C4 Text Standardization</td><td>Collapsed categorical variants to one standard label</td><td>15,250</td><td>15,250</td><td>3,765</td></tr>
  <tr><td>S5</td><td>C5 Boolean Normalization</td><td>Converted 12 token encodings to TRUE/FALSE</td><td>15,250</td><td>15,250</td><td>167,750</td></tr>
  <tr><td>S6</td><td>C6 Date Standardization</td><td>Parsed four mixed date formats to true Excel dates</td><td>15,250</td><td>15,250</td><td>59,368</td></tr>
  <tr><td>S7</td><td>C7 Numeric Validation</td><td>Nullified impossible and extreme values</td><td>15,250</td><td>15,250</td><td>449</td></tr>
  <tr><td>S8</td><td>C8 Cross Field Consistency</td><td>Tested business rules between related fields</td><td>15,250</td><td>15,250</td><td>0</td></tr>
  <tr><td>S9</td><td>C9 Missing Value Treatment</td><td>Filled or explicitly retained every missing value</td><td>15,250</td><td>15,250</td><td>5,452</td></tr>
  <tr><td>S10</td><td>C10 Feature Engineering</td><td>Derived analysis ready features</td><td>15,250</td><td>15,250</td><td>152,500</td></tr>
  <tr><td>S11</td><td>Clean Data</td><td>Published analysis ready table</td><td>15,250</td><td>15,250</td><td>0</td></tr>
</table>

## C1 Data profile

The raw export holds 15,394 rows and 45 columns. Each column was measured for blanks, distinct values and distinct values after normalising case, whitespace, trailing periods and underscores. A gap between those two distinct counts proves the column carries formatting variants.

## C2 Structural integrity

22 rows were completely blank and were removed. 122 rows were exact copies of an earlier row and were removed, keeping the first occurrence. The original Excel row number of every removed row is listed on the sheet.

## C3 Key collision repair

After deduplication, 46 values of encounter_id were still shared by different records. The collision free anchor bed_id carries the true record sequence, so 46 records were re keyed from it instead of being deleted. The live check on the sheet confirms that no key in Clean Data is duplicated.

Related identifiers that repeat for legitimate reasons are flagged but not changed:

* patient_id: 92 records flagged. Flagged in patient_id_conflict_flag. A patient identifier may legitimately repeat across readmissions, but these repeats carry different age and sex, so they are routed for health information management review rather than overwritten.

## C4 Text standardisation

10 categorical columns were standardised. 197 non standard spellings covering 3,765 cells were mapped to one governed label. Each raw value was normalised (trim, drop trailing period, underscores to spaces, collapse spaces, ignore case) and mapped to the most frequent clean spelling. Identifier columns were trimmed and upper cased.

<table>
  <tr><th>Issue type</th><th>Variants</th><th>Cells</th></tr>
  <tr><td>Inconsistent letter case</td><td>88</td><td>2,095</td></tr>
  <tr><td>Leading or trailing whitespace</td><td>45</td><td>790</td></tr>
  <tr><td>Trailing period</td><td>45</td><td>662</td></tr>
  <tr><td>Underscore used as separator</td><td>19</td><td>218</td></tr>
</table>

## C5 Boolean normalisation

11 flag columns used twelve encodings of yes and no (true, TRUE, yes, Yes, Y, 1 and their negatives). They were converted to native TRUE and FALSE. No unrecognised tokens were found.

## C6 Date standardisation

Dates arrived as text in four formats. Unambiguous patterns were parsed directly; slash dates where both parts are 12 or below were resolved as month first (the convention shown in the data dictionary) and flagged in an _ambiguous_flag column.

<table>
  <tr><th>Column</th><th>Detected pattern</th><th>Records</th><th>Parsed</th></tr>
  <tr><td>admission_date</td><td>ISO 8601 (YYYY MM DD)</td><td>3,015</td><td>3,015</td></tr>
  <tr><td>admission_date</td><td>Day Month abbrev (DD Mon YYYY)</td><td>3,010</td><td>3,010</td></tr>
  <tr><td>admission_date</td><td>Slash ISO (YYYY/MM/DD)</td><td>2,988</td><td>2,988</td></tr>
  <tr><td>admission_date</td><td>Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY</td><td>2,374</td><td>2,374</td></tr>
  <tr><td>admission_date</td><td>MM/DD/YYYY (day > 12, unambiguous)</td><td>1,846</td><td>1,846</td></tr>
  <tr><td>admission_date</td><td>DD/MM/YYYY (day > 12, unambiguous)</td><td>1,834</td><td>1,834</td></tr>
  <tr><td>admission_date</td><td>Missing</td><td>183</td><td>0</td></tr>
  <tr><td>discharge_date</td><td>ISO 8601 (YYYY MM DD)</td><td>2,987</td><td>2,987</td></tr>
  <tr><td>discharge_date</td><td>Slash ISO (YYYY/MM/DD)</td><td>2,961</td><td>2,961</td></tr>
  <tr><td>discharge_date</td><td>Day Month abbrev (DD Mon YYYY)</td><td>2,928</td><td>2,928</td></tr>
  <tr><td>discharge_date</td><td>Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY</td><td>2,406</td><td>2,406</td></tr>
  <tr><td>discharge_date</td><td>MM/DD/YYYY (day > 12, unambiguous)</td><td>1,841</td><td>1,841</td></tr>
  <tr><td>discharge_date</td><td>DD/MM/YYYY (day > 12, unambiguous)</td><td>1,746</td><td>1,746</td></tr>
  <tr><td>discharge_date</td><td>Missing</td><td>381</td><td>0</td></tr>
  <tr><td>discharge_order_time</td><td>ISO 8601 (YYYY MM DD)</td><td>3,068</td><td>3,068</td></tr>
  <tr><td>discharge_order_time</td><td>Slash ISO (YYYY/MM/DD)</td><td>2,981</td><td>2,981</td></tr>
  <tr><td>discharge_order_time</td><td>Day Month abbrev (DD Mon YYYY)</td><td>2,879</td><td>2,879</td></tr>
  <tr><td>discharge_order_time</td><td>Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY</td><td>2,384</td><td>2,384</td></tr>
  <tr><td>discharge_order_time</td><td>MM/DD/YYYY (day > 12, unambiguous)</td><td>1,745</td><td>1,745</td></tr>
  <tr><td>discharge_order_time</td><td>DD/MM/YYYY (day > 12, unambiguous)</td><td>1,735</td><td>1,735</td></tr>
  <tr><td>discharge_order_time</td><td>Missing</td><td>458</td><td>0</td></tr>
  <tr><td>actual_discharge_time</td><td>Day Month abbrev (DD Mon YYYY)</td><td>2,962</td><td>2,962</td></tr>
  <tr><td>actual_discharge_time</td><td>ISO 8601 (YYYY MM DD)</td><td>2,931</td><td>2,931</td></tr>
  <tr><td>actual_discharge_time</td><td>Slash ISO (YYYY/MM/DD)</td><td>2,927</td><td>2,927</td></tr>
  <tr><td>actual_discharge_time</td><td>Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY</td><td>2,344</td><td>2,344</td></tr>
  <tr><td>actual_discharge_time</td><td>MM/DD/YYYY (day > 12, unambiguous)</td><td>1,752</td><td>1,752</td></tr>
  <tr><td>actual_discharge_time</td><td>DD/MM/YYYY (day > 12, unambiguous)</td><td>1,724</td><td>1,724</td></tr>
  <tr><td>actual_discharge_time</td><td>Missing</td><td>610</td><td>0</td></tr>
</table>

## C7 Numeric validation

Each numeric column was tested against physical or definitional limits and, where no hard maximum exists, against the Tukey outer fence (Q3 plus 3 x IQR). Failing cells were blanked (never the whole row) and then treated in C9.

<table>
  <tr><th>Column</th><th>Unit</th><th>Minimum</th><th>Applied maximum</th><th>Below minimum</th><th>Above maximum</th><th>Rationale</th></tr>
  <tr><td>patient_age</td><td>years</td><td>0.00</td><td>110.00</td><td>0</td><td>0</td><td>Human age range; no fence because age has a natural maximum</td></tr>
  <tr><td>length_of_stay_days</td><td>days</td><td>0.00</td><td>19.34</td><td>74</td><td>77</td><td>Negative stays impossible; sentinel 25.6 and inflated stays caught by fence</td></tr>
  <tr><td>unit_capacity_beds</td><td>beds</td><td>1.00</td><td>120.00</td><td>0</td><td>0</td><td>A staffed unit needs at least one bed</td></tr>
  <tr><td>beds_occupied_at_admission</td><td>beds</td><td>0.00</td><td>120.00</td><td>0</td><td>0</td><td>Count cannot be negative</td></tr>
  <tr><td>occupancy_rate_pct</td><td>%</td><td>0.00</td><td>150.00</td><td>0</td><td>149</td><td>Values of 200 to 521 are magnitude errors; 150 allows documented surge overflow</td></tr>
  <tr><td>discharge_lag_hours</td><td>hours</td><td>0.00</td><td>25.37</td><td>75</td><td>74</td><td>Negative lag impossible; sentinel 33.6 caught by fence</td></tr>
  <tr><td>ed_boarding_hours</td><td>hours</td><td>0.00</td><td>19.14</td><td>0</td><td>0</td><td>Boarding cannot be negative</td></tr>
  <tr><td>acuity_score</td><td>score</td><td>1.00</td><td>5.00</td><td>0</td><td>0</td><td>Scale defined as 1 to 5</td></tr>
  <tr><td>nurse_to_patient_ratio</td><td>patients per nurse</td><td>1.00</td><td>12.00</td><td>0</td><td>0</td><td>Operational range</td></tr>
  <tr><td>discharge_planning_start_day</td><td>hospital day</td><td>0.00</td><td>60.00</td><td>0</td><td>0</td><td>Cannot precede admission</td></tr>
  <tr><td>housekeeping_turnover_minutes</td><td>minutes</td><td>0.00</td><td>117.07</td><td>0</td><td>0</td><td>Negative turnover impossible</td></tr>
  <tr><td>predicted_occupancy_next_72h_pct</td><td>%</td><td>0.00</td><td>150.00</td><td>0</td><td>0</td><td>Legacy model output, same bound as occupancy</td></tr>
  <tr><td>forecast_variance_pct</td><td>% points</td><td>minus 100.00</td><td>100.00</td><td>0</td><td>0</td><td>Signed variance bounded by the percentage scale</td></tr>
  <tr><td>staffing_level_ratio</td><td>ratio</td><td>0.20</td><td>2.50</td><td>0</td><td>0</td><td>Actual over budgeted staffing</td></tr>
</table>

## C8 Cross field consistency

<table>
  <tr><th>Rule</th><th>Records failing</th><th>Fail rate</th><th>Decision</th></tr>
  <tr><td>Discharge before admission</td><td>7,329</td><td>48.1 percent</td><td>Recorded discharge_date is not reliable. length_of_stay_days (a validated numeric field) is authoritative; discharge_date_derived = admission_date + LOS is used in every model.</td></tr>
  <tr><td>Recorded dates contradict length of stay</td><td>14,303</td><td>93.8 percent</td><td>Confirms the decision above: the recorded discharge date field is independent of the stay and is retained only for audit.</td></tr>
  <tr><td>Occupied beds exceed capacity</td><td>6,071</td><td>39.8 percent</td><td>occupancy_rate_pct (the unit census metric reported to the bed meeting) is authoritative. beds_occupied_at_admission is not used in modelling.</td></tr>
  <tr><td>Occupancy rate disagrees with bed counts</td><td>13,705</td><td>89.9 percent</td><td>Same resolution as above. Capacity for modelling comes from the 400 bed allocation on M1 Unit Calibration.</td></tr>
  <tr><td>Bed left before discharge order</td><td>7,091</td><td>46.5 percent</td><td>Both fields hold dates only, with no clock time, so hour level lag cannot be derived. discharge_lag_hours is authoritative for all lag analysis.</td></tr>
  <tr><td>Paediatric unit with adult patient</td><td>1,216</td><td>8.0 percent</td><td>Retained and flagged for clinical documentation review. Unit flow metrics (LOS, lag, occupancy) are unaffected by patient age, so no exclusion is applied.</td></tr>
  <tr><td>Labor and Delivery with male patient</td><td>687</td><td>4.5 percent</td><td>Retained and included in clinical_plausibility_flag for review; same reasoning as above.</td></tr>
</table>

## C9 Missing value treatment

<table>
  <tr><th>Column</th><th>Blanks before</th><th>Strategy</th><th>Filled</th><th>Blanks after</th></tr>
  <tr><td>patient_age</td><td>274</td><td>Median within facility_unit (overall median if group empty)</td><td>274</td><td>0</td></tr>
  <tr><td>length_of_stay_days</td><td>380</td><td>Median within facility_unit (overall median if group empty)</td><td>380</td><td>0</td></tr>
  <tr><td>occupancy_rate_pct</td><td>454</td><td>Median within facility_unit (overall median if group empty)</td><td>454</td><td>0</td></tr>
  <tr><td>discharge_lag_hours</td><td>454</td><td>Median within facility_unit (overall median if group empty)</td><td>454</td><td>0</td></tr>
  <tr><td>ed_boarding_hours</td><td>458</td><td>Median within admission_source (overall median if group empty)</td><td>458</td><td>0</td></tr>
  <tr><td>acuity_score</td><td>305</td><td>Median within facility_unit (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>discharge_planning_start_day</td><td>458</td><td>Median within facility_unit (overall median if group empty)</td><td>458</td><td>0</td></tr>
  <tr><td>housekeeping_turnover_minutes</td><td>305</td><td>Median within bed_type (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>beds_occupied_at_admission</td><td>305</td><td>Left blank: field superseded by occupancy_rate_pct (see C8)</td><td>0</td><td>305</td></tr>
  <tr><td>predicted_occupancy_next_72h_pct</td><td>305</td><td>Left blank: legacy model output; imputing a forecast would fabricate accuracy</td><td>0</td><td>305</td></tr>
  <tr><td>forecast_variance_pct</td><td>305</td><td>Left blank: legacy model error; imputing would fabricate accuracy</td><td>0</td><td>305</td></tr>
  <tr><td>patient_gender</td><td>152</td><td>Explicit category 'Unknown'</td><td>152</td><td>0</td></tr>
  <tr><td>race_ethnicity</td><td>458</td><td>Explicit category 'Unknown'</td><td>458</td><td>0</td></tr>
  <tr><td>primary_diagnosis_code</td><td>305</td><td>Explicit category 'Not Recorded'</td><td>305</td><td>0</td></tr>
  <tr><td>procedure_code</td><td>915</td><td>Explicit category 'Not Recorded'</td><td>915</td><td>0</td></tr>
  <tr><td>discharge_disposition</td><td>305</td><td>Explicit category 'Not Recorded'</td><td>305</td><td>0</td></tr>
  <tr><td>attending_physician</td><td>229</td><td>Explicit category 'Not Recorded'</td><td>229</td><td>0</td></tr>
  <tr><td>zip_code</td><td>305</td><td>Left blank: location cannot be inferred</td><td>0</td><td>305</td></tr>
  <tr><td>admission_date</td><td>183</td><td>Left blank: event dates are never imputed; record excluded from time series only</td><td>0</td><td>183</td></tr>
  <tr><td>discharge_date</td><td>381</td><td>Left blank: superseded by discharge_date_derived</td><td>0</td><td>381</td></tr>
</table>

## C10 Feature engineering

Derived fields are documented on the C10 sheet with their definition, worksheet equivalent and a live populated count.

## C11 Audit log and reconciliation

The final sheet lists every step with rows in, rows out and cells changed, and reconciles live: rows in Raw Data, less blank rows, less exact duplicates, must equal rows in Clean Data. The check returns PASS in the delivered workbook.
