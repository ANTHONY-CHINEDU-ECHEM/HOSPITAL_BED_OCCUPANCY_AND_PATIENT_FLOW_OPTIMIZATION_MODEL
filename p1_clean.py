"""Project 1: Hospital Bed Occupancy and Patient Flow. Cleaning pipeline."""
import glob
import numpy as np
import pandas as pd
from lib.cleaning import Cleaner

SRC = sorted(glob.glob("data/raw/*.xlsx"))[0]

BOOL = ["bed_block_flag", "transfer_in_flag", "transfer_out_flag", "isolation_precaution_flag", "case_manager_assigned_flag",
        "readmission_30day_flag", "surge_day_flag", "weekend_admission_flag", "holiday_flag", "breach_risk_flag",
        "discharge_before_noon_flag"]
CFG = dict(
    key="encounter_id", anchor="bed_id", key_prefix="ENC-", key_width=7,
    flag_identity=[("patient_id", "Flagged in patient_id_conflict_flag. A patient identifier may legitimately repeat across readmissions, "
                                  "but these repeats carry different age and sex, so they are routed for health information management review rather than overwritten.")],
    categorical=["patient_gender", "race_ethnicity", "insurance_type", "facility_unit", "admission_source", "discharge_disposition",
                 "physician_team", "bed_type", "primary_diagnosis_code", "attending_physician"],
    id_text=["patient_id", "encounter_id", "bed_id", "monte_carlo_scenario_id"],
    zip=["zip_code"], code_text=["procedure_code"],
    bool=BOOL,
    dates=["admission_date", "discharge_date", "discharge_order_time", "actual_discharge_time"],
    numeric={
        "patient_age": dict(lo=0, hi=110, fence=False, unit="years", integer=True, rationale="Human age range; no fence because age has a natural maximum"),
        "length_of_stay_days": dict(lo=0, hi=90, unit="days", rationale="Negative stays impossible; sentinel 25.6 and inflated stays caught by fence"),
        "unit_capacity_beds": dict(lo=1, hi=120, fence=False, unit="beds", integer=True, rationale="A staffed unit needs at least one bed"),
        "beds_occupied_at_admission": dict(lo=0, hi=120, fence=False, unit="beds", integer=True, rationale="Count cannot be negative"),
        "occupancy_rate_pct": dict(lo=0, hi=150, unit="%", rationale="Values of 200 to 521 are magnitude errors; 150 allows documented surge overflow"),
        "discharge_lag_hours": dict(lo=0, hi=72, unit="hours", rationale="Negative lag impossible; sentinel 33.6 caught by fence"),
        "ed_boarding_hours": dict(lo=0, hi=72, unit="hours", rationale="Boarding cannot be negative"),
        "acuity_score": dict(lo=1, hi=5, fence=False, unit="score", integer=True, rationale="Scale defined as 1 to 5"),
        "nurse_to_patient_ratio": dict(lo=1, hi=12, fence=False, unit="patients per nurse", rationale="Operational range"),
        "discharge_planning_start_day": dict(lo=0, hi=60, fence=False, unit="hospital day", integer=True, rationale="Cannot precede admission"),
        "housekeeping_turnover_minutes": dict(lo=0, hi=240, unit="minutes", rationale="Negative turnover impossible"),
        "predicted_occupancy_next_72h_pct": dict(lo=0, hi=150, fence=False, unit="%", rationale="Legacy model output, same bound as occupancy"),
        "forecast_variance_pct": dict(lo=-100, hi=100, fence=False, unit="% points", rationale="Signed variance bounded by the percentage scale"),
        "staffing_level_ratio": dict(lo=0.2, hi=2.5, fence=False, unit="ratio", rationale="Actual over budgeted staffing"),
    },
    impute={
        "patient_age": ("median", "facility_unit"),
        "length_of_stay_days": ("median", "facility_unit"),
        "occupancy_rate_pct": ("median", "facility_unit"),
        "discharge_lag_hours": ("median", "facility_unit"),
        "ed_boarding_hours": ("median", "admission_source"),
        "acuity_score": ("median", "facility_unit"),
        "discharge_planning_start_day": ("median", "facility_unit"),
        "housekeeping_turnover_minutes": ("median", "bed_type"),
        "beds_occupied_at_admission": ("leave", "field superseded by occupancy_rate_pct (see C8)"),
        "predicted_occupancy_next_72h_pct": ("leave", "legacy model output; imputing a forecast would fabricate accuracy"),
        "forecast_variance_pct": ("leave", "legacy model error; imputing would fabricate accuracy"),
        "patient_gender": ("constant", "Unknown"),
        "race_ethnicity": ("constant", "Unknown"),
        "primary_diagnosis_code": ("constant", "Not Recorded"),
        "procedure_code": ("constant", "Not Recorded"),
        "discharge_disposition": ("constant", "Not Recorded"),
        "attending_physician": ("constant", "Not Recorded"),
        "zip_code": ("leave", "location cannot be inferred"),
        "admission_date": ("leave", "event dates are never imputed; record excluded from time series only"),
        "discharge_date": ("leave", "superseded by discharge_date_derived"),
    },
)

FEATURES = [
    ("discharge_date_derived", "admission_date + length_of_stay_days (cleaned). Replaces the recorded discharge date, which contradicts length of stay (C8).", "=admission_date + length_of_stay_days"),
    ("admit_week_start", "Monday of the admission week", "=admission_date - WEEKDAY(admission_date,3)"),
    ("admit_month", "First day of the admission month", "=DATE(YEAR(d),MONTH(d),1)"),
    ("admit_year", "Calendar year of admission", "=YEAR(admission_date)"),
    ("admit_dow", "Day of week of admission", "=TEXT(admission_date,\"ddd\")"),
    ("discharge_dow", "Day of week of the derived discharge", "=TEXT(discharge_date_derived,\"ddd\")"),
    ("discharge_lag_int", "Discharge lag rounded to whole hours, the unit used for distribution fitting", "=ROUND(discharge_lag_hours,0)"),
    ("bed_unavailable_hours", "Discharge lag plus housekeeping turnover: total hours a bed is blocked after the clinical decision", "=discharge_lag_hours + housekeeping_turnover_minutes/60"),
    ("los_band", "Length of stay band for mix analysis", "=LOOKUP(los,{0,1,3,7,14},{\"0 to 1\",\"1 to 3\",\"3 to 7\",\"7 to 14\",\"14 plus\"})"),
    ("occupancy_breach_flag", "Unit occupancy at admission at or above the 85 percent safe occupancy threshold", "=occupancy_rate_pct >= 85"),
]


def run():
    raw = pd.read_excel(SRC, sheet_name="Dataset", dtype=object)
    dd = pd.read_excel(SRC, sheet_name="Data Dictionary", header=None, dtype=object)
    cl = Cleaner(raw, CFG)
    cl.profile(); cl.structural(); cl.key_collisions(); cl.text(); cl.booleans(); cl.dates(); cl.numerics()
    N = lambda s: pd.to_numeric(s, errors="coerce")
    checks = [
        dict(Check="Discharge before admission", Description="Recorded discharge_date is earlier than admission_date",
             mask=lambda d: d.discharge_date < d.admission_date, flag="date_sequence_conflict_flag",
             Decision="Recorded discharge_date is not reliable. length_of_stay_days (a validated numeric field) is authoritative; discharge_date_derived = admission_date + LOS is used in every model."),
        dict(Check="Recorded dates contradict length of stay", Description="Days between recorded admission and discharge differ from length_of_stay_days by more than 1 day",
             mask=lambda d: ((d.discharge_date - d.admission_date).dt.days - d.length_of_stay_days).abs() > 1,
             Decision="Confirms the decision above: the recorded discharge date field is independent of the stay and is retained only for audit."),
        dict(Check="Occupied beds exceed capacity", Description="beds_occupied_at_admission greater than unit_capacity_beds",
             mask=lambda d: N(d.beds_occupied_at_admission) > N(d.unit_capacity_beds),
             Decision="occupancy_rate_pct (the unit census metric reported to the bed meeting) is authoritative. beds_occupied_at_admission is not used in modelling."),
        dict(Check="Occupancy rate disagrees with bed counts", Description="|occupancy_rate_pct minus 100 x occupied / capacity| greater than 5 points",
             mask=lambda d: (N(d.occupancy_rate_pct) - 100 * N(d.beds_occupied_at_admission) / N(d.unit_capacity_beds)).abs() > 5,
             Decision="Same resolution as above. Capacity for modelling comes from the 400 bed allocation on M1 Unit Calibration."),
        dict(Check="Bed left before discharge order", Description="actual_discharge_time earlier than discharge_order_time",
             mask=lambda d: d.actual_discharge_time < d.discharge_order_time,
             Decision="Both fields hold dates only, with no clock time, so hour level lag cannot be derived. discharge_lag_hours is authoritative for all lag analysis."),
        dict(Check="Paediatric unit with adult patient", Description="facility_unit = Pediatrics and patient_age 18 or over",
             mask=lambda d: (d.facility_unit == "Pediatrics") & (N(d.patient_age) >= 18), flag="clinical_plausibility_flag",
             Decision="Retained and flagged for clinical documentation review. Unit flow metrics (LOS, lag, occupancy) are unaffected by patient age, so no exclusion is applied."),
        dict(Check="Labor and Delivery with male patient", Description="facility_unit = Labor & Delivery and patient_gender = Male",
             mask=lambda d: (d.facility_unit == "Labor & Delivery") & (d.patient_gender == "Male"),
             Decision="Retained and included in clinical_plausibility_flag for review; same reasoning as above."),
    ]
    cl.cross_field(checks)
    df = cl.df
    df["clinical_plausibility_flag"] = df["clinical_plausibility_flag"] | ((df.facility_unit == "Labor & Delivery") & (df.patient_gender == "Male"))
    cl.missing()
    df = cl.df
    # features
    df["discharge_date_derived"] = df.admission_date + pd.to_timedelta(df.length_of_stay_days, unit="D")
    df["admit_week_start"] = df.admission_date - pd.to_timedelta(df.admission_date.dt.weekday, unit="D")
    df["admit_month"] = df.admission_date.dt.to_period("M").dt.to_timestamp()
    df["admit_year"] = df.admission_date.dt.year
    df["admit_dow"] = df.admission_date.dt.strftime("%a")
    df["discharge_dow"] = df.discharge_date_derived.dt.strftime("%a")
    df["discharge_lag_int"] = df.discharge_lag_hours.round()
    df["bed_unavailable_hours"] = df.discharge_lag_hours + df.housekeeping_turnover_minutes / 60
    df["los_band"] = pd.cut(df.length_of_stay_days, [-0.01, 1, 3, 7, 14, 999], labels=["0 to 1", "1 to 3", "3 to 7", "7 to 14", "14 plus"]).astype(str)
    df["occupancy_breach_flag"] = df.occupancy_rate_pct >= 85
    cl._log("S10", "C10 Feature Engineering", "Derived analysis ready features", len(df), len(df), len(FEATURES) * len(df), f"{len(FEATURES)} features added")
    cl._log("S11", "Clean Data", "Published analysis ready table", len(df), len(df), 0, "Clean Data sheet, with named ranges for model formulas")
    cl.df = df
    return raw, dd, cl


if __name__ == "__main__":
    raw, dd, cl = run()
    print(cl.df.shape)
    print(pd.DataFrame(cl.audit)[["Step", "Action", "RowsIn", "RowsOut", "CellsChanged"]])
    print(cl.numeric_rules[["Column", "TukeyFence", "AppliedMax", "BelowMin", "AboveMax"]])
    print(cl.cross[["Check", "RecordsFailing", "FailRate"]])
    print(cl.missing_df)
    print(cl.date_summary)
    print("normalised dups", cl.normalised_dups)
    print(cl.df.facility_unit.value_counts())
