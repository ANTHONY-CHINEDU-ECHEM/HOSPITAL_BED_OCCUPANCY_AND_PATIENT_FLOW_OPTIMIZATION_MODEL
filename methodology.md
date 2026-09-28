# Methodology, Project 1

## Bed base and calibration
Staffed beds (400) are allocated to the ten units by share of patient days using the largest remainder method. For each unit the effective length of stay is length of stay plus discharge lag plus housekeeping turnover. Little's Law gives the admission rate consistent with observed occupancy: lambda equals census divided by effective length of stay.

## 72 hour forecast
Starting census is staffed beds times mean recorded occupancy over the 28 days before the forecast date. Each day admissions are Poisson with rate lambda and each occupied bed is discharged with probability p equal to one over effective length of stay, adjusted by a weekend factor normalised so the weekly average is unchanged. Mean and variance are propagated exactly: E(d) = E(d minus 1) x (1 minus p) + lambda and Var(d) = Var(d minus 1) x (1 minus p) squared + E(d minus 1) x p x (1 minus p) + lambda. Breach probability uses a normal approximation.

## Monte Carlo cross check
1,000 trials draw admissions and discharges with CRITBINOM. The simulated 72 hour mean must agree with the exact mean within 2 percent.

## Discharge lag fitting
Lags excluding imputed values are binned by hour. Poisson, Normal censored at zero, Gamma (method of moments) and a hurdle Gamma model are fitted and compared by chi square per degree of freedom using bins with expected counts of at least 5.

## Scenario economics
Annual bed unavailable hours are valued at the cost per hour implied by the brief's 2.3 million dollars. Levers shorten lag and turnover or move discharges before noon; the change in effective length of stay changes census through Little's Law.

## Validation
Weekly hospital backtest with a 4 week rolling forecast, surge week replay, Monte Carlo agreement, bed reconciliation and cleaning reconciliation.
