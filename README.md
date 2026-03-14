# RBI Repo Rate Projection Model

13-indicator weighted scoring framework for forecasting RBI monetary policy decisions.

## Model Overview

Projects likely RBI repo rate direction using a multi-factor scoring system calibrated against MPC voting patterns since 2020.

### Indicators & Weights

| Indicator | Weight | Direction |
|-----------|--------|-----------|
| CPI Inflation (YoY) | 18% | Inverse |
| Core CPI | 10% | Inverse |
| GDP Growth | 12% | Direct |
| CAD (% of GDP) | 8% | Inverse |
| Crude Oil (Brent) | 7% | Inverse |
| Fed Funds Rate | 6% | Direct |
| LAF Liquidity | 6% | Inverse |
| MIBOR-OIS Spread | 5% | Inverse |
| USD/INR | 5% | Inverse |
| WPI Inflation | 5% | Inverse |
| FII Net Flows | 4% | Direct |
| IIP Growth | 4% | Direct |
| Forex Reserves | 4% | Direct |

## Excel Sheets

1. **Dashboard** — Current state, MPC projections, 4-scenario analysis
2. **Indicator Scoring** — Enter -5 to +5 scores per indicator (yellow cells)
3. **Historical Data** — MPC decisions Feb 2020 → Feb 2026 with chart
4. **Assumptions** — Input latest macro data before each MPC meeting
5. **Weight Calibration** — Rationale and predictive power for each weight
6. **MPC Calendar** — FY25-26 schedule with pre-meeting action triggers

## Usage

1. Open **Assumptions** sheet → update with latest macro data
2. Open **Indicator Scoring** → enter scores from -5 (strongly dovish) to +5 (strongly hawkish)
3. Weighted score auto-calculates
4. Interpretation:
   - ≤ -0.30: Strongly dovish → expect 50bp+ cut
   - -0.30 to -0.10: Moderately dovish → expect 25bp cut
   - -0.10 to +0.10: Neutral → hold likely
   - +0.10 to +0.30: Moderately hawkish → possible hike
   - ≥ +0.30: Strongly hawkish → expect 25-50bp hike

## Data Sources

- CPI / WPI / IIP: MOSPI (mospi.gov.in)
- GDP: NSO advance/revised estimates
- CAD / Forex / LAF: RBI Weekly Statistical Supplement
- FII Flows: NSDL (nsdl.co.in)
- MIBOR / OIS: FIMMDA / Bloomberg
- Fed Funds Rate: federalreserve.gov
- Crude Oil: Bloomberg / Reuters
- MPC Decisions: RBI press releases

## Regenerate Model

```bash
python3 rbi_repo_projection.py
```

## License

MIT
