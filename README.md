# Cesim Global Challenge — Financial Analysis Dashboard

A multi-round financial analysis dashboard built for the **Cesim Global Challenge** business simulation. Upload your round `.xls` files and instantly get ratios, valuation models, risk indicators, and regional breakdowns — all in one interactive interface.

Built with Python & Streamlit as part of the KU Leuven Business Game course (BBA, Year 2).

---

## Features

### Core analysis tabs
- **Profitability** — EBITDA margin, ROS, ROCE, ROE, ROA
- **Liquidity & Solvency** — Current ratio, quick ratio, gearing, interest coverage
- **Valuation** — P/E, P/B, P/Sales, EV/EBITDA, EV/Sales, DDM (zero-growth & Gordon Growth), Residual Income Model
- **Capital & Risk** — WACC, Altman Z-Score
- **Growth & Trends** — YoY growth, Sustainable Growth Rate (SGR), multi-round trend charts
- **Regional Analysis** — Income statements, balance sheets, market shares, and margin comparisons across USA, Europe, and China
- **CFO Cheat Sheet** — Formula reference with benchmarks and a quick decision framework

### Key metrics computed
| Category | Metrics |
|---|---|
| Valuation | P/E, P/B, EV/EBITDA, EV/Sales, DDM, RIM |
| Risk | WACC, Altman Z-Score |
| Growth | SGR, YoY revenue & profit growth |
| Regional | Revenue contribution %, regional EBIT margins, market share by vehicle type |

---

## Tech Stack

- **Python 3.x**
- **Streamlit** — interactive web dashboard
- **Plotly** — charts and visualizations
- **Pandas / NumPy** — data processing
- **xlrd / openpyxl** — Excel file parsing

---

## How to Run

**1. Install dependencies**
```bash
pip install streamlit plotly xlrd pandas openpyxl numpy
```

**2. Run the dashboard**
```bash
streamlit run Cesim_dashboard_v3_1_0_fixed.py
```

**3. Upload your data**

In the sidebar, upload one or more Cesim round `.xls` export files. The dashboard will automatically parse and compute all metrics.

---

## Formula Reference

| Model | Formula |
|---|---|
| WACC | (E/V)×Re + (D/V)×Rd×(1−t) |
| Altman Z-Score | Public-firm model (1968); X4 uses market cap / book liabilities |
| SGR | ROE × (1 − payout ratio) |
| DDM (zero-growth) | P = DPS / Re |
| DDM (Gordon Growth) | P = DPS×(1+g) / (Re−g) |
| Residual Income Model | IV = BVps + ((EPS×(1+g)) − Re×BVps) / (Re−g) |
| Enterprise Value | Market Cap + LT Debt + ST Debt − Cash |

---

## Version History

| Version | Key additions |
|---|---|
| v1 | Basic ratio dashboard |
| v2 | Valuation tab, WACC, Altman Z-Score, DDM, multi-round upload |
| v3 | Regional analysis (USA / Europe / China), revenue contribution charts, heatmap P&L tables |
| v3.1 | Bug fixes and stability improvements |

---

## Context

This dashboard was developed during the **Cesim Global Challenge** simulation at KU Leuven, where I served as **CFO** of team ThunderVolt. It was used to support round-by-round financial decisions and competitive benchmarking against rival teams.

---

## Author

**Khai Tran** — BBA student at KU Leuven  
[GitHub](https://github.com/KhaiTranKUL)
