# 📊 DCF Valuation Model — Python + Excel

A professional-grade **Discounted Cash Flow (DCF)** model built entirely in Python. Running the script generates a fully formatted, formula-driven Excel workbook that mirrors the layout used by Wall Street analysts.

---

## 🖼️ Preview

| Section | What it shows |
|---|---|
| **Summary Bar** | Implied vs. today's share price + upside/(downside) |
| **Assumptions** | WACC and Terminal Growth Rate |
| **Income Statement** | Revenue, EBIT, Taxes — historical (2021–2025) + projected (2026–2030) |
| **Cash Flow Items** | D&A, CapEx, Δ NWC — same time span |
| **DCF Analysis** | NOPAT, Unlevered FCF, Present Value of FCF |
| **Valuation Bridge** | Terminal Value → Enterprise Value → Equity Value → **Share Price** |

---

## 🚀 Quick Start

```bash
# 1. Clone the repo
git clone https://github.com/<your-username>/dcf-model.git
cd dcf-model

# 2. Install the single dependency
pip install -r requirements.txt

# 3. Run
python dcf_model.py
# → CMG_DCF_Model.xlsx is created in the current directory
```

Open the generated `.xlsx` in Excel or Google Sheets — all formulas recalculate automatically.

---

## ⚙️ Customizing for Any Company

Everything you need to change lives in the **INPUTS** block at the top of `dcf_model.py`:

```python
TICKER        = "CMG"          # ← your ticker
COMPANY_NAME  = "Chipotle"
TODAY_PRICE   = 37.97          # current share price
SHARES        = 1_300_000      # diluted shares (thousands)
CASH          = 1_050_000      # cash on balance sheet (thousands)
DEBT          = 5_080_000      # total debt (thousands)
WACC          = 0.0755         # your WACC estimate
TGR           = 0.060          # terminal growth rate
```

Swap in real historical figures from the company's 10-K, update the projection drivers, and re-run.

---

## 🎨 Color Conventions (Industry Standard)

| Color | Meaning |
|---|---|
| 🔵 **Blue** | Hardcoded inputs — change these to run scenarios |
| ⚫ **Black** | Calculated formulas — auto-update when inputs change |
| 🟡 **Yellow highlight** | Key output (implied share price) |
| ⬜ **Grey background** | Historical data |

---

## 🧮 Model Logic

```
Revenue (projected)  = Prior Year Revenue × (1 + Growth Rate)
EBIT                 = Revenue × EBIT Margin
Taxes                = EBIT × Tax Rate
NOPAT (PAR)          = EBIT − Taxes
D&A / CapEx / ΔNWC   = Revenue × respective % driver
Unlevered FCF        = NOPAT + D&A − CapEx − ΔNWC
PV of FCF            = Unlevered FCF ÷ (1 + WACC)ⁿ
Terminal Value       = FCF₅ × (1 + TGR) ÷ (WACC − TGR)
PV of Terminal Value = Terminal Value ÷ (1 + WACC)⁵
Enterprise Value     = Σ PV of FCFs + PV of Terminal Value
Equity Value         = Enterprise Value + Cash − Debt
Implied Share Price  = Equity Value ÷ Shares Outstanding
```

---

## 📁 File Structure

```
dcf-model/
├── dcf_model.py       # Main script — edit INPUTS here
├── requirements.txt   # pip dependencies (just openpyxl)
└── README.md
```

---

## 📦 Requirements

- Python 3.8+
- openpyxl ≥ 3.1.0

---

## 📚 Where to Find Inputs

| Input | Source |
|---|---|
| Historical Revenue, EBIT, Taxes | Company 10-K (SEC EDGAR) |
| D&A, CapEx | Cash Flow Statement in 10-K |
| Shares Outstanding | Cover page of 10-K or 10-Q |
| Cash & Debt | Balance Sheet in 10-K |
| WACC | Damodaran's website (pages.stern.nyu.edu) or Bloomberg |
| Projection % Drivers | Analyst consensus (FactSet, Bloomberg) or your own research |

---

*Built as a finance portfolio project. Not investment advice.*
