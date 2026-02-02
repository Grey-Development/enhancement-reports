# Enhancement Jobs Monthly Report Generator

**Purpose:** Automated generation of monthly Enhancement & Contracted Enhancement financial analysis reports for Lawn Capital.

**Owner:** Darian - Grey Development / Lawn Capital
**Last Updated:** February 1, 2026

---

## Report Scope

This agent generates financial performance reports filtered exclusively to:
- **Enhancement** jobs
- **Contracted Enhancement** jobs

All other job types (Open Issue, etc.) are excluded from analysis.

---

## Data Sources

### Required Input Files (from Jobber exports)

1. **Jobs Report:** `One-off jobs_Report_1_of_1_YYYY-MM-DD.csv`
   - Location: `C:\Users\daria\Downloads\`
   - Export from: Jobber > Reports > One-off Jobs
   - Required columns: Job #, Client name, Scheduled start date, Closed date, Salesperson, Visits assigned to, Invoice #s, Expenses total ($), Time tracked, Labour cost total ($), Total costs ($), Total revenue ($), Profit ($), Profit %, Job Type

2. **Invoice Report:** `Invoices_Report_1_of_1_YYYY-MM-DD.csv`
   - Location: `C:\Users\daria\Downloads\`
   - Export from: Jobber > Reports > Invoices
   - Required columns: Invoice #, Client name, Created date, Issued date, Due date, Late by, Marked paid date, Days to paid, Last contacted, Job #s, Status, Total ($), Balance ($)

---

## Output File Convention

**Location:** `C:\Users\daria\OneDrive - Lawn Capital\LawnCapital\LC\Reports\Enhancements\Monthly Report\2026\`

**Naming Format:** `[Month#]-OneOffReport-[MonthName].xlsm`

| Month | Filename |
|-------|----------|
| January | `1-OneOffReport-January.xlsm` |
| February | `2-OneOffReport-February.xlsm` |
| March | `3-OneOffReport-March.xlsm` |
| April | `4-OneOffReport-April.xlsm` |
| May | `5-OneOffReport-May.xlsm` |
| June | `6-OneOffReport-June.xlsm` |
| July | `7-OneOffReport-July.xlsm` |
| August | `8-OneOffReport-August.xlsm` |
| September | `9-OneOffReport-September.xlsm` |
| October | `10-OneOffReport-October.xlsm` |
| November | `11-OneOffReport-November.xlsm` |
| December | `12-OneOffReport-December.xlsm` |

---

## Report Structure

### Sheet 1: Executive Dashboard
**KPI Cards (Top Row):**
- Total Enhancement Jobs (with breakdown: X standard / Y contracted)
- Total Revenue
- Gross Profit (with margin %)
- Total Costs (with expenses subtotal)
- Average Job Value
- Average Profit Margin

**Enhancement Type Comparison Table:**
| Job Type | Jobs | Revenue | Costs | Profit | Margin % | Avg Value |
|----------|------|---------|-------|--------|----------|-----------|
| Enhancement | formula | formula | formula | formula | formula | formula |
| Contracted Enhancement | formula | formula | formula | formula | formula | formula |
| TOTAL | formula | formula | formula | formula | formula | formula |

**Weekly Performance Table:**
- Week Starting (date)
- Jobs count
- Revenue
- Costs
- Profit (green if positive, red if negative)
- Margin %
- WoW Delta (week-over-week % change)

**P&L by Team Assignment:**
- Team name
- Job count
- Revenue, Costs, Profit
- Margin %

**Invoice Status Summary:**
- Paid (green)
- Awaiting Payment (yellow)
- Past Due (red)
- Count, Amount, % of Total, Avg Invoice

---

### Sheet 2: Weekly Trends
Detailed week-over-week analysis with:
- Week Start date
- Enhancement job count (green background)
- Contracted Enhancement count (blue background)
- Revenue (with data bars)
- Costs
- Profit (color-coded)
- Margin %
- Revenue WoW % change
- Profit WoW % change

---

### Sheet 3: Jobs (Filtered Data)
Raw job data filtered to Enhancement/Contracted Enhancement only:
- All original columns from Jobber export
- Added "Status" column: `=IF(M2>0,"Profitable",IF(M2<0,"Loss","Break-even"))`
- Job Type column color-coded:
  - Enhancement: Light green (#E2EFDA)
  - Contracted Enhancement: Light blue (#DDEBF7)
- Profit column conditional formatting (green/red)
- Frozen header row

---

### Sheet 4: Invoices (Filtered Data)
Only invoices linked to Enhancement job numbers:
- All original columns from Jobber export
- Status column color-coded:
  - Paid: Green (#C6EFCE)
  - Awaiting Payment: Yellow (#FFEB9C)
  - Past Due: Red (#FFC7CE)
- Balance column highlights non-zero values
- Frozen header row

---

### Sheet 5: Client Analysis
Enhancement clients ranked by revenue:
- Client name
- Job count
- Total Revenue
- Total Costs
- Profit (with conditional formatting)
- Margin % (formula)
- Avg Job Value (formula)

---

## Color Palette (Professional Finance Style)

| Element | Color Code | Usage |
|---------|------------|-------|
| Navy | #1B365D | Headers, titles |
| Dark Blue | #2E5090 | Subheaders |
| Medium Blue | #4472C4 | Total rows |
| Light Blue | #D6DCE5 | KPI card backgrounds |
| Green Positive | #548235 | Positive profit text |
| Green Light | #C6EFCE | Positive values fill |
| Red Negative | #C00000 | Negative profit text |
| Red Light | #FFC7CE | Negative values fill |
| Yellow Light | #FFEB9C | Caution/attention |
| Enhancement | #E2EFDA | Enhancement job type |
| Contracted | #DDEBF7 | Contracted Enhancement |
| Alt Row | #F8F9FA | Alternating row shading |

---

## Formula Standards

All summary metrics use dynamic formulas:

```excel
# Job counts by type
=COUNTIF(Jobs!O:O,"Enhancement")
=COUNTIF(Jobs!O:O,"Contracted Enhancement")

# Revenue/Cost/Profit by type
=SUMIF(Jobs!O:O,"Enhancement",Jobs!L:L)
=SUMIF(Jobs!O:O,"Enhancement",Jobs!K:K)
=SUMIF(Jobs!O:O,"Enhancement",Jobs!M:M)

# Margin calculation (with error handling)
=IFERROR(Profit/Revenue,0)

# Week-over-week change
=IFERROR((CurrentWeek-PriorWeek)/ABS(PriorWeek),0)

# Invoice status counts
=COUNTIF(Invoices!K:K,"Paid")
=COUNTIF(Invoices!K:K,"Awaiting Payment")
=COUNTIF(Invoices!K:K,"Past Due")
```

---

## How to Generate a Report

### Quick Command
Say: **"Generate the enhancement report for [Month]"**

### Manual Steps
1. Export Jobs Report from Jobber (One-off Jobs, date range = target month)
2. Export Invoice Report from Jobber (Invoices, date range = target month)
3. Save both CSVs to `C:\Users\daria\Downloads\`
4. Run the enhancement report generator script
5. Report saves to the 2026 folder with proper naming

---

## Verification Checklist

Before delivering the report, verify:

- [ ] Only Enhancement and Contracted Enhancement jobs included
- [ ] Job count matches filtered data
- [ ] Revenue/Costs/Profit totals are accurate
- [ ] All formulas calculate correctly (no #DIV/0! or #REF! errors)
- [ ] Weekly breakdown sums to total
- [ ] Invoice data only includes enhancement-related invoices
- [ ] Color coding is applied correctly
- [ ] File saved with correct naming convention

---

## Automated Weekly Generation

The report is automatically generated every **Friday at 6:00 PM EST** via GitHub Actions.

### How It Works
1. GitHub Actions triggers on schedule (cron: `0 23 * * 5`)
2. Python script authenticates with Jobber API using OAuth tokens
3. Fetches all Enhancement and Contracted Enhancement jobs for current month
4. Generates Excel report with full P&L analysis
5. Commits and pushes updated report to repository

### Manual Trigger
You can also manually trigger the workflow:
1. Go to: https://github.com/Grey-Development/enhancement-reports/actions
2. Click "Weekly Enhancement Report"
3. Click "Run workflow"
4. Optionally specify a month number (1-12)

### GitHub Secrets Required
| Secret | Description |
|--------|-------------|
| `JOBBER_ACCESS_TOKEN` | OAuth access token |
| `JOBBER_REFRESH_TOKEN` | OAuth refresh token |
| `JOBBER_CLIENT_ID` | OAuth client ID |
| `JOBBER_CLIENT_SECRET` | OAuth client secret |

---

## Trigger Phrases

Use these to invoke this report manually:

- "Generate enhancement report"
- "Monthly enhancement analysis"
- "One-off jobs report for [month]"
- "Enhancement P&L report"
- "Create the monthly enhancement report"

---

## Notes

- Reports are enhancement-focused only; excludes Open Issue, Service, and other job types
- Invoice filtering uses job number matching from the Invoice #s field
- Week-over-week calculations show growth/decline trends
- Team P&L helps identify which crews are most profitable
- Client analysis identifies top enhancement customers

---

## File Locations

| Item | Path |
|------|------|
| This CLAUDE.md | `C:\Users\daria\OneDrive - Lawn Capital\LawnCapital\LC\Reports\Enhancements\Monthly Report\2026\CLAUDE.md` |
| Monthly Reports | `C:\Users\daria\OneDrive - Lawn Capital\LawnCapital\LC\Reports\Enhancements\Monthly Report\2026\` |
| Input Data | `C:\Users\daria\Downloads\` |
| Generator Script | Generated dynamically by Claude Code |
