
## ⚠️ Note# Sales Performance Dashboard (Excel + VBA)

An Excel dashboard built with pivot tables, pivot charts, slicers, and VBA macros to track and monitor daily sales performance of sales executives across regions.

## 📊 What it does

- Tracks **5 days of sales data** (Day 1–Day 5) per sales executive
- Groups performance **region-wise** (Mumbai, Delhi, Chennai, Pune, and more)
- Calculates **Total Sales**, **Target Hit %**, and **Away From Target %** for each executive
- Lets managers filter and compare performance across regions using an interactive slicer
- Displays a live **"Showing: [Region]"** label so it's always clear what the dashboard is currently filtered to
- Gives a quick, visual way to spot who's on track to hit target and who needs support — without digging through raw rows of data

## 🧩 How it's built

| Component | Purpose |
|---|---|
| **Raw Data sheet** | Source data: Emp Code, Sales Executive, Region, Day1–Day5, Total Sales, Target, Target Hit %, Away From Target % |
| **Dashboard sheet** | 4 pivot tables + pivot charts summarizing performance by region and executive |
| **Slicer (Region)** | Single interactive filter shared across the dashboard views |
| **VBA Macros** | Automate slicer connections and keep the dashboard title in sync with the current filter |

### VBA logic
- `SlicerConnection` — dynamically connects or disconnects each pivot table from the shared Region slicer based on checkbox toggles, so users can control exactly which views respond to filtering.
- `UpdateDashboardTitle` — reads the currently selected slicer item(s) and updates a "Showing: ..." label at the top of the dashboard (defaults to "All Regions" when nothing is selected).
- `MasterUpdate` — runs both macros together to refresh the dashboard in one click.

## 🎯 Why I built this

This was my first Excel dashboard combining pivot tables with VBA automation. The goal was to give managers a single, self-updating view to monitor sales executive performance against targets — instead of manually reviewing raw sales sheets every day.

## 🛠️ Tech used

- Microsoft Excel (Pivot Tables, Pivot Charts, Slicers)
- VBA (Visual Basic for Applications)

## 📂 File

- `dashboard-1-styled.xlsm` — the dashboard workbook (macro-enabled)


Macros must be **enabled** when opening the file for the dashboard's filtering and auto-title features to work (Excel will prompt you to "Enable Content" on open).

---

*Feedback and suggestions welcome — this is an ongoing learning project.*
