# Sales Performance Dashboard — Excel + VBA

An interactive Excel dashboard designed to monitor the daily sales performance of sales executives across regions.

## Project Overview

The dashboard analyzes **5 days of sales data** for sales executives and compares their performance against assigned targets.

### Questions / Problems Solved

- How are sales executives performing against their targets?
- Which executives are hitting or missing their targets?
- How does performance change when a region is selected?
- How can multiple PivotTables respond to the same Region slicer?
- How can the dashboard clearly show which region is currently selected?

## Solution

The dashboard combines **PivotTables, PivotCharts, slicers, calculated fields, and VBA automation** to create an interactive reporting view.

It includes:
- Total Sales
- Target Hit %
- Away From Target %
- Region-wise filtering
- Visual comparison of executive performance

### VBA Automation

Three macros were used:

- **`SlicerConnection`** — dynamically connects or disconnects PivotTables from the shared Region slicer based on checkbox selections.
- **`UpdateDashboardTitle`** — reads the current slicer selection and updates the dashboard label to show the active region(s).
- **`MasterUpdate`** — runs the required VBA updates together.

## Dashboard Components

| Component | Purpose |
|---|---|
| Raw Data | Source sales and target information |
| PivotTables | Summarize executive and regional performance |
| PivotCharts | Visualize sales and target performance |
| Region Slicer | Interactive regional filtering |
| VBA | Automate slicer connections and dashboard labeling |

## Tools

**Microsoft Excel · PivotTables · PivotCharts · Slicers · VBA**

## File

- `sales executive dashboard.xlsm` — macro-enabled Excel dashboard

**Note:** Macros must be enabled in Excel for the VBA-based filtering and dashboard-title functionality to work.

## Learning Outcome

This project strengthened my understanding of PivotTable architecture, slicer connections, dashboard automation, and VBA-driven Excel reporting.
