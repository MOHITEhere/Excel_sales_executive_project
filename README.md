# Sales Executive Dashboard — Excel + VBA Project

An interactive Excel dashboard built to track the daily sales performance of 141 sales executives across 8 regions in India. The dashboard displays Total Sales, Target Hit %, and Target Away % with dynamic region-based filtering using slicers and VBA-powered checkbox control.

---

## Dashboard Preview

![Sales Executive Dashboard](dashboard_sales.png)

---

## Problem Statement

A sales organization with 141 field executives spread across Mumbai, Delhi, Nagpur, Chennai, Pune, Patna, Ranchi, and Surat needed a single-view performance tracker. Managers required the ability to switch between different KPI views and filter by region without navigating multiple sheets.

This dashboard solves that with three switchable views, a shared slicer, and VBA automation to keep each view independent.

---

## Dataset

| Property | Value |
|---|---|
| Executives Tracked | 141 |
| Regions | Mumbai, Delhi, Nagpur, Chennai, Pune, Patna, Ranchi, Surat |
| Tracking Period | 5 Days (Day 1 to Day 5) |
| Target per Executive | 500 units |
| Data Source | RAW DATA sheet (Excel) |

### Columns

| Column | Description |
|---|---|
| Emp Code | Unique employee ID (e.g. Mum-TCL001) |
| Sales Executive | Executive name |
| Region | City of operation |
| Day 1 – Day 5 | Daily sales figures |
| Total Sales | Sum of Day 1 to Day 5 |
| Target | Fixed target of 500 units |
| Target Hit % | Total Sales / Target |
| Away From Target % | 1 - Target Hit % |

---

## Dashboard Views

All three views share the same region slicer. Checkboxes on the dashboard enable or disable which Pivot Table responds to the slicer at any time.

### Dashboard 1 — Total Sales
Shows total units sold per executive. Used to identify top and bottom performers by raw output.

### Dashboard 2 — Target Hit %
Shows what percentage of the 500-unit target each executive has achieved. Helps managers see who is on track.

### Dashboard 3 — Target Away %
Shows how far each executive is from their target. Useful for identifying who needs intervention.

---

## VBA Automation

The core macro dynamically connects or disconnects each Pivot Table from the region slicer based on checkbox state. This ensures only the active dashboard view responds to filtering.

```vba
Sub Macro3()

    If Sheet3.Range("A1").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region1").PivotTables.AddPivotTable _
            ActiveSheet.PivotTables("PivotTable1")
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region1").PivotTables.RemovePivotTable _
            ActiveSheet.PivotTables("PivotTable1")
    End If

    If Sheet3.Range("D1").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region1").PivotTables.AddPivotTable _
            ActiveSheet.PivotTables("PivotTable3")
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region1").PivotTables.RemovePivotTable _
            ActiveSheet.PivotTables("PivotTable3")
    End If

    If Sheet3.Range("G1").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region1").PivotTables.AddPivotTable _
            ActiveSheet.PivotTables("PivotTable4")
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region1").PivotTables.RemovePivotTable _
            ActiveSheet.PivotTables("PivotTable4")
    End If

    If Sheet3.Range("J1").Value = True Then
        ActiveWorkbook.SlicerCaches("Slicer_Region1").PivotTables.AddPivotTable _
            ActiveSheet.PivotTables("PivotTable5")
    Else
        ActiveWorkbook.SlicerCaches("Slicer_Region1").PivotTables.RemovePivotTable _
            ActiveSheet.PivotTables("PivotTable5")
    End If

End Sub
```

Each checkbox is linked to a cell (A1, D1, G1, J1). When a checkbox is ticked, the cell turns TRUE and the macro connects that Pivot Table to the slicer. Unticking disconnects it — so filtering one view does not affect the others.

---

## Tech Stack

| Tool | Purpose |
|---|---|
| Microsoft Excel (.xlsm) | Dashboard, Pivot Tables, Charts, Slicers |
| VBA (Visual Basic for Applications) | Dynamic slicer-pivot table connection control |

---

## Project Structure

```
Excel_sales_executive_project/
├── sales_executive project.xlsm    # Main workbook with RAW DATA and DASH sheets
├── dashboard_sales.png             # Dashboard screenshot
└── README.md
```

---

## Key Learnings

- Pivot Table design and slicer integration across multiple dashboards in one sheet
- VBA macro to dynamically link and unlink Pivot Tables from a shared slicer
- Checkbox-to-cell linking for interactive dashboard control
- Performance KPI design: absolute sales, target hit %, and gap analysis

---

## Author

Atharva Mohite
B.Tech Student — DY Patil University (RAIT), Expected 2027
Specialisation: Machine Learning · Data Science · NLP
GATE 2026 Qualified — Data Science & AI (AIR 7007)
[GitHub](https://github.com/MOHITEhere) | [LinkedIn](https://linkedin.com/in/<your-profile>)
