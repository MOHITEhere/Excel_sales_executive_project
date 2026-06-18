Sales Executive Dashboard (Excel Project)

 Overview

This project is an interactive Sales Executive Dashboard built in Microsoft Excel** using Pivot Tables, Slicers, Charts, and VBA automation.

It helps track:

Total Sales performance
 Target Achievement (%)
 Target Gap (Away %)
 Region-wise filtering using slicers

The dashboard is designed to give quick insights into sales performance** and allows dynamic interaction through checkboxes and slicers.


Dashboard Preview
 Features

Interactive Slicer (Region Filter)
  Filter data by cities like Chennai, Delhi, Mumbai, etc.

  Multiple Dashboards

   Dashboard 1 → Total Sales
   Dashboard 2→ Target Hit %
   Dashboard 3 → Target Away %

  Visualizations

   Bar Chart (Sales Comparison)
   Pie Chart (Target Distribution)
   Line Chart (Target Gap Trend)

 Dynamic Dashboard Control
  Enable/Disable dashboards using checkboxes.

 VBA Automation
  Automatically connects/disconnects Pivot Tables from slicers.



 Tech Stack

 Microsoft Excel
 Pivot Tables
 Slicers
 Charts
 VBA (Visual Basic for Applications)

VBA Logic Used

The project uses VBA to dynamically control which Pivot Tables respond to slicers.

 Core Macro

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

 What This Does

 Checks if a checkbox (linked to a cell) is TRUE/FALSE
 If TRUE → connects Pivot Table to slicer
 If FALSE → disconnects Pivot Table
 Makes dashboard fully dynamic

---

Project Structure

```
 Sales-Dashboard
 ┣  README.md
 ┣  Sales_Dashboard.xlsx
 ┗  image(1).png
```


 Use Case  
Sales Performance Tracking
Business Reporting
Interview Portfolio Project
Excel + VBA Practice



Key Learnings

 Advanced Excel Dashboard Design
 Slicer & Pivot Table Integration
 VBA Automation for dynamic control
 Data visualization best practices

Future Improvements

 Add Power BI version
 Automate data refresh
 Add monthly/yearly trends
 Improve UI/UX design
