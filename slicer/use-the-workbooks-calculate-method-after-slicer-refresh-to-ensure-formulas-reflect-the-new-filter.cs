// Title: Recalculate formulas after slicer refresh with Aspose.Cells for .NET
// Description: Shows how to modify source data, refresh a slicer linked to a pivot table, and call Workbook.CalculateFormula so GETPIVOTDATA formulas reflect the new filter before saving the workbook.
// Keywords: Aspose.Cells slicer refresh | Workbook.CalculateFormula | GETPIVOTDATA update | pivot table recalculation C# | Aspose.Cells example
// Common Searches: Aspose.Cells recalculate formulas after slicer refresh | How to use Workbook.CalculateFormula with slicer.Refresh | Update GETPIVOTDATA after changing slicer filter in .NET | Refresh pivot table formulas Aspose.Cells C# | Slicer linked to pivot table Aspose.Cells example
// Developer Intent: Refresh a slicer and recalculate all workbook formulas so dependent GETPIVOTDATA cells show the latest results.
// Use Cases: Automated reporting that adjusts totals when a slicer filter changes. | Dynamic dashboards where source data is edited, the slicer is refreshed, and formulas are updated programmatically. | Batch processing of workbooks that require consistent pivot‑table calculations after slicer operations.
// AI Prompts: Provide C# code using Aspose.Cells to refresh a slicer, then recalculate the workbook so GETPIVOTDATA values are updated. | Generate an example that changes source data, calls slicer.Refresh, and invokes Workbook.CalculateFormula to keep formulas accurate. | Explain why Workbook.CalculateFormula is necessary after slicer.Refresh when using GETPIVOTDATA in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Shows how to modify source data, refresh a slicer linked to a pivot table, and call Workbook.CalculateFormula so GETPIVOTDATA formulas reflect the new filter before saving the workbook.
class SlicerRefreshCalculateDemo
{
    static void Main()
    {
        // Create a new workbook and add a data worksheet
        Workbook wb = new Workbook();
        Worksheet dataSheet = wb.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["B3"].PutValue(200);
        dataSheet.Cells["A4"].PutValue("Orange");
        dataSheet.Cells["B4"].PutValue(150);

        // Create a pivot table based on the data range
        Worksheet pivotSheet = wb.Worksheets.Add("Pivot");
        int pivotIdx = pivotSheet.PivotTables.Add("Data!A1:B4", "C3", "SalesPivot");
        PivotTable pivot = pivotSheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product field
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales field

        // Add a formula that retrieves the total sales from the pivot table
        Worksheet reportSheet = wb.Worksheets.Add("Report");
        // GETPIVOTDATA will pull the aggregated Sales value from the pivot table
        reportSheet.Cells["A1"].Formula = "=GETPIVOTDATA(\"Sales\",\"Pivot!$C$3\")";

        // Add a slicer linked to the pivot table for the Product field
        Worksheet slicerSheet = wb.Worksheets.Add("Slicer");
        int slicerIdx = slicerSheet.Slicers.Add(pivot, "A1", "Product");
        Slicer slicer = slicerSheet.Slicers[slicerIdx];

        // Modify source data to demonstrate slicer filtering effect
        dataSheet.Cells["A4"].PutValue("Apple"); // Change "Orange" to "Apple"

        // Refresh the slicer (this also refreshes the associated pivot table)
        slicer.Refresh();

        // After slicer refresh, recalculate all formulas so that the GETPIVOTDATA result is up‑to‑date
        wb.CalculateFormula();

        // Save the workbook (using the standard save method)
        wb.Save("SlicerRefreshCalculateDemo.xlsx");
    }
}
