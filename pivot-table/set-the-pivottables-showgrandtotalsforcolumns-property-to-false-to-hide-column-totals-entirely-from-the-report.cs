// Title: C# – Hide Column Grand Totals in an Aspose.Cells PivotTable (ShowColumnGrandTotals = false)
// Description: Creates a workbook with sample sales data, adds a PivotTable (Category rows, Region columns, Sales values), disables column grand totals by setting ShowColumnGrandTotals to false, refreshes the pivot data, and saves the file as an Excel report.
// Keywords: Aspose.Cells C# | .NET PivotTable | ShowColumnGrandTotals | hide column totals | pivot grand total settings | Excel report generation | sales summary pivot | Aspose.Cells API | disable column grand totals
// Common Searches: Aspose.Cells hide column grand totals C# | ShowColumnGrandTotals false example | remove column totals from PivotTable using Aspose | C# pivot table grand total visibility | Aspose.Cells pivot settings tutorial
// Developer Intent: Turn off column grand total calculations in a generated PivotTable.
// Use Cases: Produce a sales dashboard where column aggregates are redundant. | Create financial worksheets that display only detailed column figures. | Design compact Excel reports that omit column grand totals for clearer layout.
// AI Prompts: Generate C# code with Aspose.Cells that builds a PivotTable and sets ShowColumnGrandTotals to false. | Explain how the ShowColumnGrandTotals property influences PivotTable output and when to call CalculateData. | Show a complete example that hides both row and column grand totals in an Aspose.Cells PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook with sample sales data, adds a PivotTable (Category rows, Region columns, Sales values), disables column grand totals by setting ShowColumnGrandTotals to false, refreshes the pivot data, and saves the file as an Excel report.
    class HideColumnGrandTotals
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Electronics";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Electronics";
            cells["B3"].Value = "South";
            cells["C3"].Value = 1500;

            cells["A4"].Value = "Clothing";
            cells["B4"].Value = "North";
            cells["C4"].Value = 800;

            cells["A5"].Value = "Clothing";
            cells["B5"].Value = "South";
            cells["C5"].Value = 950;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, columns = Region, data = Sales
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as rows
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as columns
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data

            // Hide column grand totals by setting ShowColumnGrandTotals to false
            pivotTable.ShowColumnGrandTotals = false;

            // Optional: recalculate the pivot table data after changing settings
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_HideColumnGrandTotals.xlsx");
        }
    }
}
