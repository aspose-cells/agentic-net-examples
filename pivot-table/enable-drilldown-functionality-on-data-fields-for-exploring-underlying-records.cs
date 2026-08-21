// Title: Enable Drill‑Down in an Aspose.Cells Pivot Table (C#/.NET)
// Description: This example creates a workbook, populates sample sales data, adds a pivot table on range A1:C5, assigns Category, SubCategory and Sales to row, column and data areas, and activates drill‑down by setting EnableDrilldown, ShowDrill and PrintDrill to true. The pivot is refreshed, calculated, and saved as PivotTableDrilldownDemo.xlsx, providing interactive expand/collapse functionality and printable drill indicators.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | drill down | EnableDrilldown | ShowDrill | PrintDrill | expand collapse buttons | interactive Excel report | Excel pivot drill‑through
// Common Searches: Aspose.Cells enable drill down pivot table C# | How to show expand collapse buttons in Aspose.Cells pivot | Print drill indicators with Aspose.Cells .NET | C# code for drill‑through in Excel pivot using Aspose | Aspose.Cells pivot table interactive dashboard example
// Developer Intent: Add interactive drill‑down capability to a pivot table so end‑users can view the underlying rows behind aggregated values.
// Use Cases: Allow analysts to double‑click a total and see the detailed sales records that compose it. | Generate printable reports that retain visual drill icons for offline review. | Build Excel‑based dashboards where rows and columns can be expanded or collapsed on demand. | Create reusable workbook templates with pre‑configured drill‑down settings for multiple datasets.
// AI Prompts: Write C# code that builds a pivot table with EnableDrilldown, ShowDrill, and PrintDrill set to true, and applies custom number formatting. | Explain how to retrieve the source rows displayed after a user drills down into a pivot cell using Aspose.Cells APIs. | Show how to toggle ShowDrill and PrintDrill at runtime based on a configuration flag in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsDrilldownDemo
{
    // This example creates a workbook, populates sample sales data, adds a pivot table on range A1:C5, assigns Category, SubCategory and Sales to row, column and data areas, and activates drill‑down by setting EnableDrilldown, ShowDrill and PrintDrill to true. The pivot is refreshed, calculated, and saved as PivotTableDrilldownDemo.xlsx, providing interactive expand/collapse functionality and printable drill indicators.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "SubCategory";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = "Phones";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = "Laptops";
            sheet.Cells["C3"].Value = 2500;

            sheet.Cells["A4"].Value = "Furniture";
            sheet.Cells["B4"].Value = "Chairs";
            sheet.Cells["C4"].Value = 800;

            sheet.Cells["A5"].Value = "Furniture";
            sheet.Cells["B5"].Value = "Tables";
            sheet.Cells["C5"].Value = 1500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");      // Row field
            pivot.AddFieldToArea(PivotFieldType.Column, "SubCategory"); // Column field
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");       // Data field

            // Enable drill‑down functionality and show the expand/collapse buttons
            pivot.EnableDrilldown = true;   // Allows users to double‑click a cell to see underlying records
            pivot.ShowDrill = true;        // Displays the drill indicators in the UI
            pivot.PrintDrill = true;       // Ensures the indicators are printed if needed

            // Refresh and calculate the pivot table data
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableDrilldownDemo.xlsx");
        }
    }
}
