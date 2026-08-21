// Title: Aspose.Cells C# – Show Row Grand Totals & Hide Column Grand Totals in a Pivot Table
// Description: Creates a workbook with sample sales data, builds a pivot table on a separate sheet, assigns Product to rows, Region to columns, Sales to values, then enables row grand totals and disables column grand totals before refreshing and saving the file.
// Keywords: Aspose.Cells pivot table row grand total | Aspose.Cells hide column grand total | C# Aspose.Cells pivot totals visibility | ShowRowGrandTotals Aspose.Cells | ShowColumnGrandTotals false | Aspose.Cells .NET pivot table settings
// Common Searches: Aspose.Cells C# show only row grand totals in pivot table | how to hide column grand totals using Aspose.Cells | set ShowRowGrandTotals true Aspose.Cells | disable column grand totals Aspose.Cells pivot | pivot table grand total options Aspose.Cells .NET
// Developer Intent: Configure a pivot table so that row grand totals are displayed while column grand totals are suppressed.
// Use Cases: Generate a product‑wise sales summary where column totals would clutter the view. | Create a financial report that emphasizes total per row category without column aggregates. | Export an Excel workbook for presentation that needs only row grand totals for clarity.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot table that shows row grand totals but hides column grand totals. | Demonstrate how to set ShowRowGrandTotals and ShowColumnGrandTotals properties on an Aspose.Cells pivot table. | Explain how to modify an existing Aspose.Cells pivot table to change grand total visibility without rebuilding the table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // Creates a workbook with sample sales data, builds a pivot table on a separate sheet, assigns Product to rows, Region to columns, Sales to values, then enables row grand totals and disables column grand totals before refreshing and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (data sheet)
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].Value = "Product";
            dataSheet.Cells["B1"].Value = "Region";
            dataSheet.Cells["C1"].Value = "Sales";

            dataSheet.Cells["A2"].Value = "Product A";
            dataSheet.Cells["B2"].Value = "North";
            dataSheet.Cells["C2"].Value = 1000;

            dataSheet.Cells["A3"].Value = "Product B";
            dataSheet.Cells["B3"].Value = "South";
            dataSheet.Cells["C3"].Value = 1500;

            dataSheet.Cells["A4"].Value = "Product A";
            dataSheet.Cells["B4"].Value = "South";
            dataSheet.Cells["C4"].Value = 2000;

            dataSheet.Cells["A5"].Value = "Product B";
            dataSheet.Cells["B5"].Value = "North";
            dataSheet.Cells["C5"].Value = 1200;

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table (source range A1:C5, destination top‑left cell E3)
            int pivotIndex = pivotSheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Sales as data field

            // Set grand total visibility:
            // Show row grand totals
            pivotTable.ShowRowGrandTotals = true;
            // Hide column grand totals
            pivotTable.ShowColumnGrandTotals = false;

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_RowGrand_HideColumnGrand.xlsx");
        }
    }
}
