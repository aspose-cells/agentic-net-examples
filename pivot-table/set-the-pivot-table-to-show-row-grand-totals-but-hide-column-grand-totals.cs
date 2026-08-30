// Title: Enable row grand totals and hide column grand totals in an Aspose.Cells pivot table using C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a pivot table, sets ShowRowGrandTotals = true and ShowColumnGrandTotals = false, then saves the workbook. | Write a .NET example that adds sample data, defines row, column, and data fields for a pivot table, and configures it to display only row grand totals.
// Common Searches: how to show only row grand totals in Aspose.Cells pivot table C# | Aspose.Cells hide column grand totals example | C# pivot table ShowRowGrandTotals true ShowColumnGrandTotals false | Aspose.Cells set pivot table grand total visibility programmatically | C# create pivot table with row totals but no column totals using Aspose.Cells
// Tags: Aspose.Cells pivot table grand total visibility | C# ShowRowGrandTotals property | Aspose.Cells column grand totals off | pivot table row totals only .NET | Aspose.Cells create pivot table programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // The program creates a workbook, inserts sample data, adds a pivot table on a new sheet, assigns Product as rows, Region as columns, Sales as data, enables row grand totals, disables column grand totals, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (data sheet)
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

            // Show row grand totals and hide column grand totals
            pivotTable.ShowRowGrandTotals = true;   // Enable row grand totals
            pivotTable.ShowColumnGrandTotals = false; // Disable column grand totals

            // Save the workbook
            workbook.Save("PivotTable_RowGrand_HideColumnGrand.xlsx");
        }
    }
}
