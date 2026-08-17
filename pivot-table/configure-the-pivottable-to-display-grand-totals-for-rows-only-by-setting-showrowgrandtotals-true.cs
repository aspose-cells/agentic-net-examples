// Title: Aspose.Cells C# PivotTable – Show Row Grand Totals Only
// Description: Creates a workbook, adds sample product‑region‑sales data, builds a PivotTable, assigns Product to rows, Region to columns, Sales to values, then enables row grand totals while hiding column grand totals, and saves the file.
// Keywords: Aspose.Cells | C# PivotTable | ShowRowGrandTotals | disable column grand totals | Excel automation .NET | pivot table grand totals | row totals only | Aspose.Cells example | pivot table settings
// Common Searches: Aspose.Cells show only row grand totals | C# pivot table hide column totals Aspose | Set ShowRowGrandTotals property in .NET | PivotTable grand total configuration Aspose.Cells | How to display row totals without column totals in Excel using Aspose
// Developer Intent: Configure an Aspose.Cells PivotTable in .NET to display grand totals for rows only.
// Use Cases: Generate a sales summary where only the overall row total is required, simplifying the report layout. | Create dashboards that emphasize product‑level aggregates while omitting column‑wise totals. | Export Excel files with PivotTables that need row‑only grand totals for downstream data processing.
// AI Prompts: Write C# code with Aspose.Cells to create a PivotTable that shows only row grand totals. | Explain the impact of ShowRowGrandTotals and ShowColumnGrandTotals on PivotTable output in Aspose.Cells. | Provide a step‑by‑step tutorial for configuring row‑only grand totals in an Aspose.Cells PivotTable for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook, adds sample product‑region‑sales data, builds a PivotTable, assigns Product to rows, Region to columns, Sales to values, then enables row grand totals while hiding column grand totals, and saves the file.
    class ShowRowGrandTotalsOnly
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Product A";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1000;

            cells["A3"].Value = "Product B";
            cells["B3"].Value = "South";
            cells["C3"].Value = 1500;

            cells["A4"].Value = "Product A";
            cells["B4"].Value = "South";
            cells["C4"].Value = 2000;

            cells["A5"].Value = "Product B";
            cells["B5"].Value = "North";
            cells["C5"].Value = 1200;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Sales as data field

            // Configure grand totals: show only row grand totals
            pivotTable.ShowRowGrandTotals = true;    // Enable row grand totals
            pivotTable.ShowColumnGrandTotals = false; // Disable column grand totals

            // Save the workbook to a file
            workbook.Save("PivotTable_ShowRowGrandTotalsOnly.xlsx");
        }
    }
}
