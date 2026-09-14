// Title: Show row grand totals in an Aspose.Cells pivot table using C# (ShowRowGrandTotals property)
// AI Prompts: Create a new workbook, populate sample sales data, add a pivot table, and enable row grand totals by setting PivotTable.ShowRowGrandTotals = true in C#. | Update an existing Aspose.Cells pivot table to display row totals at the bottom of the report by configuring the ShowRowGrandTotals property. | Programmatically generate an Excel file with a pivot table that shows row grand totals and save it as .xlsx using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to turn on row grand totals for a pivot table | set ShowRowGrandTotals to true in Aspose.Cells pivot table example | enable row totals at the bottom of pivot report using Aspose.Cells .NET | C# code to display row grand totals in Excel pivot table with Aspose
// Tags: Aspose.Cells pivot table ShowRowGrandTotals | C# enable row grand totals Excel pivot | Aspose.Cells generate pivot table with row totals | programmatic Excel pivot row totals Aspose | C# workbook save .xlsx with pivot grand totals

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // The sample creates a workbook, fills cells A1:C5 with product, region, and sales data, adds a pivot table at E3, assigns Product, Region, and Sales fields, enables row grand totals by setting ShowRowGrandTotals to true, and saves the file as ShowGrandTotalsForRowsDemo.xlsx.
    class ShowGrandTotalsForRows
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Laptop";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Laptop";
            cells["B3"].Value = "South";
            cells["C3"].Value = 1500;

            cells["A4"].Value = "Phone";
            cells["B4"].Value = "North";
            cells["C4"].Value = 800;

            cells["A5"].Value = "Phone";
            cells["B5"].Value = "South";
            cells["C5"].Value = 1100;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");   // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region"); // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // Sales as data field

            // Enable grand totals for rows (display row totals at the bottom)
            pivotTable.ShowRowGrandTotals = true;

            // Save the workbook to a file
            workbook.Save("ShowGrandTotalsForRowsDemo.xlsx");
        }
    }
}
