// Title: Set a custom data field caption (e.g., "Sales Amount") for an Aspose.Cells PivotTable in C#
// AI Prompts: Generate C# code that creates a PivotTable with Aspose.Cells and assigns a custom caption such as "Sales Amount" to the data field using the appropriate property. | Modify an existing Aspose.Cells PivotTable in C# to rename its data column header programmatically.
// Common Searches: Aspose.Cells C# how to rename pivot table data field header | set DataFieldHeaderName property for pivot table using Aspose.Cells | change pivot table data caption programmatically in .NET | example of customizing pivot table column name with Aspose.Cells | C# Aspose.Cells pivot table custom data caption tutorial
// Tags: Aspose.Cells pivot table data caption | C# rename pivot table data field | custom data header for Excel pivot via Aspose | programmatic pivot table column naming .NET | Aspose.Cells set pivot data column title

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // The example creates a new workbook, fills it with sample product, region, and sales data, adds a PivotTable on range A1:C5, assigns Product as rows, Region as columns, and Sales as the data field, then sets the DataFieldHeaderName to "Sales Amount" to provide a clearer column heading, refreshes and calculates the pivot, and finally saves the file as PivotTableWithDataCaption.xlsx.
    class SetDataCaption
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Bike";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Bike";
            cells["B3"].Value = "South";
            cells["C3"].Value = 1500;

            cells["A4"].Value = "Car";
            cells["B4"].Value = "North";
            cells["C4"].Value = 2000;

            cells["A5"].Value = "Car";
            cells["B5"].Value = "South";
            cells["C5"].Value = 2500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Sales as data field

            // Set the data field header caption to "Sales Amount"
            pivotTable.DataFieldHeaderName = "Sales Amount";

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableWithDataCaption.xlsx");
        }
    }
}
