// Title: Create a new workbook and add a PivotTable with Aspose.Cells for .NET
// Description: Demonstrates how to instantiate a Workbook, fill cells A1:C6 with product, region, and sales data, define the source range, add a PivotTable named "SalesPivot" at cell E5 using the string overload, assign Product to rows, Region to columns, Sales to data, refresh and calculate the pivot, and save the file as PivotTableDemo.xlsx.
// Keywords: Aspose.Cells PivotTable example | C# add PivotTable string overload | create workbook Aspose.Cells .NET | populate cells programmatically | refresh pivot data Aspose | calculate pivot Aspose.Cells | save Excel file C# | dynamic PivotTable generation
// Common Searches: How to add a PivotTable to a new workbook using Aspose.Cells for .NET | Aspose.Cells string overload for PivotTable source range | Set row, column, and data fields in Aspose.Cells PivotTable | Refresh and calculate PivotTable programmatically Aspose | Save Excel file after creating PivotTable with Aspose.Cells
// Developer Intent: Programmatically build a PivotTable from in‑memory data in a freshly created workbook and export it as an Excel file.
// Use Cases: Generate a sales summary report without external data files. | Automate region‑wise product analysis for recurring business dashboards. | Create, refresh, and calculate PivotTables on the fly before distributing Excel workbooks.
// AI Prompts: Show how to add multiple data fields to the PivotTable using Aspose.Cells. | Provide code to format PivotTable headers and apply number formatting to the Sales field. | Explain how to change the source data range of an existing PivotTable and refresh it in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Demonstrates how to instantiate a Workbook, fill cells A1:C6 with product, region, and sales data, define the source range, add a PivotTable named "SalesPivot" at cell E5 using the string overload, assign Product to rows, Region to columns, Sales to data, refresh and calculate the pivot, and save the file as PivotTableDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample source data for the pivot table
            Cells cells = worksheet.Cells;
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Product1";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1000;

            cells["A3"].Value = "Product2";
            cells["B3"].Value = "South";
            cells["C3"].Value = 2000;

            cells["A4"].Value = "Product3";
            cells["B4"].Value = "East";
            cells["C4"].Value = 3000;

            cells["A5"].Value = "Product1";
            cells["B5"].Value = "West";
            cells["C5"].Value = 4000;

            cells["A6"].Value = "Product2";
            cells["B6"].Value = "North";
            cells["C6"].Value = 5000;

            // Define the source range, destination cell, and pivot table name
            string sourceData = "A1:C6";
            string destCell = "E5";
            string tableName = "SalesPivot";

            // Add the pivot table using the (string, string, string) overload
            int pivotIndex = worksheet.PivotTables.Add(sourceData, destCell, tableName);

            // Retrieve the newly created pivot table
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableDemo.xlsx");
        }
    }
}
