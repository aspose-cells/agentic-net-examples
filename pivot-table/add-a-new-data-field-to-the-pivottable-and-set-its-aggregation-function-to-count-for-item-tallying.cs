// Title: Add a Count aggregation field to a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, populates it with sample data, adds a PivotTable, and configures a data field to use the Count aggregation via Aspose.Cells. | Modify an existing Aspose.Cells PivotTable in C# by inserting a new data field and setting its ConsolidationFunction to Count for item tallying.
// Common Searches: Aspose.Cells C# how to set Count function for a pivot table data field | programmatically add a count data field to a pivot table with Aspose.Cells .NET | C# example of using ConsolidationFunction.Count in Aspose.Cells PivotTable | create pivot table with product row and count of items using Aspose.Cells | Aspose.Cells pivot table aggregation options count example
// Tags: Aspose.Cells add count data field pivot table | C# set ConsolidationFunction.Count Aspose.Cells | pivot table data field aggregation Aspose.Cells | Aspose.Cells create pivot table with row field | Aspose.Cells count aggregation example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCountExample
{
    // The sample creates a new workbook, fills it with product, region, and sales data, adds a PivotTable at E3, assigns 'Product' as a row field, adds 'Product' again as a data field with the Count aggregation, refreshes and calculates the pivot, and saves the file as PivotTable_CountField.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Columns: Product, Region, Sales
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(150);

            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B4"].PutValue("East");
            sheet.Cells["C4"].PutValue(200);

            sheet.Cells["A5"].PutValue("Banana");
            sheet.Cells["B5"].PutValue("West");
            sheet.Cells["C5"].PutValue(180);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add "Product" as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add "Product" as a data field to count the number of items per product
            int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Product");
            PivotField dataField = pivotTable.DataFields[dataFieldPos];
            dataField.Function = ConsolidationFunction.Count; // Set aggregation to Count

            // Refresh and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CountField.xlsx");
        }
    }
}
