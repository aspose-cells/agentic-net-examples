// Title: Sum Units Sold with a Pivot Table in Aspose.Cells for .NET (C#)
// Description: The example builds a workbook, adds product and UnitsSold rows, creates a pivot table on range A1:B6, assigns Product as a row field, sets UnitsSold as a Sum data field, refreshes and calculates the pivot, and saves the file as TotalUnitsSoldPivot.xlsx.
// Keywords: Aspose.Cells | C# pivot table sum | ConsolidationFunction.Sum | total units sold | Excel aggregation .NET | pivot table example | sales summary pivot | Aspose.Cells API | generate Excel pivot | calculate totals C#
// Common Searches: Aspose.Cells set sum aggregation for a pivot data field C# | how to total units sold with a pivot table using Aspose.Cells | C# example creating a sales pivot table in Aspose.Cells | pivot table sum function Aspose.Cells .NET | calculate product totals with Aspose.Cells pivot
// Developer Intent: Create an Excel pivot table that aggregates the UnitsSold column using the Sum function to display total units per product.
// Use Cases: Produce a concise sales summary that lists each product with its total units sold. | Export aggregated sales numbers to an Excel workbook for distribution to business stakeholders. | Automate daily reporting by generating and refreshing a pivot table in a .NET service.
// AI Prompts: Show how to add an Average aggregation for UnitsSold alongside the Sum in the same pivot table. | Provide code to format the summed UnitsSold values as currency and highlight the total row with bold styling. | Explain how to include a column field for product categories and programmatically refresh the pivot after data changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSumExample
{
    // The example builds a workbook, adds product and UnitsSold rows, creates a pivot table on range A1:B6, assigns Product as a row field, sets UnitsSold as a Sum data field, refreshes and calculates the pivot, and saves the file as TotalUnitsSoldPivot.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Product | UnitsSold
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("UnitsSold");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(85);
            cells["A4"].PutValue("Apple");
            cells["B4"].PutValue(95);
            cells["A5"].PutValue("Banana");
            cells["B5"].PutValue(110);
            cells["A6"].PutValue("Cherry");
            cells["B6"].PutValue(60);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add Product as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add UnitsSold as a data field
            int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "UnitsSold");
            PivotField unitsSoldField = pivotTable.DataFields[dataFieldPos];

            // Ensure the aggregation function is Sum (default, but set explicitly)
            unitsSoldField.Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("TotalUnitsSoldPivot.xlsx");
        }
    }
}
