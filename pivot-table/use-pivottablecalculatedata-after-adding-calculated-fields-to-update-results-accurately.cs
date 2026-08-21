// Title: Use PivotTable.CalculateData to Refresh a Calculated Field in Aspose.Cells for .NET (C#)
// Description: Creates a workbook with product, quantity, and price data, adds a pivot table, inserts a calculated field (TotalSales = Quantity × Price), refreshes the pivot cache, calls PivotTable.CalculateData to populate the new values, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells PivotTable CalculateData | C# calculated field pivot table | RefreshData vs CalculateData Aspose | Add calculated field Aspose.Cells | Update pivot results .NET
// Common Searches: Aspose.Cells PivotTable.CalculateData example | how to recalculate pivot after adding calculated field | C# Aspose.Cells refresh pivot cache | add total sales calculated field to pivot | Aspose.Cells pivot table refresh data
// Developer Intent: Refresh the pivot cache and recompute the pivot so that a newly added calculated field displays correct results.
// Use Cases: Generate a sales summary where TotalSales = Quantity × Price and the pivot reflects the calculation. | Build a dynamic reporting workbook that updates automatically after source data changes and custom calculations are added. | Export an Excel file with a pivot table containing custom formulas for downstream analysis in Excel or Power BI.
// AI Prompts: Show how to add multiple calculated fields to a PivotTable and recalculate them with Aspose.Cells for .NET. | Explain the difference between RefreshData and CalculateData when working with Aspose.Cells pivot tables. | Provide code to update an existing pivot table after changing source data and adding a new calculated field, ensuring results are refreshed.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook with product, quantity, and price data, adds a pivot table, inserts a calculated field (TotalSales = Quantity × Price), refreshes the pivot cache, calls PivotTable.CalculateData to populate the new values, and saves the file as an Excel workbook.
    public class PivotTableCalculateDataWithCalculatedFieldDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(2);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(5);
                sheet.Cells["C3"].PutValue(3);

                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B4"].PutValue(8);
                sheet.Cells["C4"].PutValue(1.5);

                // Add a pivot table based on the data range A1:C4, place it at E3
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Product, data = Quantity and Price
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Price");

                // Add a calculated field that computes total sales = Quantity * Price
                // The third parameter 'true' drags the field to the data area automatically
                pivotTable.AddCalculatedField("TotalSales", "=Quantity*Price", true);

                // Refresh the pivot cache from the source data (correct API)
                pivotTable.RefreshData();

                // Calculate the pivot data so that the calculated field values are populated
                pivotTable.CalculateData();

                // Save the workbook to a file
                workbook.Save("PivotTable_With_CalculatedField.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
            }
        }
    }
}
