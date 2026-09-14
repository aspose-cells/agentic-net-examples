// Title: Generate a Pivot Table that Sums Units Sold Using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to create a workbook, populate Product, Region, and UnitsSold columns, add a pivot table on range A1:C7, set Product as the row field, Region as the column field, and configure UnitsSold as a summed data field. | Show how to refresh the pivot cache, calculate the pivot data, and save the resulting workbook as an .xlsx file using Aspose.Cells.
// Common Searches: aspnet aspocells how to set sum aggregation for a pivot table data field | c# create pivot table from worksheet range and calculate total units sold | aspocells pivot table refresh data and calculate after adding fields | example of using ConsolidationFunction.Sum in Aspose.Cells pivot table | save pivot table workbook as xlsx with Aspose.Cells C#
// Tags: Aspose.Cells create pivot table C# | pivot table sum aggregation Aspose.Cells | UnitsSold data field ConsolidationFunction.Sum | refresh pivot cache Aspose.Cells | save workbook as .xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates building a workbook, inserting sample product/region/units data, adding a pivot table on A1:C7, assigning Product to rows, Region to columns, summing UnitsSold, refreshing and calculating the pivot, then saving the file as TotalUnitsSoldPivot.xlsx.
    public class TotalUnitsSoldPivot
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // Populate sample data: Product, Region, UnitsSold
                // -------------------------------------------------
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Region");
                cells["C1"].PutValue("UnitsSold");

                // Sample rows
                cells["A2"].PutValue("Apple");   cells["B2"].PutValue("North"); cells["C2"].PutValue(120);
                cells["A3"].PutValue("Apple");   cells["B3"].PutValue("South"); cells["C3"].PutValue(80);
                cells["A4"].PutValue("Banana");  cells["B4"].PutValue("North"); cells["C4"].PutValue(150);
                cells["A5"].PutValue("Banana");  cells["B5"].PutValue("South"); cells["C5"].PutValue(130);
                cells["A6"].PutValue("Cherry");  cells["B6"].PutValue("North"); cells["C6"].PutValue(90);
                cells["A7"].PutValue("Cherry");  cells["B7"].PutValue("South"); cells["C7"].PutValue(110);

                // -------------------------------------------------
                // Create a pivot table based on the data range
                // -------------------------------------------------
                // Data range: A1:C7
                // Destination top‑left cell for the pivot table: E3
                // Pivot table name: "UnitsSoldPivot"
                int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "UnitsSoldPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // -------------------------------------------------
                // Configure pivot fields
                // -------------------------------------------------
                // Row field: Product
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

                // Column field: Region (optional, shows breakdown by region)
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");

                // Data field: UnitsSold – we want the sum of units sold
                int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "UnitsSold");
                PivotField unitsSoldField = pivotTable.DataFields[dataFieldPos];
                unitsSoldField.Function = ConsolidationFunction.Sum; // Explicitly set aggregation to Sum

                // -------------------------------------------------
                // Refresh and calculate the pivot table data
                // -------------------------------------------------
                // Refresh the pivot cache and recalculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("TotalUnitsSoldPivot.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TotalUnitsSoldPivot.Run();
        }
    }
}
