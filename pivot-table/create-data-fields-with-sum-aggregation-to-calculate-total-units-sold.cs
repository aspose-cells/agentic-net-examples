// Title: C# – Create a Pivot Table with Sum Aggregation for Total Units Sold using Aspose.Cells
// Description: This example demonstrates how to build a new workbook, populate it with region, product, and units‑sold data, add a pivot table on range A1:C5 at cell E3, set the Units Sold field to use the Sum consolidation function, refresh and calculate the pivot, and save the file as TotalUnitsSoldPivot.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# pivot table | sum aggregation | total units sold | ConsolidationFunction.Sum | Excel report automation | Aspose.Cells .NET example | pivot data refresh | calculate pivot data
// Common Searches: Aspose.Cells set sum aggregation for pivot data field | C# create pivot table total units sold Aspose | how to refresh and calculate pivot table Aspose.Cells | example of pivot table with sum function in .NET | add pivot table from range A1:C5 using Aspose.Cells
// Developer Intent: Generate a pivot table that sums the Units Sold column to display total sales per region and product.
// Use Cases: Produce management‑ready sales summaries grouped by region and product. | Automate Excel report generation in a .NET service that shows total units sold. | Add a reusable routine to any workbook for creating sum‑aggregated sales pivots.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot table that calculates the sum of a numeric column. | Show how to change a pivot data field’s aggregation function to Average in Aspose.Cells. | Provide an example of refreshing and recalculating pivot data after modifying the source range with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This example demonstrates how to build a new workbook, populate it with region, product, and units‑sold data, add a pivot table on range A1:C5 at cell E3, set the Units Sold field to use the Sum consolidation function, refresh and calculate the pivot, and save the file as TotalUnitsSoldPivot.xlsx with Aspose.Cells for .NET.
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

                // Populate sample data
                cells["A1"].Value = "Region";
                cells["B1"].Value = "Product";
                cells["C1"].Value = "Units Sold";

                cells["A2"].Value = "North";
                cells["B2"].Value = "Widget";
                cells["C2"].Value = 120;

                cells["A3"].Value = "North";
                cells["B3"].Value = "Gadget";
                cells["C3"].Value = 80;

                cells["A4"].Value = "South";
                cells["B4"].Value = "Widget";
                cells["C4"].Value = 150;

                cells["A5"].Value = "South";
                cells["B5"].Value = "Gadget";
                cells["C5"].Value = 70;

                // Add a pivot table based on the data range A1:C5, placed at E3
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotUnitsSold");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Units Sold");

                // Set aggregation to Sum (default)
                PivotField unitsSoldField = pivotTable.DataFields[dataFieldIndex];
                unitsSoldField.Function = ConsolidationFunction.Sum;

                // Refresh and calculate pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("TotalUnitsSoldPivot.xlsx");
                Console.WriteLine("Workbook saved as TotalUnitsSoldPivot.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
