// Title: C# – Aspose.Cells PivotTable Show Zero for Blank Cells
// Description: Demonstrates how to set the worksheet DisplayZeros property and configure an Aspose.Cells PivotTable so that empty data cells appear as 0 in the generated Excel report.
// Keywords: Aspose.Cells C# | PivotTable display zeros | show zero for blank cells | DisplayZeros property | Aspose.Cells pivot options | Excel automation .NET | pivot table empty cell zero | Aspose.Cells example
// Common Searches: Aspose.Cells show zero values in pivot table | C# pivot table blank cells as zero Aspose | DisplayZeros property Aspose.Cells example | how to display zeros in Aspose.Cells pivot | Aspose.Cells pivot table options for empty cells
// Developer Intent: Enable a PivotTable to render empty cells as zero values using Aspose.Cells in C#.
// Use Cases: Sales summary where missing sales figures are displayed as 0. | Financial workbook that treats blank expense entries as zero during aggregation. | Inventory report that shows zero quantity for items without recorded stock.
// AI Prompts: Generate C# code with Aspose.Cells that creates a PivotTable and shows zero for blank data cells. | Explain how to apply the DisplayZeros property to a worksheet and ensure the PivotTable reflects zero values. | Provide a step‑by‑step Aspose.Cells example that refreshes and calculates a PivotTable after enabling zero display for empty cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to set the worksheet DisplayZeros property and configure an Aspose.Cells PivotTable so that empty data cells appear as 0 in the generated Excel report.
    public class PivotTableDisplayZeroValuesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (some rows have no numeric value)
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["A3"].PutValue("Banana");
                // B3 left empty – will be treated as zero in the pivot table
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(0); // explicit zero

                // Ensure the worksheet displays zero values (default is true, but set explicitly)
                sheet.DisplayZeros = true;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D2", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Product as row field, Sales as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTableDisplayZeroValues.xlsx");
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
            PivotTableDisplayZeroValuesDemo.Run();
        }
    }
}
