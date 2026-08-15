// Title: Refresh and Recalculate a PivotTable with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, populate source data, add a PivotTable, refresh its cache, recalculate the pivot values, and save the result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | PivotTable refresh | CalculateData | RefreshData | update pivot cache | programmatic pivot table | Excel automation .NET | Aspose.Cells example | refresh pivot table .NET
// Common Searches: Aspose.Cells refresh PivotTable C# example | How to recalculate PivotTable with Aspose.Cells | RefreshData and CalculateData Aspose.Cells code | Update pivot cache programmatically .NET | PivotTable automation Aspose.Cells
// Developer Intent: Programmatically refresh a PivotTable’s cache and recompute its data in a .NET workbook.
// Use Cases: Generate a sales summary pivot, modify source rows, then call RefreshData and CalculateData before exporting. | Automate monthly reporting where source data changes and the pivot must stay current without manual intervention. | Create Excel files with ready‑to‑use pivots for downstream systems that require pre‑calculated results.
// AI Prompts: Write C# code using Aspose.Cells to refresh a PivotTable after source data changes and then calculate the pivot values. | Explain when to use RefreshData versus RefreshDataOnOpen in Aspose.Cells and how CalculateData fits into the workflow. | Provide a step‑by‑step guide to programmatically update a PivotTable cache and recalculate totals with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, populate source data, add a PivotTable, refresh its cache, recalculate the pivot values, and save the result using Aspose.Cells for .NET.
    public class RefreshAndCalculatePivotDemo
    {
        public static void Main()
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

                // Populate source data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Drink");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Food");
                sheet.Cells["B4"].PutValue(200);

                // Add a pivot table that uses the source range A1:B4 and place it at D1
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Define the row and data fields for the pivot table
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh the pivot cache from the source data
                // The RefreshData method on PivotTable is still available (though marked obsolete)
                pivot.RefreshData();

                // Calculate the pivot data and write it to the worksheet
                pivot.CalculateData();

                // Save the workbook with the refreshed and calculated pivot table
                workbook.Save("RefreshAndCalculatePivot.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
                throw;
            }
        }
    }
}
