// Title: Create a Pivot Table on a New Worksheet with Aspose.Cells for .NET
// Description: This example shows how to generate a workbook, add sample data, determine the full data range with MaxDisplayRange, create a new worksheet, insert a pivot table at a specified cell, assign "Category" to rows and "Amount" to values, refresh the cache, calculate results, and save the file as PivotTableDemo.xlsx using C# and Aspose.Cells.
// Keywords: Aspose.Cells | .NET | C# | pivot table | add pivot table | source range | destination cell | MaxDisplayRange | RefreshData | CalculateData | PivotTableCollection | PivotFieldType
// Common Searches: how to add a pivot table to a new worksheet using Aspose.Cells | specify source range and destination cell for Aspose.Cells pivot table | programmatically set row and data fields in Aspose.Cells pivot table | Aspose.Cells C# create pivot table from dynamic range | refresh and calculate pivot data with Aspose.Cells
// Developer Intent: Add a pivot table to a separate worksheet by defining the source data range and the target cell programmatically.
// Use Cases: Automatically generate a sales summary pivot table from raw data in the same workbook. | Create reporting templates that add a pivot table to a fresh worksheet, handling variable data sizes with MaxDisplayRange. | Refresh and calculate pivot results after populating source cells before exporting the workbook for downstream processing.
// AI Prompts: Show code to add a column field and format the "Amount" values as currency in an Aspose.Cells pivot table. | Provide an example of applying a built‑in pivot style (e.g., PivotStyleMedium9) after creating the pivot table. | Explain how to rename a pivot table and move it to another worksheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotDemo
{
    // This example shows how to generate a workbook, add sample data, determine the full data range with MaxDisplayRange, create a new worksheet, insert a pivot table at a specified cell, assign "Category" to rows and "Amount" to values, refresh the cache, calculate results, and save the file as PivotTableDemo.xlsx using C# and Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and add sample data
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "SourceData";

                // Populate sample data
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Amount");
                dataSheet.Cells["A2"].PutValue("Food");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["A3"].PutValue("Beverage");
                dataSheet.Cells["B3"].PutValue(80);
                dataSheet.Cells["A4"].PutValue("Food");
                dataSheet.Cells["B4"].PutValue(150);
                dataSheet.Cells["A5"].PutValue("Beverage");
                dataSheet.Cells["B5"].PutValue(90);

                // Add a new worksheet that will contain the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Determine the source data range (including the sheet name)
                AsposeRange sourceRange = dataSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{sourceRange.Address}";

                // Add a pivot table to the pivot sheet
                // Parameters: source data range, destination cell (upper‑left corner), pivot table name
                PivotTableCollection pivots = pivotSheet.PivotTables;
                int pivotIndex = pivots.Add(sourceData, "A1", "MyPivotTable");

                // Retrieve the created pivot table
                PivotTable pivotTable = pivots[pivotIndex];

                // Configure the pivot table: put "Category" in rows and "Amount" in data area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh the pivot cache and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTableDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
