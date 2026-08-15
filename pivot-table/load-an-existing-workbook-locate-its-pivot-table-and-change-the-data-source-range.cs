// Title: Aspose.Cells for .NET – Change Pivot Table Data Source Range in an Existing Workbook (C#)
// Description: Loads an existing Excel file, accesses the first worksheet, verifies a pivot table exists, assigns a new source range with ChangeDataSource, refreshes the pivot, and saves the updated workbook. Demonstrates how to re‑point a pivot table to a different cell block without recreating it.
// Keywords: Aspose.Cells | C# | PivotTable | ChangeDataSource | Update pivot source | Refresh pivot | Excel workbook manipulation | Modify pivot table range
// Common Searches: Aspose.Cells change pivot table source range C# | How to update pivot table data source with Aspose.Cells | ChangeDataSource example for PivotTable in .NET | Refresh pivot after changing source in Aspose.Cells | Load workbook and edit pivot table programmatically
// Developer Intent: Load an existing workbook, locate a pivot table, set a new data source range, refresh the pivot, and save the file.
// Use Cases: Redirect a pivot table to a newly expanded data block after adding rows. | Switch a pivot table to a different worksheet range without rebuilding the layout. | Automate batch updates of multiple workbooks to point all pivots to a standardized source range.
// AI Prompts: Write C# code using Aspose.Cells that changes the data source of a pivot table identified by its name. | Show how to iterate through every pivot table in a workbook and assign each a new source range. | Explain how to preserve pivot table formatting while updating its data source with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // Loads an existing Excel file, accesses the first worksheet, verifies a pivot table exists, assigns a new source range with ChangeDataSource, refreshes the pivot, and saves the updated workbook. Demonstrates how to re‑point a pivot table to a different cell block without recreating it.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook that contains a pivot table
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is in the first worksheet; adjust index as needed
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Define the new data source range.
            // The array contains the source range and the sheet name.
            // Example: data in C1:D10 on the same sheet ("Sheet1")
            string[] newDataSource = new string[] { "C1:D10", worksheet.Name };

            // Change the data source of the pivot table
            pivotTable.ChangeDataSource(newDataSource);

            // Refresh the pivot table to reflect the new source data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Pivot table data source changed and workbook saved to '{outputPath}'.");
        }
    }
}
