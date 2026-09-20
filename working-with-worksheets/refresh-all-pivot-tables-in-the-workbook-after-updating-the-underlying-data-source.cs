// Title: Refresh every PivotTable in an Excel workbook after modifying source data with Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file using Aspose.Cells, change values in the source worksheet, then loop through all worksheets and call PivotTable.RefreshData() followed by PivotTable.CalculateData() for each pivot table before saving the workbook. | Write C# code that verifies the input workbook exists, updates a specific cell, iterates over Workbook.Worksheets and their PivotTables collection to refresh the pivot cache and recalculate results, and handles any runtime exceptions. | Create a reusable method that accepts a workbook path, updates the data sheet, refreshes all pivot tables via the Aspose.Cells API, and returns the path of the newly saved workbook with refreshed pivots.
// Common Searches: C# Aspose.Cells how to refresh all pivot tables after changing data source | programmatically update pivot cache in .xlsx using Aspose.Cells for .NET | refresh multiple PivotTable objects in a workbook with Aspose.Cells C# example | Aspose.Cells RefreshData CalculateData on each PivotTable after editing cells
// Tags: aspocells pivot cache update | c# refresh every pivot table | modify workbook data source aspocells | iterate worksheets for pivot refresh | save workbook after pivot refresh aspocells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefresh
{
    // The sample loads InputWorkbook.xlsx, optionally modifies a cell in the 'Data' sheet, iterates through every worksheet and each PivotTable to invoke RefreshData() and CalculateData(), then saves the result as OutputWorkbook.xlsx while handling missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook containing the data source and pivot tables
                Workbook workbook = new Workbook(inputPath);

                // (Optional) Update the underlying data source here
                // Example: modify cells in the data sheet
                Worksheet dataSheet = workbook.Worksheets["Data"];
                if (dataSheet != null)
                {
                    dataSheet.Cells["A2"].PutValue(12345); // sample update
                }

                // Refresh all pivot tables in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each pivot table on the current worksheet
                    foreach (PivotTable pivotTable in sheet.PivotTables)
                    {
                        // Refresh the pivot cache with the latest data
                        pivotTable.RefreshData();

                        // Recalculate the pivot table values after refresh
                        pivotTable.CalculateData();
                    }
                }

                // Save the workbook with refreshed pivot tables
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
