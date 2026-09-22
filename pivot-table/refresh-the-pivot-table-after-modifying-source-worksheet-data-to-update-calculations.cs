// Title: Refresh all PivotTables in an Excel workbook after modifying source data with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that changes a cell in a worksheet, then iterates through every PivotTable in the workbook to call RefreshData and CalculateData using Aspose.Cells. | Generate a .NET snippet that loads an existing .xlsx file, updates source data, refreshes all pivot caches, and saves the workbook. | Provide an example of programmatically refreshing PivotTables after data changes in Aspose.Cells, including error handling for missing files.
// Common Searches: Aspose.Cells C# how to refresh pivot tables after updating source worksheet | programmatically recalculate all PivotTables in an Excel file using Aspose.Cells .NET | C# example to modify cell value and refresh pivot cache with Aspose.Cells | refresh pivot table data source in .xlsx using Aspose.Cells library
// Tags: Aspose.Cells refresh pivot tables C# | update source worksheet Aspose.Cells | PivotTable RefreshData CalculateData .NET | iterate worksheets refresh all pivots | save workbook after pivot refresh Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // For PivotTable class

// The example loads an existing workbook, updates cell B2 in the "Data" worksheet, iterates through every worksheet and its PivotTables to invoke RefreshData and CalculateData, and then saves the modified workbook to a new XLSX file.
class PivotTableRefreshExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the source worksheet that contains the data for the pivot table
            Worksheet sourceSheet = workbook.Worksheets["Data"]; // or use index: workbook.Worksheets[0]

            // Modify source data (example: update cell B2)
            sourceSheet.Cells["B2"].PutValue(12345);

            // Refresh all pivot tables in the workbook to reflect the updated data
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the pivot table's cache with the latest source data
                    pivotTable.RefreshData();

                    // Recalculate the pivot table (necessary for totals, subtotals, etc.)
                    pivotTable.CalculateData();
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
