// Title: Log the duration of a pivot table refresh after loading cache records with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that opens an XLSX workbook, calls PivotTable.RefreshData, measures the elapsed time with Stopwatch, prints the milliseconds to the console, and saves the file. | Add comprehensive try‑catch error handling to the pivot refresh routine and redirect the measured refresh duration to a log file instead of the console. | Extend the example to iterate over all pivot tables in a workbook, refresh each one, calculate formulas, and write each table's refresh time to a CSV performance report.
// Common Searches: Aspose.Cells C# how to measure pivot table refresh time after loading cache | log execution time of PivotTable.RefreshData in a .NET application | benchmark Excel pivot refresh performance using Stopwatch and Aspose.Cells | record pivot table refresh duration and save to a log file with Aspose.Cells | C# example for timing pivot table refresh and saving the workbook
// Tags: Aspose.Cells pivot table refresh timing | C# Stopwatch measurement Aspose.Cells | log pivot refresh duration .NET | refreshdata calculateformula performance | measure Excel pivot refresh Aspose.Cells | pivot table performance logging C#

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// Opens an Excel workbook, refreshes the first pivot table, measures the refresh time with Stopwatch, logs the elapsed milliseconds, recalculates formulas, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            // Retrieve the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Start measuring the refresh duration
            Stopwatch stopwatch = Stopwatch.StartNew();

            // Refresh the pivot table data
            pivotTable.RefreshData();

            // Recalculate all formulas in the workbook (which also updates the pivot table)
            workbook.CalculateFormula();

            // Stop the timer
            stopwatch.Stop();

            // Log the elapsed time (in milliseconds)
            Console.WriteLine($"Pivot table refresh duration: {stopwatch.ElapsedMilliseconds} ms");

            // Save the workbook if further processing is required
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
