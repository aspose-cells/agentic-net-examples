// Title: Measure and log the time taken to refresh an Excel pivot table cache using Aspose.Cells for .NET
// Common Searches: how can I profile the RefreshData call of a pivot table with Aspose.Cells in C# | log the elapsed milliseconds of Excel pivot table recalculation using Aspose.Cells .NET | benchmark the time required to load pivot cache records with Aspose.Cells | using Stopwatch to measure Aspose.Cells pivot table refresh performance | determine pivot table refresh duration programmatically in a .NET workbook
// Tags: Aspose.Cells pivot table refresh timing | C# Stopwatch Excel pivot performance | measure RefreshData execution .NET | log pivot cache load duration | benchmark Aspose.Cells pivot recalculation

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The example loads an Excel workbook, verifies a pivot table exists, starts a Stopwatch, refreshes and recalculates the pivot table data with RefreshData and CalculateData, stops the timer, outputs the elapsed milliseconds, ensures the output directory exists, saves the workbook, and handles any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure that at least one pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Start timing the refresh operation
            Stopwatch timer = Stopwatch.StartNew();

            // Refresh the pivot table data (loads cache records and recalculates)
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Stop timing
            timer.Stop();

            // Log the duration required to refresh the pivot table
            Console.WriteLine($"Pivot table refresh duration: {timer.ElapsedMilliseconds} ms");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook if you need to persist changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
