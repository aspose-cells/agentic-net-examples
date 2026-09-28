// Title: Asynchronously refresh an Excel PivotTable with Aspose.Cells in a C# desktop app to keep the UI responsive
// AI Prompts: Write an async C# method that opens an .xlsx file with Aspose.Cells, refreshes the first PivotTable on a background thread using Task.Run, and saves the workbook to a new file. | Extend the example to iterate over all worksheets and refresh every PivotTable, accepting a CancellationToken so the operation can be cancelled mid‑process. | Add comprehensive error handling that catches any exception during the pivot refresh, logs relevant details, and rethrows it wrapped in an ApplicationException while preserving the original stack trace.
// Common Searches: how to refresh an Excel pivot table asynchronously using Aspose.Cells in C# | prevent UI freeze when updating pivot cache with Aspose.Cells .NET | run Aspose.Cells PivotTable.RefreshData on a background thread | c# async method to refresh pivot tables in large workbook without blocking UI | Aspose.Cells example for non‑blocking pivot table refresh in WinForms
// Tags: aspose.cells async pivot refresh | c# background thread refreshdata | excel workbook non blocking pivot update | aspose.cells refreshdata obsolete suppression | c# cancellationtoken pivot table refresh

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample defines an async RefreshPivotTableAsync method that loads an .xlsx workbook with Aspose.Cells, verifies a PivotTable exists, refreshes its data on a background thread via Task.Run (suppressing the obsolete RefreshData warning), saves the updated file, and wraps any exception in an ApplicationException. The Main method demonstrates calling the async method, enabling UI‑responsive pivot refresh in desktop .NET applications.
class Program
{
    // Refreshes the first pivot table in the first worksheet of the workbook.
    // inputPath  - path to the source workbook.
    // outputPath - path where the updated workbook will be saved.
    static async Task RefreshPivotTableAsync(string inputPath, string outputPath)
    {
        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException($"Input file not found: {inputPath}");
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust if needed).
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table.
            if (sheet.PivotTables.Count == 0)
            {
                throw new InvalidOperationException("No pivot tables found in the first worksheet.");
            }

            // Get the first pivot table.
            PivotTable pivotTable = sheet.PivotTables[0];

            // Refresh the pivot table on a background thread.
            await Task.Run(() =>
            {
                // Refresh the pivot table data (method is marked obsolete but still functional).
#pragma warning disable CS0618 // Suppress obsolete warning
                pivotTable.RefreshData();
#pragma warning restore CS0618
            });

            // Save the updated workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to let the caller handle it.
            throw new ApplicationException("Error while refreshing pivot table.", ex);
        }
    }

    // Example usage from a console application.
    static async Task Main(string[] args)
    {
        try
        {
            // Adjust paths as appropriate.
            await RefreshPivotTableAsync("input.xlsx", "output.xlsx");
            Console.WriteLine("Pivot table refreshed and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Operation failed: {ex.Message}");
        }
    }
}
