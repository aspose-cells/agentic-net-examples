// Title: Refresh an Excel PivotTable asynchronously with Aspose.Cells for .NET and trigger a completion callback
// Common Searches: c# refresh pivot table without blocking UI using Aspose.Cells | how to run Aspose.Cells pivot refresh on a background thread | invoke a method after Aspose.Cells pivot table refresh completes | example of Task.Run to refresh Excel pivot table in .NET | asynchronous pivot table refresh with callback in C#
// Tags: asynchronous pivot table refresh Aspose.Cells | background thread Excel pivot update .NET | pivot table RefreshData CalculateData Aspose | completion callback after pivot refresh C# | non‑blocking UI pivot operation Aspose.Cells

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // PivotTable resides in this namespace

// The example shows how to refresh a specific PivotTable in an Excel workbook using Aspose.Cells without freezing the UI. The RefreshPivotTableAsync method loads the workbook, locates the worksheet and PivotTable, calls RefreshData and CalculateData, saves the file, and runs all of this on a thread‑pool thread via Task.Run. A continuation invokes a user‑provided Action callback so the UI can be updated after the operation finishes. Error handling and a ManualResetEventSlim usage demonstrate waiting for completion in a console app.
public class PivotTableRefresher
{
    /// <param name="filePath">Path to the Excel file.</param>
    /// <param name="sheetName">Worksheet containing the PivotTable.</param>
    /// <param name="pivotTableName">Name of the PivotTable to refresh.</param>
    /// <param name="onCompleted">Callback invoked after refresh completes.</param>
    public void RefreshPivotTableAsync(string filePath, string sheetName, string pivotTableName, Action onCompleted)
    {
        // Run heavy work on a thread‑pool thread
        Task.Run(() =>
        {
            try
            {
                // Verify the source file exists
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("Excel file not found.", filePath);

                // Load the workbook
                var workbook = new Workbook(filePath);

                // Locate the worksheet
                var sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                    throw new ArgumentException($"Worksheet '{sheetName}' not found.");

                // Find the requested PivotTable
                PivotTable pivotTable = null;
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    if (pt.Name.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        pivotTable = pt;
                        break;
                    }
                }

                if (pivotTable == null)
                    throw new ArgumentException($"PivotTable '{pivotTableName}' not found in worksheet '{sheetName}'.");

                // Refresh data and recalculate the PivotTable
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save changes (overwrite original file)
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                // Log any exception that occurs during processing
                Console.Error.WriteLine("Error refreshing PivotTable:");
                Console.Error.WriteLine(ex.Message);
            }
        })
        .ContinueWith(t =>
        {
            // Invoke the completion callback regardless of success or failure
            onCompleted?.Invoke();
        });
    }
}

public class Program
{
    public static void Main()
    {
        var refresher = new PivotTableRefresher();

        string filePath = @"C:\Data\Report.xlsx";
        string sheetName = "DataSheet";
        string pivotName = "PivotTable1";

        // Use a synchronization primitive to wait for async completion
        using (var doneEvent = new ManualResetEventSlim(false))
        {
            refresher.RefreshPivotTableAsync(filePath, sheetName, pivotName, () =>
            {
                Console.WriteLine("PivotTable refreshed successfully.");
                doneEvent.Set();
            });

            // Wait for the operation to finish
            doneEvent.Wait();
        }
    }
}
