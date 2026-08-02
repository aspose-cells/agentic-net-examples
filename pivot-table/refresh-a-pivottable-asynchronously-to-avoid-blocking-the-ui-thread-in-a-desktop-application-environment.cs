// Title: Asynchronously Refresh an Aspose.Cells PivotTable in a C# Desktop Application
// Description: Shows how to load a workbook, find the first PivotTable, refresh its cache on a background thread using Task.Run, recalculate values, and save the file—eliminating UI blocking in WinForms or WPF apps.
// Keywords: Aspose.Cells async pivot refresh | C# PivotTable background update | Task.Run RefreshData CalculateData | non‑blocking UI Aspose.Cells | desktop .NET pivot cache refresh | WinForms PivotTable refresh | WPF PivotTable async | PivotRefreshState handling
// Common Searches: refresh Aspose.Cells PivotTable without freezing UI | async PivotTable refresh C# WinForms | Task.Run RefreshData Aspose.Cells example | how to recalculate pivot after RefreshData | handle PivotRefreshState errors in async refresh
// Developer Intent: Update a PivotTable in an Aspose.Cells workbook on a background thread to keep the UI responsive.
// Use Cases: Load a workbook, locate the first PivotTable, and refresh its data cache on a separate thread. | Recalculate pivot values after the cache refresh and save the updated workbook. | Detect and log PivotRefreshState results while continuing processing even on failure.
// AI Prompts: Generate a WinForms button click handler that disables the button, awaits an async PivotTable refresh, and re‑enables the button after completion. | Create a method that returns true only when RefreshData returns PivotRefreshState.Success, otherwise logs the issue. | Provide robust error‑handling code for async pivot refresh that captures exceptions, logs PivotRefreshState, and falls back to CalculateData.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsAsyncPivotRefresh
{
    // Shows how to load a workbook, find the first PivotTable, refresh its cache on a background thread using Task.Run, recalculate values, and save the file—eliminating UI blocking in WinForms or WPF apps.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "PivotData.xlsx";
            string outputPath = "PivotData_Refreshed.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule)
                Workbook workbook = new Workbook(inputPath);

                // Assume the pivot table is in the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure a pivot table exists
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found.");
                    return;
                }

                PivotTable pivotTable = worksheet.PivotTables[0];

                // Refresh the pivot table asynchronously
                await RefreshPivotTableAsync(pivotTable);

                // Save the workbook after refresh (lifecycle rule)
                workbook.Save(outputPath);

                Console.WriteLine($"Pivot table refreshed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="pivotTable">The pivot table to refresh.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        private static Task RefreshPivotTableAsync(PivotTable pivotTable)
        {
            // Execute potentially time‑consuming operations off the UI thread
            return Task.Run(() =>
            {
                // Refresh the pivot cache from the data source
                PivotRefreshState state = pivotTable.RefreshData();

                // Optionally handle refresh state
                if (state != PivotRefreshState.Success)
                {
                    Console.WriteLine("Pivot refresh encountered an issue, proceeding with calculation.");
                }

                // Recalculate the pivot table values after the refresh
                pivotTable.CalculateData();
            });
        }
    }
}
