// Title: Refresh all PivotTables in an Excel workbook asynchronously using Aspose.Cells for .NET and trigger a UI callback after completion
// AI Prompts: Write C# code that loads a workbook with Aspose.Cells, executes workbook.Worksheets.RefreshPivotTables() on a Task thread, saves the file, and invokes a supplied Action delegate when the refresh finishes. | Show how to combine Task.Run and ContinueWith to perform a pivot‑table refresh in the background and notify the UI layer via a callback method in a .NET application.
// Common Searches: how to refresh pivot tables asynchronously with Aspose.Cells in C# | run Aspose.Cells RefreshPivotTables on a background thread and get notified | C# Aspose.Cells refresh all pivot tables without blocking UI | callback after Aspose.Cells workbook save task in .NET | asynchronous pivot table refresh example using Aspose.Cells for .NET
// Tags: async RefreshPivotTables Aspose.Cells | background thread Excel workbook processing .NET | UI callback after Aspose.Cells async task | programmatic pivot table refresh C# | save workbook after pivot refresh Aspose.Cells

using System;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    // Loads 'PivotData.xlsx', refreshes every pivot table on a Task thread, saves the workbook, and calls OnRefreshCompleted to update the UI.
    class Program
    {
        // Path to the Excel file containing the pivot table
        private const string WorkbookPath = "PivotData.xlsx";

        static void Main()
        {
            // Start the background refresh operation
            RefreshPivotTableAsync(WorkbookPath, OnRefreshCompleted);
            
            // Simulate UI thread work (e.g., keep console alive)
            Console.WriteLine("Refresh started in background thread...");
            Console.ReadLine(); // Prevent application exit until user presses Enter
        }

        /// <param name="filePath">Path to the workbook file.</param>
        /// <param name="callback">Method to call after refresh finishes.</param>
        private static void RefreshPivotTableAsync(string filePath, Action callback)
        {
            // Run the refresh logic on a separate thread
            Task.Run(() =>
            {
                // Load the workbook (uses Aspose.Cells load rule)
                Workbook workbook = new Workbook(filePath);

                // Refresh all pivot tables in the workbook (uses RefreshPivotTables rule)
                workbook.Worksheets.RefreshPivotTables();

                // Save the workbook after refresh (uses Aspose.Cells save rule)
                workbook.Save(filePath);
            })
            .ContinueWith(t =>
            {
                // Invoke the UI callback after the background task completes
                callback?.Invoke();
            });
        }

        /// <summary>
        /// Callback method executed after the pivot tables have been refreshed.
        /// This simulates updating the UI (e.g., displaying a message).
        /// </summary>
        private static void OnRefreshCompleted()
        {
            // UI update logic – here we simply write to the console
            Console.WriteLine("Pivot tables have been refreshed and workbook saved.");
        }
    }
}
