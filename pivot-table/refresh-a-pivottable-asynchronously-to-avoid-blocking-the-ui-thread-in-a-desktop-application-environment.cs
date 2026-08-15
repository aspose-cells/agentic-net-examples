// Title: Asynchronous PivotTable Refresh with Aspose.Cells for .NET – Keep UI Responsive
// Description: Demonstrates how to load an Excel workbook, refresh all PivotTables, recalculate their data, and save the file using Aspose.Cells in C#. All operations are wrapped in Task.Run and awaited, allowing the work to run on background threads and preventing UI thread blockage in WinForms or WPF applications.
// Keywords: Aspose.Cells async pivot refresh | C# refresh PivotTable background thread | non‑blocking Excel pivot update | Task.Run Aspose.Cells | .NET desktop UI responsiveness | RefreshPivotTables asynchronous
// Common Searches: refresh pivot tables asynchronously Aspose.Cells | C# non blocking pivot refresh example | how to keep UI responsive while updating Excel pivots | Aspose.Cells RefreshPivotTables on background thread | async calculatedata for pivot tables .NET
// Developer Intent: Update all PivotTables in an Excel workbook on a background thread to avoid freezing the desktop UI.
// Use Cases: Refresh large workbooks in a WinForms/WPF app without UI lag. | Integrate async pivot updates into a reporting service that runs alongside user interactions. | Batch‑process multiple Excel files, refreshing and saving each workbook concurrently to improve throughput.
// AI Prompts: Generate a C# async method that refreshes PivotTables with Aspose.Cells and reports progress to a progress bar. | Show how to add cancellation support to the asynchronous pivot refresh routine in a WPF MVVM command. | Create robust error handling for the async workflow, covering missing files, load failures, and save exceptions.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotAsyncDemo
{
    // Demonstrates how to load an Excel workbook, refresh all PivotTables, recalculate their data, and save the file using Aspose.Cells in C#. All operations are wrapped in Task.Run and awaited, allowing the work to run on background threads and preventing UI thread blockage in WinForms or WPF applications.
    class Program
    {
        // Entry point for a console application.
        static async Task Main(string[] args)
        {
            try
            {
                // Path to the Excel file containing the pivot table.
                string inputPath = "PivotData.xlsx";
                string outputPath = "PivotData_Refreshed.xlsx";

                // Refresh the pivot tables without blocking the UI thread.
                await RefreshPivotTablesAsync(inputPath, outputPath);

                Console.WriteLine("Pivot tables refreshed and workbook saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Asynchronously loads a workbook, refreshes all pivot tables, and saves the result.
        private static async Task RefreshPivotTablesAsync(string sourceFile, string destinationFile)
        {
            // Verify that the source file exists before attempting to load it.
            if (!File.Exists(sourceFile))
                throw new FileNotFoundException($"Source file not found: {sourceFile}");

            // Load the workbook on a background thread.
            Workbook workbook = await Task.Run(() => new Workbook(sourceFile));

            // Refresh all pivot tables in the workbook on a background thread.
            await Task.Run(() => workbook.Worksheets.RefreshPivotTables());

            // Recalculate the pivot data after refresh.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    await Task.Run(() => pt.CalculateData());
                }
            }

            // Save the updated workbook on a background thread.
            await Task.Run(() => workbook.Save(destinationFile));
        }
    }
}
