// Title: C# CLI Tool to Batch Recalculate Excel Workbooks with Aspose.Cells
// Description: A .NET command‑line utility that accepts a folder path, scans for Excel and CSV files, loads each workbook with Aspose.Cells, recalculates all formulas, refreshes pivots and charts, saves the changes, and prints total files and worksheets processed.
// Keywords: Aspose.Cells | C# | .NET | command line tool | batch recalculate formulas | Excel workbook processing | refresh pivots | save workbook | folder scan | CSV support | GitHub example
// Common Searches: C# batch recalculate Excel formulas Aspose.Cells | command line tool to process multiple workbooks | how to refresh pivots in Excel files using Aspose | calculate all formulas in a folder of spreadsheets | Aspose.Cells CLI example for .NET
// Developer Intent: Recalculate formulas in every supported workbook within a given directory and report the number of files and worksheets processed.
// Use Cases: Automate nightly recalculation of financial models stored in a shared folder before distribution. | Refresh pivot tables and charts across a batch of monthly reports to ensure up‑to‑date visuals. | Provide a quick audit of how many worksheets were updated after bulk processing.
// AI Prompts: Generate a PowerShell script that runs the compiled .NET tool for a target folder and writes the summary to a log file. | Extend the program to search subfolders recursively while maintaining accurate total counts. | Add robust error handling that writes detailed exceptions to a separate log and returns distinct exit codes for success, partial failure, and total failure.

using System;
using System.IO;
using Aspose.Cells;

namespace RecalculateWorkbooks
{
    // A .NET command‑line utility that accepts a folder path, scans for Excel and CSV files, loads each workbook with Aspose.Cells, recalculates all formulas, refreshes pivots and charts, saves the changes, and prints total files and worksheets processed.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input folder path
            if (args.Length == 0 || !Directory.Exists(args[0]))
            {
                Console.WriteLine("Please provide a valid folder path as the first argument.");
                return;
            }

            string folderPath = args[0];
            string[] supportedExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv" };

            int totalFiles = 0;
            int totalWorksheets = 0;

            // Iterate through all supported files in the folder (non‑recursive)
            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                if (Array.IndexOf(supportedExtensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue; // skip unsupported files

                try
                {
                    // Load workbook (create‑load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Recalculate all formulas in the workbook
                    workbook.CalculateFormula();

                    // Optional: refresh pivot tables and charts
                    workbook.Worksheets.RefreshAll();

                    // Save the workbook back to the same file (create‑save rule)
                    workbook.Save(filePath);

                    // Update statistics
                    totalFiles++;
                    totalWorksheets += workbook.Worksheets.Count;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            // Output summary statistics
            Console.WriteLine("Processing completed.");
            Console.WriteLine($"Total workbooks processed: {totalFiles}");
            Console.WriteLine($"Total worksheets processed: {totalWorksheets}");
        }
    }
}
