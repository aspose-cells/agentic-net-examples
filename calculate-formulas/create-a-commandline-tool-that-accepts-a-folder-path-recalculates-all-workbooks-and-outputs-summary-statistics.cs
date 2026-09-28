// Title: C# command‑line utility using Aspose.Cells to recalculate formulas in all Excel workbooks within a folder and report processing statistics
// AI Prompts: Write a C# console program that takes a directory path, loops through every .xls and .xlsx file, loads each workbook with Aspose.Cells, invokes CalculateFormula, saves the file, and prints the workbook name together with its worksheet count. | Enhance the batch recalculation utility to create a CSV log that records for each processed workbook: file name, worksheet count, any error messages, and the processing timestamp. | Add an optional command‑line switch that disables saving the modified workbooks and instead outputs the total number of formulas that would be recalculated across all files.
// Common Searches: how to batch recalculate formulas in multiple Excel files using Aspose.Cells C# console app | C# program to process all .xls and .xlsx files in a directory and count worksheets | Aspose.Cells calculate all formulas in a folder of workbooks and display summary statistics | command line tool to update Excel workbooks with Aspose.Cells and log processing results | run formula recalculation on every Excel file in a folder with C# and Aspose.Cells
// Tags: Aspose.Cells calculateformula batch processing | C# console recalculate Excel formulas | process multiple xlsx files Aspose.Cells | Excel workbook batch update command line | summary statistics worksheet count Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace WorkbookRecalculator
{
    // A C# console application that accepts a folder path, finds all .xls and .xlsx files, recalculates each workbook's formulas using Aspose.Cells, saves the changes, and prints per‑file details along with total files and worksheets processed.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a folder path was provided
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WorkbookRecalculator <folderPath>");
                return;
            }

            string folderPath = args[0];

            // Check that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Find Excel files in the folder (xlsx and xls)
            var excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (!excelFiles.Any())
            {
                Console.WriteLine("No Excel files found in the specified folder.");
                return;
            }

            int processedFileCount = 0;
            int totalWorksheetCount = 0;

            foreach (var filePath in excelFiles)
            {
                try
                {
                    // Ensure the file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook
                    var workbook = new Workbook(filePath);

                    // Recalculate all formulas in the workbook
                    workbook.CalculateFormula();

                    // Save the workbook back to the same file
                    workbook.Save(filePath);

                    // Update statistics
                    processedFileCount++;
                    totalWorksheetCount += workbook.Worksheets.Count;

                    // Placeholder summary (AI features not available in this SDK version)
                    string summary = "Summary not available (Aspose.Cells.AI not referenced).";

                    // Output per‑file information
                    Console.WriteLine($"--- Processed: {Path.GetFileName(filePath)} ---");
                    Console.WriteLine($"Summary: {summary}");
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    // Report any errors but continue processing other files
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            // Output overall statistics
            Console.WriteLine("=== Summary Statistics ===");
            Console.WriteLine($"Total files processed: {processedFileCount}");
            Console.WriteLine($"Total worksheets processed: {totalWorksheetCount}");
        }
    }
}
