// Title: Skip Corrupted Excel Workbooks in Batch Processing with Aspose.Cells for .NET
// Description: Shows how to load multiple Excel files using Aspose.Cells, catch CellsException (FileCorrupted) and generic errors, log the issue, optionally enable RepairLoad, and continue processing the remaining workbooks.
// Keywords: Aspose.Cells | C# | .NET | batch workbook processing | error handling | corrupted Excel file | CellsException | FileCorrupted | RepairLoad | continue loop | Excel automation
// Common Searches: Aspose.Cells ignore corrupted file | C# batch process Excel workbooks with error handling | skip bad Excel files Aspose | continue after CellsException | repair load Aspose.Cells example
// Developer Intent: Add try‑catch logic so a failed workbook load does not abort the entire batch operation.
// Use Cases: Automated reporting pipelines that handle dozens of spreadsheets. | Data migration where some source files may be damaged. | Scheduled server jobs that must finish even if individual files are unreadable. | Applying RepairLoad to attempt recovery before moving on to the next file.
// AI Prompts: Write C# code that iterates over a list of .xlsx paths, loads each with Aspose.Cells, catches CellsException with ExceptionType.FileCorrupted, logs a warning, and proceeds to the next file. | Show how to enable Workbook.Settings.RepairLoad and handle generic exceptions during batch Excel processing.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // Shows how to load multiple Excel files using Aspose.Cells, catch CellsException (FileCorrupted) and generic errors, log the issue, optionally enable RepairLoad, and continue processing the remaining workbooks.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            List<string> workbookFiles = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "CorruptedWorkbook.xlsx", // Example of a corrupted file
                "Workbook3.xlsx"
            };

            // Output directory for processed workbooks
            string outputDir = "ProcessedWorkbooks";

            // Ensure the output directory exists
            System.IO.Directory.CreateDirectory(outputDir);

            foreach (string filePath in workbookFiles)
            {
                try
                {
                    // Load the workbook (uses the Workbook(string) constructor)
                    Workbook workbook = new Workbook(filePath);

                    // Enable repair mode for future operations (optional but demonstrates usage)
                    workbook.Settings.RepairLoad = true;

                    // Example processing: write the number of worksheets to console
                    Console.WriteLine($"Loaded '{filePath}' successfully. Worksheets count: {workbook.Worksheets.Count}");

                    // Save the processed workbook to the output folder (uses Workbook.Save(string))
                    string outputPath = System.IO.Path.Combine(outputDir, System.IO.Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                    Console.WriteLine($"Saved processed workbook to '{outputPath}'.");
                }
                catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
                {
                    // Specific handling for corrupted files – log and continue with next file
                    Console.WriteLine($"[Warning] The file '{filePath}' is corrupted (ExceptionType.FileCorrupted). Skipping this workbook.");
                    continue;
                }
                catch (Exception ex)
                {
                    // General error handling – log and continue
                    Console.WriteLine($"[Error] Failed to process '{filePath}'. Reason: {ex.Message}");
                    continue;
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
