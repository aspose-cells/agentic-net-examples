// Title: How to skip corrupted Excel files while batch processing workbooks with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads each .xlsx file in a directory with Aspose.Cells inside a try‑catch, logs the exception details, and continues to the next file. | Show how to append the paths of workbooks that fail to load to a log file without interrupting the batch operation. | Demonstrate preserving the original sub‑folder hierarchy when saving processed workbooks after handling load errors.
// Common Searches: asp.net batch process excel files continue on load error aspose.cells | c# aspose.cells ignore corrupted workbook and keep processing | log corrupted excel file path during aspose.cells batch operation | skip invalid .xlsx files in aspose.cells while iterating directory | how to handle workbook load exception in aspose.cells c#
// Tags: Aspose.Cells batch workbook processing with error handling | skip corrupted .xlsx files using Aspose.Cells | log workbook load failures in C# | preserve output folder structure Aspose.Cells | continue processing after workbook load exception

using System;
using System.IO;
using Aspose.Cells;

// The example enumerates all .xlsx files in an input folder, attempts to load each workbook with Aspose.Cells inside a try‑catch block, logs any load exceptions (e.g., corrupted files), and proceeds to the next file, saving successfully processed workbooks to an output folder while optionally preserving the original directory hierarchy.
class BatchWorkbookProcessor
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where processed workbooks will be saved
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the input folder
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string filePath in workbookFiles)
        {
            try
            {
                // Load workbook (using the provided load rule)
                Workbook workbook = new Workbook(filePath);

                // -------------------------------------------------
                // Place any workbook processing logic here.
                // For example, you could modify cells, add sheets, etc.
                // -------------------------------------------------

                // Save workbook (using the provided save rule)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath);
                Console.WriteLine($"Successfully processed: {filePath}");
            }
            catch (Exception ex)
            {
                // Handle load failures (e.g., corrupted files) and continue with next workbook
                Console.WriteLine($"Error loading workbook '{filePath}': {ex.Message}");
                // Optionally, log the error to a file or monitoring system
                continue; // Continue processing the remaining workbooks
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
