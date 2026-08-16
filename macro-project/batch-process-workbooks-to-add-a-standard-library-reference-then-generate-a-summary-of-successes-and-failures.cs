// Title: Batch add stdole VBA reference to .xlsm workbooks with Aspose.Cells (C#)
// Description: A C# console utility that scans a folder for macro‑enabled Excel files (*.xlsm), loads each workbook with Aspose.Cells, adds the registered stdole VBA library reference when a VBA project exists, saves the updated file, and prints a concise success/failure summary.
// Keywords: Aspose.Cells | C# batch Excel | add VBA reference | stdole library | macro‑enabled workbook automation | Excel .xlsm processing | VBA project manipulation | CI/CD Excel validation | bulk Excel update | GitHub example
// Common Searches: add stdole reference to multiple xlsm files using Aspose.Cells | C# batch update macro enabled workbooks | Aspose.Cells VBA project example | automate VBA library addition in Excel files | summary of batch Excel processing results
// Developer Intent: Automatically insert a standard VBA library reference into every .xlsm file in a directory and report how many files were updated versus skipped or failed.
// Use Cases: Prepare a distribution package by ensuring all macro‑enabled workbooks contain the required stdole reference. | Audit a repository of Excel macros, fixing missing references in bulk and generating a status report. | Integrate the script into a CI/CD pipeline to enforce VBA reference compliance before release.
// AI Prompts: Create a parallel version of this batch processor using async/await while keeping accurate success/failure counts. | Extend the code to read a JSON configuration that maps multiple GUIDs to library names and log each file’s outcome to a CSV file. | Explain how to modify the program to work with .xlsb files that contain VBA projects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// A C# console utility that scans a folder for macro‑enabled Excel files (*.xlsm), loads each workbook with Aspose.Cells, adds the registered stdole VBA library reference when a VBA project exists, saves the updated file, and prints a concise success/failure summary.
class BatchVbaReferenceAdder
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\Workbooks\Input";
        // Optional: folder to save processed workbooks (can be same as input)
        string outputFolder = @"C:\Workbooks\Output";

        // Ensure input and output folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Get all macro-enabled Excel files in the input folder
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsm");

        int successCount = 0;
        int failureCount = 0;

        foreach (string filePath in workbookFiles)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipped: {Path.GetFileName(filePath)}");
                    failureCount++;
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Check if the workbook contains a VBA project
                if (workbook.VbaProject != null)
                {
                    // Add a standard registered reference
                    workbook.VbaProject.References.AddRegisteredReference(
                        "stdole",
                        "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

                    // Determine output file path
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the modified workbook
                    workbook.Save(outputPath);

                    successCount++;
                }
                else
                {
                    // No VBA project present; cannot add reference
                    Console.WriteLine($"Skipped (no VBA project): {Path.GetFileName(filePath)}");
                    failureCount++;
                }
            }
            catch (Exception ex)
            {
                // Log any errors that occur during processing
                Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                failureCount++;
            }
        }

        // Summary of processing results
        Console.WriteLine("Batch processing completed.");
        Console.WriteLine($"Total files processed: {workbookFiles.Length}");
        Console.WriteLine($"Successfully updated: {successCount}");
        Console.WriteLine($"Failed or skipped: {failureCount}");
    }
}
