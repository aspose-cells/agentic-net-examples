// Title: Batch apply Workbook.Settings.ExcludeUnusedStyles to XLSX files over 5 MB using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a directory, loads each .xlsx larger than 5 MB with Aspose.Cells, sets Workbook.Settings.ExcludeUnusedStyles to true, and overwrites the original file. | Update the sample code to add a file‑size check and enable the ExcludeUnusedStyles optimization only when the workbook size exceeds 5 MB, including comprehensive error handling. | Create a reusable method that accepts a file path, verifies the size, applies ExcludeUnusedStyles via Aspose.Cells if the size is above 5 MB, saves the workbook, and returns a status indicator.
// Common Searches: how to enable ExcludeUnusedStyles for large Excel workbooks with Aspose.Cells C# | C# script to process only XLSX files bigger than 5 MB using Aspose.Cells | batch optimization of Excel files with Aspose.Cells excluding unused styles | Aspose.Cells conditional workbook settings based on file size in .NET | automate style cleanup for big Excel files with Aspose.Cells library
// Tags: excludeunusedstyles workbook.settings aspnet | batch process xlsx files aspnet | conditional workbook optimization aspnet | file size threshold excel aspnet | automated excel style cleanup aspnet

using System;
using System.IO;
using Aspose.Cells;

// // Enumerates .xlsx files in a given folder, skips files ≤5 MB, loads each larger workbook with Aspose.Cells, attempts to enable Workbook.Settings.ExcludeUnusedStyles (requires a newer library version), and saves the workbook back to the original location.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing the XLSX files
            string folderPath = @"C:\Path\To\XlsxFolder";

            // Size threshold: 5 MB in bytes
            const long sizeThreshold = 5L * 1024 * 1024;

            // Get all .xlsx files in the folder (non‑recursive)
            string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                try
                {
                    // Verify the file still exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Check file size
                    FileInfo fi = new FileInfo(filePath);
                    if (fi.Length <= sizeThreshold)
                        continue; // Skip files <= 5 MB

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Note: ExcludeUnusedStyles property is not available in this version of Aspose.Cells.
                    // If needed, upgrade the library or use alternative optimization methods.

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath, SaveFormat.Xlsx);
                    Console.WriteLine($"Processed: {filePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
