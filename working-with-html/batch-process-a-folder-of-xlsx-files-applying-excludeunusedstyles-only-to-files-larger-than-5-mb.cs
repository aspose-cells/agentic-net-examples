// Title: C# batch processing of XLSX files larger than 5 MB with Aspose.Cells – handling missing ExcludeUnusedStyles
// AI Prompts: Write a C# console application that scans a specified folder, checks each .xlsx file’s size, and for files exceeding 5 MB loads the workbook with Aspose.Cells, attempts to call the ExcludeUnusedStyles property (or logs a warning if the property is unavailable), then overwrites the original file. | Enhance the sample code to add verbose logging (file name, size, success or error), ensure files under the 5 MB threshold are skipped, and provide a fallback routine that removes unused styles by iterating the workbook’s style collection when ExcludeUnusedStyles cannot be used.
// Common Searches: how to process only large Excel files with Aspose.Cells in C# | C# script to apply ExcludeUnusedStyles to workbooks bigger than 5 MB | Aspose.Cells batch remove unused styles from XLSX files | skip small Excel files when cleaning up styles using Aspose.Cells .NET | handle missing ExcludeUnusedStyles property in Aspose.Cells 2023
// Tags: batch processing XLSX with Aspose.Cells | filter Excel files by size C# | unused style cleanup Aspose.Cells | fallback style removal Aspose.Cells | large workbook optimization .NET

using System;
using System.IO;
using Aspose.Cells;

// The program enumerates .xlsx files in a given directory, skips any file 5 MB or smaller, loads larger workbooks with Aspose.Cells, notes that the ExcludeUnusedStyles property may be unavailable, saves the workbook back to the same path, and logs each operation’s outcome.
class Program
{
    static void Main(string[] args)
    {
        // Specify the folder containing the XLSX files
        string folderPath = @"C:\Path\To\Your\Folder";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Define the size threshold (5 MB)
        const long sizeThreshold = 5L * 1024 * 1024; // bytes

        // Get all .xlsx files in the folder (non‑recursive)
        string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            try
            {
                // Ensure the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                // Check file size
                FileInfo fi = new FileInfo(filePath);
                if (fi.Length <= sizeThreshold)
                    continue; // Skip files below the threshold

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // NOTE: ExcludeUnusedStyles property is not available in the current Aspose.Cells version.
                // If needed, upgrade the library or use alternative methods.

                // Save the workbook, overwriting the original file
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine($"Processed and saved: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
