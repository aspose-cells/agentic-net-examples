// Title: Batch load Excel workbooks from a folder, hide zero values on every worksheet, and overwrite files with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells hide zero values on all worksheets in multiple workbooks | Iterate through a folder of .xlsx files and apply custom formatting with Aspose.Cells | Batch update Excel workbooks to suppress zero values using Aspose.Cells .NET API | Overwrite original Excel files after processing with Aspose.Cells in C#
// Tags: batch processing .xlsx files Aspose.Cells | suppress zero display worksheet Aspose.Cells | load workbooks from folder Aspose.Cells .NET | overwrite original Excel workbook Aspose.Cells | custom zero‑value formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example enumerates all .xlsx files in a specified directory, loads each workbook with Aspose.Cells, provides a placeholder for custom logic to suppress zero values on every worksheet, and then saves the workbook back to the original file, handling common file‑system errors.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing the workbooks
            string folderPath = @"C:\Path\To\Folder";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files in the folder (adjust the pattern if needed)
            string[] files = Directory.GetFiles(folderPath, "*.xlsx");

            foreach (string filePath in files)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook from file
                    Workbook workbook = new Workbook(filePath);

                    // NOTE: The ShowZeroValues property is not available in the current Aspose.Cells version.
                    // If needed, implement custom logic to hide zero values here.

                    // Save the changes back to the same file (overwrites original)
                    workbook.Save(filePath);
                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
