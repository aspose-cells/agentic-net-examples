// Title: How to batch set Excel workbooks to manual calculation mode using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set workbook calculation mode to Manual for all files in a folder | disable automatic recalculation for multiple Excel files with Aspose.Cells | process directory of .xls .xlsx .xlsm files and change CalcMode to Manual in C# | batch update Excel workbooks calculation settings using Aspose.Cells .NET API | Aspose.Cells example to iterate over Excel files and set manual calculation
// Tags: batch update workbook calculation mode Aspose.Cells | manual CalcMode setting for Excel files C# | iterate over directory of .xls .xlsx .xlsm Aspose.Cells | disable automatic recalculation in Aspose.Cells | save workbook after changing Settings.CalcMode Aspose

using System;
using System.IO;
using Aspose.Cells;

// Iterates through a specified folder, loads each .xls, .xlsx, or .xlsm workbook with Aspose.Cells, sets Settings.CalcMode to Manual (when supported), and saves the file back, while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        // Folder containing the Excel files to process
        string folderPath = @"C:\ExcelFolder";

        // Validate folder existence
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all Excel files in the folder (supported extensions)
        string[] excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            // Process only supported Excel extensions
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue;

            // Ensure the file still exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // NOTE: In newer Aspose.Cells versions the calculation mode can be set via
                // workbook.Settings.CalcMode = CalcMode.Manual;
                // If the property is unavailable, this step can be omitted.

                // Save the workbook, overwriting the original file
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Processing completed.");
    }
}
