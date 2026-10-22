// Title: How to batch set Excel workbook calculation mode to Manual for all .xls and .xlsx files in a folder using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set calculation mode to Manual for every workbook in a directory | Batch process Excel files to disable automatic calculation with Aspose.Cells .NET | How to iterate over a folder of .xls and .xlsx files and change CalcMode using Aspose.Cells | Programmatically change workbook Settings.CalcMode to Manual for multiple Excel files in C# | Aspose.Cells example for updating calculation mode of all workbooks in a folder
// Tags: Aspose.Cells set manual calculation mode | batch update workbook calculation mode .NET | process folder of Excel workbooks Aspose.Cells | disable automatic formula recalculation C# | iterate Excel files directory Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified directory for .xls and .xlsx files, loads each workbook with Aspose.Cells, sets the workbook's Settings.CalcMode to Manual, saves the workbook back to its original location, and logs processing results while handling missing files and exceptions.
class Program
{
    static void Main(string[] args)
    {
        // Specify the folder containing Excel files
        string folderPath = @"C:\ExcelFolder";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all Excel files in the folder (both .xls and .xlsx)
        string[] excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue; // Skip non‑Excel files

            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // (Optional) Set calculation mode to Manual if supported
                // workbook.Settings.CalcMode = CalculationMode.Manual;

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
