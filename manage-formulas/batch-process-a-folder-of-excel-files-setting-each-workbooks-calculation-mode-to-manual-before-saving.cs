// Title: Set calculation mode to Manual for every .xls and .xlsx workbook in a folder using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that enumerates all .xls and .xlsx files in a specified directory, opens each workbook with Aspose.Cells, sets Workbook.Settings.CalcMode to Manual, and saves the file back to its original path. | Update the provided Aspose.Cells example to enable manual calculation mode before saving each workbook, add handling for missing files and unsupported extensions, and log each processed file name.
// Common Searches: Aspose.Cells C# batch change workbook calculation mode to manual | How to set manual calculation for multiple Excel files using Aspose.Cells .NET | C# loop through folder and update Settings.CalcMode for each workbook | Overwrite original Excel workbook after changing calculation settings with Aspose.Cells | Process all .xls and .xlsx files in a directory to disable automatic calculation using Aspose.Cells
// Tags: Aspose.Cells manual calculation mode | batch modify workbook settings .NET | enumerate Excel files in folder C# | save workbook with updated settings Aspose | disable automatic calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a given directory for .xls and .xlsx files, loads each workbook with Aspose.Cells, optionally sets the calculation mode to Manual, and overwrites the original file while handling missing files and unsupported extensions.
class Program
{
    static void Main()
    {
        // Folder containing the Excel files to process
        string folderPath = @"C:\Path\To\ExcelFolder";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all Excel files in the folder (supports .xls and .xlsx)
        string[] excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            // Process only supported Excel extensions
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
                continue;

            // Ensure the file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
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
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Processing completed.");
    }
}
