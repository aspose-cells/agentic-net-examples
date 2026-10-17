// Title: How to batch set the workbook language to British English (en‑GB) for all Excel files in a folder using Aspose.Cells for .NET
// Common Searches: C# batch update Excel workbook language to en-GB with Aspose.Cells | Set CultureInfo for multiple .xlsx files using Aspose.Cells .NET | Iterate through a directory and change workbook Settings.CultureInfo to British English | Preserve original Excel format while updating document language in C# | Automate language property change for .xls, .xlsx, .xlsm files using Aspose.Cells
// Tags: batch workbook culture update Aspose.Cells | set workbook language en-GB C# | iterate Excel directory Aspose.Cells | preserve original file format Aspose.Cells | apply Settings.CultureInfo to multiple workbooks

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

// The example scans a specified folder, loads each .xls, .xlsx, or .xlsm workbook with Aspose.Cells, sets its Settings.CultureInfo to British English (en-GB), selects the appropriate SaveFormat based on the original extension, saves the workbook back preserving its original format, and logs progress while handling any errors.
class Program
{
    static void Main(string[] args)
    {
        // Determine the folder containing the workbooks.
        string folderPath = args.Length > 0 ? args[0] : @"C:\Workbooks";

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Retrieve all files in the folder (filter later for supported extensions).
        string[] allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in allFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();

            // Process only supported Excel formats.
            if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                continue;

            // Ensure the file exists before attempting to load.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Set the workbook culture to British English.
                workbook.Settings.CultureInfo = new CultureInfo("en-GB");

                // Determine the appropriate save format based on the original extension.
                SaveFormat format = extension switch
                {
                    ".xlsx" => SaveFormat.Xlsx,
                    ".xlsm" => SaveFormat.Xlsm,
                    ".xls" => SaveFormat.Excel97To2003,
                    _ => SaveFormat.Xlsx
                };

                // Save the workbook back to the same file, preserving its format.
                workbook.Save(filePath, format);
                Console.WriteLine($"Updated language for: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                // Log the error and continue processing other files.
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("All applicable workbooks have been processed.");
    }
}
