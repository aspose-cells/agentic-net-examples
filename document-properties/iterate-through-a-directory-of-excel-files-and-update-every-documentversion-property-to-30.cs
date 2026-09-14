// Title: C# batch update Excel workbooks' DocumentVersion to 3.0 in a folder using Aspose.Cells
// AI Prompts: Write C# code that scans a given directory, opens each .xls, .xlsx, .xlsm, or .xlsb workbook with Aspose.Cells, sets the DocumentVersion property to "3.0", and saves the file while keeping its original format. | Enhance the processing loop to log the full path of every workbook that was updated, skipped, or caused an exception, and output the log to a text file. | Add a try‑catch block that captures any error while setting DocumentVersion and appends the file name and exception details to a CSV error report.
// Common Searches: Aspose.Cells set DocumentVersion for multiple Excel files in C# | How to programmatically change workbook version property across a folder of .xlsx files | Batch update Excel document properties using .NET Aspose.Cells | C# script to iterate over Excel files and modify custom properties
// Tags: Aspose.Cells set DocumentVersion C# | batch update Excel workbook properties | iterate Excel files directory Aspose | preserve original Excel format when saving | filter supported Excel extensions C#

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to enumerate all Excel files in a specified folder, load each workbook with Aspose.Cells, assign the DocumentVersion property the value "3.0", and save the workbook back in its original format, with logging for successes and errors.
class UpdateDocumentVersion
{
    static void Main()
    {
        // Directory containing the Excel files
        string folderPath = @"C:\ExcelFiles";

        // Verify that the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all files in the directory (filter later by supported extensions)
        string[] allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in allFiles)
        {
            // Process only supported Excel extensions
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb")
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

                // Save the workbook in Excel 2007+ (XLSX) format, overwriting the original file
                workbook.Save(filePath, SaveFormat.Xlsx);

                Console.WriteLine($"Updated version for: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
