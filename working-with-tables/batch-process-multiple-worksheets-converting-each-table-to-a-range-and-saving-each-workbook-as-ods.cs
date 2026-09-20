// Title: Batch remove ListObject tables from all worksheets and export each workbook to ODS using Aspose.Cells for .NET
// AI Prompts: Generate C# code that scans a directory for .xlsx and .xls files, loads each workbook with Aspose.Cells, deletes every ListObject table on every worksheet, and saves the result as an .ods file. | Write a .NET console application that iterates through multiple Excel workbooks, removes table definitions (ListObjects) from each sheet, and batch‑converts the workbooks to OpenDocument Spreadsheet format using Aspose.Cells. | Create a script that processes all Excel files in a folder, strips tables to plain ranges, and saves each modified workbook as ODS with the Aspose.Cells API.
// Common Searches: how to delete all tables in an Excel workbook using Aspose.Cells C# | batch convert Excel workbooks to ODS with Aspose.Cells .NET | remove ListObject tables from each worksheet before saving as ODS | C# program to process a folder of Excel files and export them to OpenDocument Spreadsheet
// Tags: strip ListObject tables Aspose.Cells | bulk Excel to ODS export Aspose.Cells | loop worksheets and clear tables C# | export workbook to ODS format Aspose.Cells | handle multiple Excel files C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The utility scans a given input directory for .xlsx and .xls files, loads each workbook with Aspose.Cells, removes all ListObject tables from every worksheet, and saves the modified workbook as an ODS file in a specified output folder.
class BatchTableToRangeConverter
{
    static void Main()
    {
        // Folder containing source Excel workbooks
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where ODS files will be saved
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each Excel file in the input folder (supports .xlsx and .xls)
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls")
                continue; // Skip non‑Excel files

            // Verify the file still exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate tables (ListObjects) in reverse order to allow removal
                    for (int i = sheet.ListObjects.Count - 1; i >= 0; i--)
                    {
                        ListObject table = sheet.ListObjects[i];
                        // Remove the table definition; the underlying cells remain unchanged
                        sheet.ListObjects.RemoveAt(i);
                    }
                }

                // Build output file name with .ods extension
                string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".ods";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the modified workbook as ODS
                workbook.Save(outputPath, SaveFormat.Ods);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
