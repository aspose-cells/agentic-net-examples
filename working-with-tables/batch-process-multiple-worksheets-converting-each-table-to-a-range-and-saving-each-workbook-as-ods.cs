// Title: Batch convert Excel workbooks to ODS while removing all tables using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, deletes every ListObject from all worksheets while preserving cell values, and saves the result as an .ods file in a specified output directory. | Enhance the batch conversion script to log the names of each removed table to a text file and to gracefully skip files that contain no tables. | Create a reusable C# method that takes a Workbook object, removes all tables from its worksheets, and returns the modified workbook ready for saving as ODS with Aspose.Cells.
// Common Searches: aspnet remove all tables from Excel workbook and save as ODS | c# batch process multiple .xlsx files to .ods with Aspose.Cells | how to delete ListObject but keep data using Aspose.Cells | convert Excel files to OpenDocument Spreadsheet while stripping tables in .NET
// Tags: batch workbook conversion to ODS Aspose.Cells | remove ListObject tables while preserving data | iterate worksheets delete tables C# | convert .xlsx files to OpenDocument Spreadsheet | Aspose.Cells table removal automation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program iterates through all .xlsx files in a given input folder, loads each workbook with Aspose.Cells, removes every ListObject (Excel table) from each worksheet while leaving the underlying cell data intact, and saves the cleaned workbook as an ODS file in an output folder.
class BatchTableToRangeConverter
{
    static void Main()
    {
        // Folder containing source Excel workbooks
        string inputFolder = @"C:\InputWorkbooks";
        // Folder to save converted ODS workbooks
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure input and output directories exist
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Remove each table (ListObject) while preserving its data as a normal range
                    // Iterate backwards because we are modifying the collection
                    for (int i = sheet.ListObjects.Count - 1; i >= 0; i--)
                    {
                        ListObject table = sheet.ListObjects[i];
                        // Remove the table; the underlying cells remain unchanged
                        sheet.ListObjects.RemoveAt(i);
                    }
                }

                // Build output file path with .ods extension
                string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".ods";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the workbook as ODS format
                workbook.Save(outputPath, SaveFormat.Ods);

                Console.WriteLine($"Converted: {Path.GetFileName(filePath)} -> {outputFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
