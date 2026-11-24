// Title: Batch remove all ListObject tables from Excel worksheets and save each workbook as ODS using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells delete all tables in a workbook before exporting to ODS | How to iterate through worksheets and remove ListObjects in bulk with Aspose.Cells | Batch convert a folder of .xlsx files to .ods while stripping Excel tables using Aspose.Cells | Aspose.Cells example for converting Excel tables to ranges and saving as OpenDocument Spreadsheet | Programmatically remove Excel tables and save as ODS in .NET
// Tags: Aspose.Cells delete ListObjects | Excel to ODS batch conversion C# | remove tables from worksheets Aspose.Cells | save workbook as OpenDocument Spreadsheet | process multiple workbooks programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample scans a given input directory for .xlsx files, loads each workbook with Aspose.Cells, loops through every worksheet, removes all ListObject tables (optionally accessing their data ranges), and then saves the modified workbook as an .ods file in an output folder, handling missing files and runtime errors.
class BatchTableToRangeConverter
{
    static void Main()
    {
        // Folder containing source Excel workbooks
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where ODS files will be saved
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all .xlsx files in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string excelPath in excelFiles)
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate backwards because we will remove tables while iterating
                    for (int i = sheet.ListObjects.Count - 1; i >= 0; i--)
                    {
                        ListObject table = sheet.ListObjects[i];

                        // Get the data range occupied by the table (optional use)
                        Aspose.Cells.Range tableRange = table.DataRange;

                        // Remove the table definition; cell data remains intact
                        sheet.ListObjects.RemoveAt(i);

                        // Additional processing with tableRange can be added here if needed
                    }
                }

                // Build the output ODS file path (same name, different extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                string odsPath = Path.Combine(outputFolder, fileNameWithoutExt + ".ods");

                // Save the modified workbook as ODS
                workbook.Save(odsPath, SaveFormat.Ods);
                Console.WriteLine($"Converted: {excelPath} -> {odsPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
