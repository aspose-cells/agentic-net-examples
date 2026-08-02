// Title: Batch convert Excel tables to ranges and save as ODS – Aspose.Cells C#
// Description: C# utility scans a folder of .xlsx files, converts every ListObject table on each worksheet to a normal range, applies OdsSaveOptions (e.g., ignore pivots), and saves each workbook as an .ods file using Aspose.Cells.
// Keywords: Aspose.Cells C# batch conversion | convert Excel tables to ranges | save workbook as ODS | ListObject ConvertToRange | OdsSaveOptions ignore pivot tables | process multiple Excel files | automate Excel to OpenDocument | C# Excel to ODS utility
// Common Searches: batch convert Excel tables to ranges C# | Aspose.Cells save as ODS multiple files | convert ListObject to range loop Aspose | ignore pivot tables when saving ODS Aspose.Cells | C# script to convert .xlsx to .ods folder
// Developer Intent: Automatically transform all tables in each worksheet of many Excel workbooks into ranges and export the modified workbooks as ODS files.
// Use Cases: Migrate legacy Excel reports to OpenDocument format while flattening table structures for downstream systems. | Prepare a large set of financial spreadsheets for platforms that only accept .ods, ensuring compatibility by removing ListObject tables. | Automate cleanup of generated Excel files—convert tables to ranges and produce ODS versions in a single batch operation.
// AI Prompts: Write C# code that iterates through every worksheet in a workbook, converts each ListObject to a range, and saves the file as ODS with custom OdsSaveOptions. | Show how to add robust logging and error handling to a batch process that converts .xlsx files to .ods using Aspose.Cells. | Explain how to extend the sample to delete empty worksheets after table conversion before saving the workbook as ODS.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

// C# utility scans a folder of .xlsx files, converts every ListObject table on each worksheet to a normal range, applies OdsSaveOptions (e.g., ignore pivots), and saves each workbook as an .ods file using Aspose.Cells.
class BatchTableToRangeToOds
{
    static void Main()
    {
        // Folder containing source Excel workbooks
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where ODS files will be saved
        string outputFolder = @"C:\OutputOds";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string sourcePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Verify the source file exists (redundant but safe)
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"File not found: {sourcePath}");
                continue;
            }

            try
            {
                // Load the workbook from the source file
                Workbook workbook = new Workbook(sourcePath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Convert every table (ListObject) on the worksheet to a normal range
                    for (int i = 0; i < sheet.ListObjects.Count; i++)
                    {
                        ListObject table = sheet.ListObjects[i];
                        table.ConvertToRange(); // Uses ListObject.ConvertToRange method
                    }
                }

                // Configure ODS save options (optional settings)
                OdsSaveOptions odsOptions = new OdsSaveOptions
                {
                    IgnorePivotTables = true // Example: ignore pivot tables when saving
                };

                // Build the output file path with .ods extension
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                string outputPath = Path.Combine(outputFolder, fileNameWithoutExt + ".ods");

                // Save the modified workbook as ODS using the specified options
                workbook.Save(outputPath, odsOptions);

                Console.WriteLine($"Converted: {sourcePath} -> {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any errors for the current file and continue processing others
                Console.WriteLine($"Error processing file '{sourcePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing of tables to ranges and ODS conversion completed.");
    }
}
