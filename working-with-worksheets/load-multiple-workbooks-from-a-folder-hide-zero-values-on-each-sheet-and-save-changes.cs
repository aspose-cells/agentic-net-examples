// Title: C# batch hide zero values in all worksheets of multiple Excel workbooks with Aspose.Cells
// Description: Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, sets Worksheet.DisplayZeros = false for every sheet, and overwrites the files. Includes folder validation, file‑existence checks, and exception handling.
// Keywords: Aspose.Cells | C# | .NET | DisplayZeros | hide zero values | batch process Excel | multiple workbooks | folder iteration | Excel automation | worksheet settings
// Common Searches: Aspose.Cells hide zeros in all worksheets | C# batch process Excel files to disable zero display | set DisplayZeros false for every sheet using Aspose.Cells | iterate over folder of .xlsx files and modify worksheets | bulk update Excel workbooks Aspose.Cells
// Developer Intent: Load each workbook in a directory, disable zero display on every worksheet, and save the changes back to the original files.
// Use Cases: Standardizing financial statements where zero amounts should not appear | Preparing large sets of report templates for client delivery | Nightly cleanup of Excel data exports to improve visual clarity | Automating data‑cleaning pipelines for BI dashboards
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through all .xlsx files in a given folder, set Worksheet.DisplayZeros = false for each sheet, and save the workbook. | Show how to add robust logging and continue‑on‑error handling when batch‑processing Excel workbooks with Aspose.Cells. | Explain how to make the folder path and file pattern configurable via appsettings in a .NET console app that hides zero values.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, sets Worksheet.DisplayZeros = false for every sheet, and overwrites the files. Includes folder validation, file‑existence checks, and exception handling.
    class Program
    {
        static void Main()
        {
            // Specify the folder containing the Excel workbooks
            string folderPath = @"C:\ExcelFiles";

            // Verify that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files in the folder (adjust the pattern as needed)
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

            foreach (string filePath in excelFiles)
            {
                // Ensure the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Hide zero values in each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.DisplayZeros = false;
                    }

                    // Save the modified workbook back to the same file
                    workbook.Save(filePath);
                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other files
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed for all workbooks in the folder.");
        }
    }
}
