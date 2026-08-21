// Title: C# Batch: Convert Excel Tables to Ranges and Export Workbooks as ODS with Aspose.Cells
// Description: Scans a directory for .xlsx files, loads each workbook using Aspose.Cells, iterates every worksheet and ListObject, converts each table to a normal range, and saves the result as an ODS file in a target folder.
// Keywords: Aspose.Cells C# | convert Excel table to range | batch process Excel workbooks | save workbook as ODS | ListObject to range | automate Excel to ODS conversion | C# Excel to OpenDocument
// Common Searches: C# Aspose.Cells convert all tables to ranges | batch export Excel files to ODS format | remove ListObjects from worksheets programmatically | convert Excel tables to ranges before saving as ODS | Aspose.Cells example for bulk workbook conversion
// Developer Intent: Automatically change every table in multiple Excel files into plain ranges and generate corresponding ODS files.
// Use Cases: Migrate legacy Excel reports with embedded tables to the open‑source ODS format for cross‑platform compatibility. | Pre‑process uploaded .xlsx documents in a web service, flatten table structures, and deliver ODS versions to downstream systems. | Integrate into a build or CI pipeline to ensure all workbooks are table‑free before publishing them as ODS assets.
// AI Prompts: Generate C# code that uses Aspose.Cells to iterate all worksheets, convert each ListObject to a range, and save the workbook as ODS. | Add comprehensive error handling and logging to the batch conversion script, skipping files that cannot be opened or saved. | Extend the example to delete empty worksheets after table conversion and then export the workbook to ODS.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Utility;

namespace BatchTableToRangeToOds
{
    // Scans a directory for .xlsx files, loads each workbook using Aspose.Cells, iterates every worksheet and ListObject, converts each table to a normal range, and saves the result as an ODS file in a target folder.
    class Program
    {
        static void Main()
        {
            // Folder containing source Excel workbooks
            string inputFolder = @"C:\InputWorkbooks";

            // Folder where ODS files will be saved
            string outputFolder = @"C:\OutputOds";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder (you can adjust the pattern as needed)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string excelPath in excelFiles)
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each table (ListObject) in the worksheet
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        // Convert the table to a normal range
                        table.ConvertToRange();
                    }
                }

                // Prepare the output ODS file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                string odsPath = Path.Combine(outputFolder, fileNameWithoutExt + ".ods");

                // Save the modified workbook as ODS
                workbook.Save(odsPath, SaveFormat.Ods);
            }

            Console.WriteLine("Batch processing completed successfully.");
        }
    }
}
