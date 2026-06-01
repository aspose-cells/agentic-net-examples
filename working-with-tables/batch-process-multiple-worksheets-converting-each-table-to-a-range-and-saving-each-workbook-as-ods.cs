using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

namespace BatchTableToRangeToOds
{
    class Program
    {
        static void Main()
        {
            // Folder containing the source Excel workbooks
            string sourceFolder = @"C:\InputWorkbooks";

            // Folder where the ODS files will be saved
            string outputFolder = @"C:\OutputOds";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the source folder
            foreach (string excelFile in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Load the workbook using Aspose.Cells (create/load rule)
                Workbook workbook = new Workbook(excelFile);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all tables (ListObjects) in the worksheet
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        // Convert the table to a normal range (use provided ConvertToRange method)
                        table.ConvertToRange();
                    }
                }

                // Prepare ODS save options (optional: ignore pivot tables)
                OdsSaveOptions odsOptions = new OdsSaveOptions
                {
                    IgnorePivotTables = true
                };

                // Build the output file path with .ods extension
                string odsFileName = Path.GetFileNameWithoutExtension(excelFile) + ".ods";
                string odsPath = Path.Combine(outputFolder, odsFileName);

                // Save the modified workbook as ODS (save rule)
                workbook.Save(odsPath, odsOptions);

                Console.WriteLine($"Converted '{excelFile}' to ODS: '{odsPath}'");
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}