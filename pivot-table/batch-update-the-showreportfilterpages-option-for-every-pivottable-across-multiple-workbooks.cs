using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace BatchUpdatePivotReportFilterPages
{
    class Program
    {
        static void Main()
        {
            // Folder containing source workbooks
            string sourceFolder = @"C:\InputWorkbooks";
            // Folder where updated workbooks will be saved
            string outputFolder = @"C:\UpdatedWorkbooks";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (supports .xlsx and .xls)
            string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in workbookFiles)
            {
                // Load the workbook (create rule)
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all pivot tables in the worksheet
                    foreach (PivotTable pivotTable in sheet.PivotTables)
                    {
                        // For each page field, show its report filter page
                        foreach (PivotField pageField in pivotTable.PageFields)
                        {
                            // ShowReportFilterPage method (rule)
                            pivotTable.ShowReportFilterPage(pageField);
                        }
                    }
                }

                // Build output file path (preserve original name)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the modified workbook (save rule)
                workbook.Save(outputPath);
            }

            Console.WriteLine("Batch update completed.");
        }
    }
}