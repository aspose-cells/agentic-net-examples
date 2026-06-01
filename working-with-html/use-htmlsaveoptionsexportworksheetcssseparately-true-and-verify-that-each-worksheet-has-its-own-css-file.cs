using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportWorksheetCssSeparatelyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with multiple worksheets
            Workbook workbook = new Workbook();

            // Add sample data to the default worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Data in Sheet 1");

            // Add a second worksheet and populate it
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["B2"].PutValue("Data in Sheet 2");

            // Add a third worksheet and populate it
            Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
            sheet3.Cells["C3"].PutValue("Data in Sheet 3");

            // Configure HTML save options to export CSS for each worksheet separately
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true,
                // Ensure the output directory is created automatically if it does not exist
                CreateDirectory = true
            };

            // Define the output folder and HTML file name
            string outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AsposeHtmlExport");
            string htmlFilePath = Path.Combine(outputFolder, "Workbook.html");

            // Save the workbook as HTML
            workbook.Save(htmlFilePath, saveOptions);

            // Verify that each worksheet has its own CSS file
            // Aspose.Cells creates CSS files named like "sheet001.css", "sheet002.css", etc.
            string[] cssFiles = Directory.GetFiles(outputFolder, "sheet*.css");

            Console.WriteLine($"Total worksheets in workbook: {workbook.Worksheets.Count}");
            Console.WriteLine($"CSS files found: {cssFiles.Length}");

            bool allSheetsHaveCss = cssFiles.Length == workbook.Worksheets.Count;

            if (allSheetsHaveCss)
            {
                Console.WriteLine("Verification succeeded: each worksheet has its own CSS file.");
                foreach (string cssFile in cssFiles)
                {
                    Console.WriteLine($" - {Path.GetFileName(cssFile)}");
                }
            }
            else
            {
                Console.WriteLine("Verification failed: number of CSS files does not match number of worksheets.");
            }
        }
    }
}