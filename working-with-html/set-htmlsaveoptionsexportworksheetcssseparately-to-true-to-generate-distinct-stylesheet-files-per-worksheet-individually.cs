using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportWorksheetCssSeparatelyDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1.Cells["A1"].PutValue("Data in Sheet 1");

                // Add a second worksheet with its own data
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                sheet2.Cells["A1"].PutValue("Data in Sheet 2");

                // Configure HTML save options to export CSS for each worksheet separately
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportWorksheetCSSSeparately = true // Enable separate CSS files per worksheet
                };

                // Define output directory
                string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "HtmlExport");
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML; each worksheet will have its own CSS file
                string outputPath = Path.Combine(outputDir, "Workbook.html");
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"HTML saved to: {outputPath}");
                Console.WriteLine("Separate CSS files for each worksheet have been generated in the same folder.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorksheetCssSeparatelyDemo.Run();
        }
    }
}