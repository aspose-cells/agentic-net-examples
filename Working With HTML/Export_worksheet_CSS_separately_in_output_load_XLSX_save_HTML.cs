using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportWorksheetCssSeparatelyDemo
    {
        public static void Run()
        {
            // Path to the source XLSX file on the desktop
            string inputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "input.xlsx");

            // If the input file does not exist, create a sample workbook
            if (!File.Exists(inputPath))
            {
                var sampleWb = new Workbook();
                var sheet = sampleWb.Worksheets[0];
                sheet.Name = "SampleSheet";
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B2"].PutValue(123);
                sampleWb.Save(inputPath);
            }

            // Load the workbook from the XLSX file
            var workbook = new Workbook(inputPath);

            // Configure HTML save options to export worksheet CSS separately
            var saveOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true
            };

            // Path for the generated HTML file on the desktop
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "output.html");

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("HTML file with separate worksheet CSS saved to: " + outputPath);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorksheetCssSeparatelyDemo.Run();
        }
    }
}