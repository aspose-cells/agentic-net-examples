using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportWorksheetCssSeparatelyDemo
    {
        public static void Run()
        {
            // Determine desktop folder
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string outputHtmlPath = Path.Combine(desktop, "sample.html");

            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DemoSheet";
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B2"].PutValue(12345);

            // Configure HTML save options to export CSS separately
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true,
                ExportActiveWorksheetOnly = true
            };

            // Save as HTML; a separate CSS file will be generated in the same folder
            workbook.Save(outputHtmlPath, saveOptions);

            Console.WriteLine($"HTML file saved to: {outputHtmlPath}");
            Console.WriteLine("A separate CSS file has been generated in the same directory.");
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