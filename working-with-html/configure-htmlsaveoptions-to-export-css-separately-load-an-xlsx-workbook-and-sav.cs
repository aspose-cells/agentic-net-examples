using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    public class ExportHtmlWithSeparateCss
    {
        public static void Run()
        {
            string inputPath = Path.Combine(Environment.CurrentDirectory, "input.xlsx");
            Workbook workbook = new Workbook(inputPath);
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportWorksheetCSSSeparately = true;
            string outputPath = Path.Combine(Environment.CurrentDirectory, "output.html");
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved as HTML with separate CSS to: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportHtmlWithSeparateCss.Run();
        }
    }
}