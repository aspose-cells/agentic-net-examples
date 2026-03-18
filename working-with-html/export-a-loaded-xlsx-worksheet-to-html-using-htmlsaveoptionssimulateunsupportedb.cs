using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = Path.Combine(Environment.CurrentDirectory, "input.xlsx");
            Workbook workbook = new Workbook(inputPath);

            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            string outputPath = Path.Combine(Environment.CurrentDirectory, "output.html");
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully exported to HTML at: {outputPath}");
        }
    }
}