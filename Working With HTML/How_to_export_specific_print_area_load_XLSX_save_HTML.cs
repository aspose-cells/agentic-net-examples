using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportPrintAreaToHtml
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PrintArea = "B2:F10";

            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true,
                ExportGridLines = true
            };

            workbook.Save("output.html", htmlOptions);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPrintAreaToHtml.Run();
        }
    }
}