using System;
using Aspose.Cells;

namespace ExportPrintAreaOnlyDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Optionally set a print area on the first worksheet
            // (If the workbook already has a print area, this step can be omitted)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PrintArea = "B2:F10";

            // Configure HTML save options to export only the defined print area
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
    }
}