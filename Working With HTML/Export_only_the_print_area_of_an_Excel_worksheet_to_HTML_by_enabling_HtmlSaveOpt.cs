using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class ExportPrintAreaToHtml
    {
        static void Main(string[] args)
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Set HTML save options to export only the defined print area
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportPrintAreaOnly = true;

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", options);
        }
    }
}