using System;
using Aspose.Cells;

class ExportPrintAreaToHtml
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Define the print area on the first worksheet (e.g., B2:F10)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.PageSetup.PrintArea = "B2:F10";

        // Configure HTML save options to export only the defined print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportPrintAreaOnly = true;

        // Save the workbook as HTML using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}