using System;
using Aspose.Cells;

class ExportPrintAreaOnlyDemo
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to export only the defined print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportPrintAreaOnly = true;   // Export only the worksheet's print area

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}