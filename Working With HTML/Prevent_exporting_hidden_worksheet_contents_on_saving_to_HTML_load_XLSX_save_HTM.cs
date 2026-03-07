using System;
using Aspose.Cells;

class PreventHiddenWorksheetExport
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to exclude hidden worksheets
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false   // Do not export hidden worksheet content
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}