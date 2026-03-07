using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and disable exporting hidden worksheets
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}