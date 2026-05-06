using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and disable exporting of hidden worksheets
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false // Do not include hidden sheets in the HTML output
        };

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", saveOptions);
    }
}