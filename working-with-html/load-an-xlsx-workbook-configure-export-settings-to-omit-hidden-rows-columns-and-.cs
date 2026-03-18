using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML export options to exclude hidden worksheets, rows, and columns
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false,                     // Omit hidden worksheets
            HiddenColDisplayType = HtmlHiddenColDisplayType.Remove, // Remove hidden columns
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove  // Remove hidden rows
        };

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", saveOptions);
    }
}