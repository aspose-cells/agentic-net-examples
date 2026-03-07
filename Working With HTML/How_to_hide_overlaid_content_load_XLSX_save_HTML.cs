using System;
using Aspose.Cells;

class HideOverlaidContent
{
    static void Main()
    {
        // Load the XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to hide hidden worksheets, columns, rows, and invisible shapes
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false,                 // Do not export hidden worksheets
            HiddenColDisplayType = HtmlHiddenColDisplayType.Remove, // Remove hidden columns
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove, // Remove hidden rows
            IgnoreInvisibleShapes = true                  // Do not export invisible shapes
        };

        // Save the workbook as HTML with the configured options
        workbook.Save("output.html", saveOptions);
    }
}