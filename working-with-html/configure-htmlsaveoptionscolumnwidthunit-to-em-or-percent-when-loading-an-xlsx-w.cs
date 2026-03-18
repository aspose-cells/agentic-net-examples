using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook.
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Example: set column width to be expressed in pixels (optional).
        // htmlOptions.ColumnWidthInPixels = true;

        // Export the workbook to HTML using the configured options.
        workbook.Save("output.html", htmlOptions);
    }
}