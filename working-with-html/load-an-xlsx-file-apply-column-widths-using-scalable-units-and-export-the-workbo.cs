using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Apply custom column widths (in characters)
        // These widths will be exported using scalable units because of the option set below
        worksheet.Cells.SetColumnWidth(0, 20); // Column A
        worksheet.Cells.SetColumnWidth(1, 30); // Column B

        // Create HTML save options and enable scalable column width export
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.WidthScalable = true; // Use scalable units for column widths

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}