using System;
using Aspose.Cells;

class ExportExcelToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set custom column widths (in characters)
        worksheet.Cells.SetColumnWidth(0, 20); // Column A
        worksheet.Cells.SetColumnWidth(1, 30); // Column B

        // Add sample data to demonstrate column width preservation
        worksheet.Cells["A1"].PutValue("Short");
        worksheet.Cells["B1"].PutValue("This is a longer text that should respect the column width when exported to HTML.");

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Export column widths as fixed CSS width values (default behavior)
        // Setting WidthScalable to false ensures pixel‑based widths are used.
        saveOptions.WidthScalable = false;

        // Keep the original Excel behavior for overflowing data
        saveOptions.FormatDataIgnoreColumnWidth = false;

        // Save the workbook as an HTML file with the specified options
        workbook.Save("ExportedWithColumnWidths.html", saveOptions);
    }
}