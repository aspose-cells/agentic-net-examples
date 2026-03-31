using System;
using Aspose.Cells;

class ExportWorksheetToHtml
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Optional: keep column widths fixed (not scalable)
        htmlOptions.WidthScalable = false;

        // Export the workbook (or active worksheet) to HTML
        workbook.Save("output.html", htmlOptions);
    }
}