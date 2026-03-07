using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure the first worksheet has gridlines visible (optional, but aligns with ExportGridLines)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.IsGridlinesVisible = true;

        // Create HTML save options and enable gridline export
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = true,
            // Export only the active worksheet (optional, can be omitted if whole workbook is needed)
            ExportActiveWorksheetOnly = true
        };

        // Save the workbook as an HTML file with visible gridlines
        workbook.Save("output.html", htmlOptions);
    }
}