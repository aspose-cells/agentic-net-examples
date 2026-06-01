using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains hidden rows
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Make sure gridlines are visible in the worksheet
        worksheet.IsGridlinesVisible = true;

        // Configure HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            // Export the gridlines to the HTML output
            ExportGridLines = true,
            // Remove hidden rows so they do not affect gridline rendering
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove
        };

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", options);
    }
}