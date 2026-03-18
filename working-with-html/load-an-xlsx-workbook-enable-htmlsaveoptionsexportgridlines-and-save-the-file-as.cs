using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Optionally make gridlines visible in the worksheet (not required for ExportGridLines)
        workbook.Worksheets[0].IsGridlinesVisible = true;

        // Create HTML save options and enable grid line export
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = true
        };

        // Save the workbook as an HTML file with grid lines rendered
        workbook.Save("output.html", htmlOptions);
    }
}