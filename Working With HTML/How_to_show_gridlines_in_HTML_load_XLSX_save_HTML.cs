using System;
using Aspose.Cells;

namespace AsposeCellsGridlinesHtmlDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure gridlines are visible in the worksheet (optional, but aligns with ExportGridLines)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.IsGridlinesVisible = true;

            // Configure HTML save options to export gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Show gridlines in the generated HTML
                ExportActiveWorksheetOnly = true      // Export only the active sheet (optional)
            };

            // Save the workbook as an HTML file with gridlines
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("HTML file saved with gridlines visible.");
        }
    }
}