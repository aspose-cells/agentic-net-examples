using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(sourcePath);

            // Ensure gridlines are visible in the worksheet (optional, but aligns with ExportGridLines)
            workbook.Worksheets[0].IsGridlinesVisible = true;

            // Create HTML save options and enable gridline export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Preserve gridlines in the HTML output
                ExportActiveWorksheetOnly = true      // Export only the active worksheet (adjust as needed)
            };

            // Save the workbook as HTML with the specified options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully exported to HTML with gridlines: {outputPath}");
        }
    }
}