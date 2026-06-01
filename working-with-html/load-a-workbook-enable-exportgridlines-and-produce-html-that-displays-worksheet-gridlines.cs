using System;
using Aspose.Cells;

namespace AsposeCellsExportGridLines
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific worksheet as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure that gridlines are visible in the worksheet
            worksheet.IsGridlinesVisible = true;

            // Create HTML save options and enable gridline export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Export the gridlines to HTML
                ExportActiveWorksheetOnly = true      // Optional: export only the active sheet
            };

            // Path for the generated HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with gridlines exported.");
        }
    }
}