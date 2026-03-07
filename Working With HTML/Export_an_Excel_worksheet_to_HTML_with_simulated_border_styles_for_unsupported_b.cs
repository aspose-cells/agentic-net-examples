using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            string inputPath = "input.xlsx"; // Path to the source Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to simulate unsupported border styles
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportSimilarBorderStyle = true // Enable simulation of unsupported borders
            };

            // Save the workbook as an HTML file with the specified options
            string outputPath = "output.html"; // Desired output HTML file path
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML with simulated borders: {outputPath}");
        }
    }
}