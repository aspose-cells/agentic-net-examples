using System;
using Aspose.Cells;

namespace ExportConditionalFormattingToHtml
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLSX)
            string inputPath = "input.xlsx";

            // Path where the resulting HTML file will be saved
            string outputPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export all data (including conditional formatting) to HTML
            htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook has been exported to HTML: {outputPath}");
        }
    }
}