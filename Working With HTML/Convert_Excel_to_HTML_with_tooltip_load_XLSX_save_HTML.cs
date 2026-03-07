using System;
using Aspose.Cells;

namespace ExcelToHtmlWithTooltip
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Path where the resulting HTML file will be saved
            string outputPath = "output.html";

            // Load the workbook from the specified Excel file
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options and enable tooltip generation
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.AddTooltipText = true; // Show tooltip when cell content is truncated

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Excel file '{sourcePath}' has been converted to HTML with tooltips at '{outputPath}'.");
        }
    }
}