using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable CSS custom properties to optimize the HTML output
            htmlOptions.EnableCssCustomProperties = true;

            // Path for the resulting HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved as HTML with CSS custom properties enabled: {outputPath}");
        }
    }
}