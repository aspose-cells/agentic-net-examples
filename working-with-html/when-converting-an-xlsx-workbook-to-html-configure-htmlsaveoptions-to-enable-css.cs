using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable CSS custom properties to optimize the HTML output
            htmlOptions.EnableCssCustomProperties = true;

            // Optionally set a page title for the generated HTML
            htmlOptions.PageTitle = "Exported HTML with CSS Custom Properties";

            // Save the workbook as an HTML file using the configured options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook has been saved to HTML with CSS custom properties enabled: {outputPath}");
        }
    }
}