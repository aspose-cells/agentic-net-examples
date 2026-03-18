using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string inputPath = "input.xlsx";

            // Path where the resulting HTML file will be saved
            string outputPath = "output.html";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Remove any styles that are not used in the workbook
            workbook.RemoveUnusedStyles();

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Explicitly set to exclude unused styles (default is true)
            htmlOptions.ExcludeUnusedStyles = true;

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook converted to HTML and saved to: {outputPath}");
        }
    }
}