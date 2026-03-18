using System;
using Aspose.Cells;

namespace AsposeCellsHtmlConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook (replace with your file path)
            string inputPath = "sample.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Ensure unused styles are excluded to reduce HTML size
            // This is true by default, but set explicitly for clarity
            htmlOptions.ExcludeUnusedStyles = true;

            // Define output HTML file path
            string outputPath = "sample_converted.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook converted to HTML with unused styles omitted: {outputPath}");
        }
    }
}