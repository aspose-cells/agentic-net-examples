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

            // Path for the generated HTML file
            string outputPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable external CSS; use only inline styles
            htmlOptions.DisableCss = true;

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine("Workbook has been exported to HTML with CSS disabled.");
        }
    }
}