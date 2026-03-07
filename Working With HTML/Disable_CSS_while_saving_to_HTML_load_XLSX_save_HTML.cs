using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLSX)
            string inputPath = "input.xlsx";

            // Path for the generated HTML file
            string outputPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable CSS generation – use only inline styles
            htmlOptions.DisableCss = true;

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML with inline styles only: {outputPath}");
        }
    }
}