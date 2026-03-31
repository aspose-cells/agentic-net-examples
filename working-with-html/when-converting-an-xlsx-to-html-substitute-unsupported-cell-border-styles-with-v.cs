using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the generated HTML file
            string htmlPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // When true, Aspose.Cells will replace border styles that are not supported by browsers
                // with visually similar alternatives, ensuring the HTML looks close to the original Excel.
                ExportSimilarBorderStyle = true,

                // Optional: keep CSS separate (default) and collapse table borders for cleaner output
                IsBorderCollapsed = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Excel file '{sourcePath}' has been converted to HTML at '{htmlPath}'.");
        }
    }
}