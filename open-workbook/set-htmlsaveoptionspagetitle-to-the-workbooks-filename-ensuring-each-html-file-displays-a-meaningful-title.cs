using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Author: Generated example demonstrating HtmlSaveOptions.PageTitle usage
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string excelPath = "sample.xlsx";

            // Load the workbook from file
            Workbook workbook = new Workbook(excelPath);

            // Determine a meaningful page title (e.g., the workbook file name without extension)
            string pageTitle = Path.GetFileNameWithoutExtension(excelPath);

            // Configure HTML save options and set the PageTitle property
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                PageTitle = pageTitle
            };

            // Define the output HTML file path
            string htmlPath = "sample.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, saveOptions);

            Console.WriteLine($"HTML file saved to '{htmlPath}' with page title '{pageTitle}'.");
        }
    }
}