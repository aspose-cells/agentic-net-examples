// Title: Export an Excel workbook to separate HTML files per worksheet while keeping inter‑sheet hyperlinks using Aspose.Cells C# IFilePathProvider
// AI Prompts: Write C# code that implements a custom file path provider to name each worksheet's HTML file and configures HtmlSaveOptions to export all sheets, preserving hyperlinks between them. | Show how to create a navigation page that links to the individual sheet HTML files when saving a workbook with Aspose.Cells. | Demonstrate loading a workbook, setting HtmlSaveOptions (ExportActiveWorksheetOnly = false) together with a custom naming provider, and saving the workbook as multiple HTML pages with functional inter‑sheet links.
// Common Searches: Aspose.Cells export each worksheet to its own HTML file C# | keep hyperlinks between sheets when saving Excel as HTML using Aspose.Cells | how to use IFilePathProvider for multi‑sheet HTML export in Aspose.Cells | generate index.html with navigation to sheet HTML pages Aspose.Cells | C# Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example
// Tags: Aspose.Cells multi‑sheet HTML export | custom sheet HTML naming provider | maintain inter‑sheet hyperlinks Aspose.Cells | HtmlSaveOptions for exporting all worksheets | navigation page creation for worksheet HTML

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // Custom file path provider that determines the HTML file name for each worksheet.
    // The example loads InputWorkbook.xlsx, sets HtmlSaveOptions to export all worksheets, assigns a custom IFilePathProvider that returns "{sheetName}.html" for each sheet, and saves the workbook as index.html. This produces separate HTML files for each worksheet and preserves inter‑sheet hyperlinks.
    class SheetFilePathProvider : IFilePathProvider
    {
        // This method is called for every worksheet during HTML export.
        // It receives the worksheet name and returns the desired file path.
        public string GetFullName(string sheetName)
        {
            // Create a simple file name based on the worksheet name.
            // You can customize the path or naming convention as needed.
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";

                // Ensure the input workbook exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML export options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export all worksheets (required for inter‑worksheet link preservation).
                    ExportActiveWorksheetOnly = false,

                    // Use the custom file path provider to generate separate HTML files.
                    FilePathProvider = new SheetFilePathProvider()
                };

                // Save the workbook to HTML. The main file (index.html) will contain navigation and references to the sheet files.
                workbook.Save("index.html", htmlOptions);
                Console.WriteLine("Workbook successfully exported to HTML.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
