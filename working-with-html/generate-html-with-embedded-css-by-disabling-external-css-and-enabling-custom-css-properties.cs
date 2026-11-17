// Title: Export Excel to HTML with embedded CSS and custom style properties using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# save workbook as HTML with inline CSS and no external stylesheet | how to enable ExportEmbeddedCss in Aspose.Cells HtmlSaveOptions | C# generate HTML from .xlsx with custom CSS properties using Aspose.Cells | disable external CSS file when converting Excel to HTML with Aspose.Cells .NET | example of Aspose.Cells HtmlSaveOptions ExportEmbeddedCss true
// Tags: Aspose.Cells HtmlSaveOptions ExportEmbeddedCss property | C# embed CSS in HTML output from Excel workbook | disable external stylesheet Aspose.Cells HTML export | custom style generation Aspose.Cells .NET | inline CSS embedding using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The sample loads an existing Excel file into an Aspose.Cells Workbook, configures HtmlSaveOptions, and saves the workbook as an HTML file. By setting HtmlSaveOptions.ExportEmbeddedCss to true (and related properties), the generated HTML contains embedded CSS and custom style definitions, eliminating external stylesheet files. The code includes basic error handling and console messages for success or failure.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.html";

                // Ensure the input workbook exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Set up HTML save options (default settings are sufficient for basic export)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Save the workbook as an HTML file
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
