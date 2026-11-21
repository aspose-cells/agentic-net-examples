// Title: How to enable CSS custom properties and set TableCssId when exporting Excel to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# enable CSS variables in HtmlSaveOptions for HTML export | Set TableCssId in HtmlSaveOptions to style tables in generated HTML using Aspose.Cells | Export Excel workbook to HTML with custom CSS properties using Aspose.Cells .NET | Configure HtmlSaveOptions to use CSS custom properties and table ID in Aspose.Cells | C# sample for enabling EnableCssCustomProperties and TableCssId in Aspose.Cells HTML output
// Tags: Aspose.Cells HtmlSaveOptions EnableCssCustomProperties | Aspose.Cells HtmlSaveOptions TableCssId | C# export Excel to HTML with CSS variables | custom table styling Aspose.Cells HTML output | HTML save options configuration Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This example loads an Excel workbook, configures HtmlSaveOptions to enable CSS custom properties and assign a TableCssId, then saves the workbook as HTML, allowing flexible styling of tables with CSS variables.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Enable CSS custom properties (variables) in the generated HTML/CSS
                    EnableCssCustomProperties = true
                };

                // Save the workbook as HTML using the configured options
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
