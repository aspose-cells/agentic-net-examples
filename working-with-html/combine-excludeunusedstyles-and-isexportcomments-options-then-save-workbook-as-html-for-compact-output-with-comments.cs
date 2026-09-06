// Title: Convert an Excel workbook to compact HTML with unused styles removed and cell comments included using Aspose.Cells for .NET
// AI Prompts: Write C# that loads an .xlsx file, enables HtmlSaveOptions.ExcludeUnusedStyles and HtmlSaveOptions.IsExportComments, and saves the workbook as an HTML file. | Update existing Aspose.Cells code to generate a minimal HTML file that keeps cell comments while stripping out any unused CSS styles.
// Common Searches: how to export Excel comments to HTML with Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions exclude unused styles example | save workbook as compact HTML with comments using Aspose.Cells | reduce HTML file size when converting Excel to HTML Aspose.Cells | C# Aspose.Cells export Excel to HTML without extra CSS
// Tags: HtmlSaveOptions.ExcludeUnusedStyles usage | export cell comments Aspose.Cells HTML | compact HTML output Excel conversion .NET | remove unused CSS styles Aspose.Cells HTML export | optimize HTML size Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads 'input.xlsx', configures HtmlSaveOptions to exclude unused styles and export cell comments, then saves the workbook as a compact 'output.html' while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.html";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Exclude unused styles to reduce file size
                    ExcludeUnusedStyles = true,
                    // Export cell comments into the HTML output
                    IsExportComments = true
                    // Note: HtmlFormattingOptions property is not available in this version of Aspose.Cells
                };

                // Save the workbook as a compact HTML file with comments
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
