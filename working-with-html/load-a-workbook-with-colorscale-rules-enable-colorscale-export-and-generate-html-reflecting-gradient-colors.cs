// Title: C# Example: Export Excel ColorScale Conditional Formatting to HTML with Aspose.Cells
// Description: Loads an existing workbook (or creates a new one if missing), configures HtmlSaveOptions to keep all styles—including ColorScale gradients—by setting ExportWorksheetCSSSeparately to true and ExcludeUnusedStyles to false, and saves the file as HTML that mirrors the original conditional formatting.
// Keywords: Aspose.Cells | C# | .NET | ColorScale export | conditional formatting HTML | HtmlSaveOptions | ExportWorksheetCSSSeparately | ExcludeUnusedStyles | Excel to HTML conversion | gradient colors HTML
// Common Searches: Aspose.Cells export ColorScale to HTML C# | preserve Excel conditional formatting when converting to HTML | HtmlSaveOptions keep gradient colors | export workbook with conditional formatting as HTML | C# code sample for Excel ColorScale HTML export
// Developer Intent: Create an HTML file that retains Excel ColorScale gradient formatting using Aspose.Cells.
// Use Cases: Render heat‑map reports as web‑ready HTML pages without losing color gradients. | Automate conversion of Excel dashboards for email newsletters or intranet portals. | Generate a placeholder HTML report when the source Excel file is unavailable.
// AI Prompts: Show how to enable ColorScale export in Aspose.Cells HtmlSaveOptions and save as HTML. | Provide C# code that loads an Excel file with conditional formatting, preserves gradient colors, and writes the HTML output. | Explain the impact of ExportWorksheetCSSSeparately and ExcludeUnusedStyles on HTML representation of ColorScale rules.

using System;
using System.IO;
using Aspose.Cells;

namespace ColorScaleHtmlExport
{
    // Loads an existing workbook (or creates a new one if missing), configures HtmlSaveOptions to keep all styles—including ColorScale gradients—by setting ExportWorksheetCSSSeparately to true and ExcludeUnusedStyles to false, and saves the file as HTML that mirrors the original conditional formatting.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file that already contains ColorScale conditional formatting
            string inputPath = "ColorScaleTemplate.xlsx";

            Workbook workbook = null;

            try
            {
                // Load the workbook if the file exists; otherwise create a new workbook
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    Console.WriteLine("Creating a new workbook as a fallback.");
                    workbook = new Workbook(); // creates a workbook with a default worksheet
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export worksheet CSS separately to preserve conditional formatting styles (including ColorScale gradients)
                    ExportWorksheetCSSSeparately = true,

                    // Keep all styles in the HTML (useful for round‑trip scenarios)
                    ExcludeUnusedStyles = false
                };

                // Define the output HTML file path
                string outputPath = "ColorScaleOutput.html";

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook has been exported to HTML: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
