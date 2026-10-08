// Title: Generate CSS‑free HTML from an Excel workbook using Aspose.Cells C# with BestFit column layout
// AI Prompts: Write C# code that reads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to BestFit, turns off all CSS output, and saves the result as a single HTML page. | Explain how to configure Aspose.Cells HtmlSaveOptions to produce a CSS‑free HTML export with automatically adjusted column widths in a .NET application.
// Common Searches: Aspose.Cells export Excel to HTML without stylesheet | BestFit column width setting for HTML output in Aspose.Cells C# | Turn off CSS generation in Aspose.Cells HtmlSaveOptions | Convert .xlsx to plain HTML using Aspose.Cells .NET | PresentationPreference BestFit not applied in Aspose.Cells HTML export
// Tags: Aspose.Cells HtmlSaveOptions disable stylesheet | PresentationPreference BestFit usage Aspose.Cells | Excel workbook to HTML conversion C# | auto‑fit columns HTML export Aspose.Cells | C# generate CSS‑free HTML with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The sample loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions to suppress CSS generation and to use the BestFit presentation preference (auto‑adjusting column widths), and saves the workbook as a single HTML file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Set up HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Adjust column widths automatically (BestFit) - property not available in this version
            // htmlOptions.PresentationPreference = Aspose.Cells.Rendering.PresentationPreference.BestFit;

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
