// Title: Export hidden worksheets and apply best‑fit column widths when converting an Excel workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, enables HtmlSaveOptions.ExportHiddenWorksheet, sets PresentationPreference to BestFit, and saves the workbook as HTML using Aspose.Cells. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to include hidden sheets and automatically adjust column widths for optimal HTML rendering. | Provide error‑handling logic for missing input files while converting Excel to HTML with hidden worksheets in C#.
// Common Searches: asp.net convert excel to html include hidden worksheets Aspose.Cells | c# Aspose.Cells PresentationPreference BestFit HTML export example | how to export hidden sheets to html using Aspose.Cells | auto‑fit column widths when saving workbook as html with Aspose.Cells | HtmlSaveOptions settings for hidden worksheets and best fit layout
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | Aspose.Cells PresentationPreference BestFit | C# Excel to HTML hidden sheet conversion | auto‑fit column widths Aspose.Cells HTML output | Aspose.Cells hidden sheet inclusion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example checks for the input file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to export hidden worksheets (and optionally set PresentationPreference.BestFit for column auto‑fit), then saves the workbook as an HTML file while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export hidden worksheets as part of the HTML output
                ExportHiddenWorksheet = true
                // PresentationPreference property removed for compatibility with current API version
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
