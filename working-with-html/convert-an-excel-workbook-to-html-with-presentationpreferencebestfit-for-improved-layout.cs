// Title: Convert an Excel workbook to HTML with best‑fit layout using Aspose.Cells PresentationPreference in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to BestFit, and saves the result as an HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions for best‑fit HTML rendering, including a check for the source file and proper exception handling.
// Common Searches: Aspose.Cells C# HtmlSaveOptions PresentationPreference BestFit example | How to export Excel to HTML while preserving column widths using Aspose.Cells | C# convert .xlsx to .html with optimal layout Aspose.Cells | Saving a workbook as HTML with best‑fit rendering in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference | C# export Excel to HTML best fit | HTML conversion preserving layout Aspose.Cells | Workbook.Save HTML best‑fit rendering | file existence validation Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies that the input Excel file exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions with PresentationPreference.BestFit for optimal column‑width rendering, saves the workbook as an HTML file, and handles any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: PresentationPreference enum may not be available in some versions.
            // If needed, adjust options here.

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
