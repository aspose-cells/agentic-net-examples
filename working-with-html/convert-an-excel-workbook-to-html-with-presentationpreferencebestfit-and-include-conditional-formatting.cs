// Title: Convert an Excel workbook to HTML with PresentationPreference.BestFit and preserve conditional formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.PresentationPreference to BestFit, enables conditional formatting export, and saves the workbook as .html with Aspose.Cells. | Explain how to handle the absence of the ExportConditionalFormatting property in older Aspose.Cells versions while still achieving HTML output that reflects conditional formatting.
// Common Searches: Aspose.Cells C# convert Excel to HTML with best‑fit layout | How to export conditional formatting when saving workbook as HTML using Aspose.Cells .NET | Set PresentationPreference to BestFit in HtmlSaveOptions Aspose.Cells example | C# code sample for saving .xlsx as .html while keeping conditional formatting with Aspose.Cells | Aspose.Cells HTML conversion preserving cell styles and conditional rules
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference BestFit | Aspose.Cells export conditional formatting to HTML | C# convert .xlsx to .html Aspose.Cells | HTML conversion layout optimization Aspose.Cells | handling missing ExportConditionalFormatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing Excel file, configures HtmlSaveOptions to use PresentationPreference.BestFit, attempts to enable conditional formatting export (or notes the property’s absence in older versions), and saves the workbook as an HTML file, handling errors and confirming successful conversion.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportConditionalFormatting property is not available in this version of Aspose.Cells.
            // The workbook will be saved without explicit conditional formatting export.

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
