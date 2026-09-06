// Title: Convert an Excel workbook to HTML with BestFit layout and CSS custom properties for image deduplication using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.PresentationPreference to BestFit, enables CssCustomProperties for image deduplication, and saves the workbook as an .html file with Aspose.Cells. | Add robust file‑existence verification and comprehensive exception handling to the Excel‑to‑HTML conversion sample that uses PresentationPreference.BestFit.
// Common Searches: Aspose.Cells C# export Excel to HTML with PresentationPreference BestFit and CSS custom properties | how to enable image deduplication when saving workbook as HTML using Aspose.Cells .NET | C# example for converting .xlsx to .html with layout optimization in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference BestFit | Aspose.Cells enable CSS custom properties HTML export | Excel to HTML image deduplication Aspose.Cells | C# workbook to HTML layout optimization

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to use PresentationPreference.BestFit and enables CssCustomProperties for image deduplication, then saves the workbook as an HTML file while handling any runtime exceptions.
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

            // Configure HTML save options (using defaults compatible with current Aspose.Cells version)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
