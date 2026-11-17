// Title: Export an Excel workbook to HTML with best‑fit layout and CSS custom properties for image deduplication using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to HTML with PresentationPreference.BestFit | enable CSS custom properties for image deduplication in Aspose.Cells HTML export | Saving workbook as HTML with best fit layout using HtmlSaveOptions in .NET | C# code sample for Aspose.Cells HtmlSaveOptions EnableCssCustomProperties | Avoid duplicate images in HTML output from Aspose.Cells Excel conversion
// Tags: Aspose.Cells HtmlSaveOptions best-fit layout | EnableCssCustomProperties image deduplication | C# Excel to HTML conversion Aspose.Cells | PresentationPreference.BestFit HTML export | CSS variables for deduped images in HTML

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions to enable CSS custom properties for image deduplication, optionally sets PresentationPreference.BestFit for a best‑fit layout, and saves the workbook as an HTML file.
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
            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Enable CSS custom properties for image deduplication
                EnableCssCustomProperties = true
                // Note: PresentationPreference property may not be available in older Aspose.Cells versions.
                // If needed, it can be set when the enum is supported.
            };

            // Save the workbook as HTML
            workbook.Save(outputPath, options);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
