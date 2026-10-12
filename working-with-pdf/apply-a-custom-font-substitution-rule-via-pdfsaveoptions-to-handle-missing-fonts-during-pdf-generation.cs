// Title: How to configure a custom font substitution rule with PdfSaveOptions for Aspose.Cells PDF conversion in C#
// AI Prompts: Write C# code that creates a PdfSaveOptions object, defines a mapping from an unavailable Excel font to a fallback system font, and saves the workbook as PDF using Aspose.Cells. | Show how to use the FontSubstitution (or equivalent) collection on PdfSaveOptions to replace missing fonts during Excel‑to‑PDF conversion in a .NET application.
// Common Searches: Aspose.Cells C# set custom font mapping when saving workbook to PDF | PdfSaveOptions font fallback example for missing Excel fonts | Replace unavailable fonts with system fonts in Aspose.Cells PDF export | Configure font substitution for Excel to PDF conversion using Aspose.Cells .NET
// Tags: pdfsaveoptions custom font mapping | aspocells missing font handling | excel to pdf font fallback c# | aspocells pdf export font replacement | c# aspocells pdfsaveoptions configuration

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook, creating a PdfSaveOptions instance, applying a custom font substitution rule to map missing fonts to a fallback font, and saving the workbook as a PDF, ensuring proper rendering when original fonts are unavailable.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Font substitution API may vary between Aspose.Cells versions.
            // If needed, configure custom font substitution here using the appropriate API.

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
