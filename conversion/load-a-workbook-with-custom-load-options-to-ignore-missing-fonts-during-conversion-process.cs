// Title: Convert an Excel workbook to PDF while substituting missing fonts using Aspose.Cells LoadOptions in C#
// AI Prompts: Write C# code that loads an XLSX file with Aspose.Cells LoadOptions, sets a font substitute for a missing font, and saves the workbook as PDF. | Demonstrate how to configure IndividualFontConfigs in LoadOptions to replace a specific unavailable font with Arial during a workbook-to-PDF conversion. | Show how to use ConversionUtility.Convert together with LoadOptions and PdfSaveOptions to perform an Excel‑to‑PDF conversion that ignores missing fonts.
// Common Searches: Aspose.Cells C# ignore missing fonts when converting XLSX to PDF | How to set font substitution in LoadOptions for Excel to PDF conversion using Aspose.Cells | Replace unavailable font with Arial during workbook conversion Aspose.Cells .NET | LoadOptions FontConfigs SetFontSubstitutes example for PDF output
// Tags: Excel to PDF conversion with custom font mapping | LoadOptions IndividualFontConfigs usage | Aspose.Cells missing font handling | PdfSaveOptions configuration for workbook conversion | ConversionUtility custom load and save options

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// // Loads an XLSX workbook with Aspose.Cells LoadOptions, configures IndividualFontConfigs to replace a missing font (e.g., "MissingFont") with Arial, and converts the workbook to PDF using ConversionUtility and PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Source Excel file that may contain fonts not installed on the system
        string sourceFile = "input.xlsx";

        // Destination file after conversion (e.g., PDF)
        string destinationFile = "output.pdf";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure font substitution to handle missing fonts.
        // Any occurrence of a font named "MissingFont" will be replaced with "Arial".
        // Adjust the original font name as needed for your scenario.
        loadOptions.FontConfigs = new IndividualFontConfigs();
        loadOptions.FontConfigs.SetFontSubstitutes("MissingFont", new string[] { "Arial" });

        // Optional: set additional save options (here we use PDF format)
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Perform conversion using the custom load options.
        // This utilizes the provided ConversionUtility.Convert method that accepts
        // both LoadOptions and SaveOptions, complying with the lifecycle rules.
        ConversionUtility.Convert(sourceFile, loadOptions, destinationFile, saveOptions);
    }
}
