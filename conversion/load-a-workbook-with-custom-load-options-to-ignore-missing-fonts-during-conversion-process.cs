// Title: C# – Convert Excel to PDF with Aspose.Cells, ignore missing fonts via LoadOptions
// Description: Shows how to configure LoadOptions with IndividualFontConfigs to substitute unavailable fonts (e.g., NonExistentFont → Arial), enable workbook default‑font checking in PdfSaveOptions, and convert an Excel workbook to PDF using ConversionUtility.
// Keywords: Aspose.Cells | LoadOptions | font substitution | ignore missing fonts | Excel to PDF | C# | .NET | PdfSaveOptions | ConversionUtility | IndividualFontConfigs
// Common Searches: Aspose.Cells replace missing font during conversion | C# load Excel with custom font config Aspose | Convert Excel to PDF ignoring unavailable fonts .NET | LoadOptions font substitute example Aspose.Cells | How to prevent font errors when converting spreadsheets to PDF
// Developer Intent: Load a workbook with custom LoadOptions that replace missing fonts and convert it to PDF.
// Use Cases: Render user‑uploaded spreadsheets on a server that lacks the original fonts. | Batch‑process large sets of Excel files into PDFs while ensuring consistent appearance. | Integrate a document‑generation pipeline that avoids conversion failures caused by absent fonts.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions to map any missing font to Arial before converting an Excel file to PDF. | Write a reusable method that accepts input and output paths, applies font substitution, and returns conversion status with error handling.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

// Shows how to configure LoadOptions with IndividualFontConfigs to substitute unavailable fonts (e.g., NonExistentFont → Arial), enable workbook default‑font checking in PdfSaveOptions, and convert an Excel workbook to PDF using ConversionUtility.
class Program
{
    static void Main()
    {
        // Source workbook path (any format supported by Aspose.Cells)
        string sourcePath = "input.xlsx";

        // Destination file path after conversion (e.g., PDF)
        string destPath = "output.pdf";

        // Create LoadOptions and configure font substitution.
        // This tells the loader to replace a missing font with a known one (Arial).
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontConfigs = new IndividualFontConfigs();
        loadOptions.FontConfigs.SetFontSubstitutes("NonExistentFont", new string[] { "Arial" });

        // Create PDF save options (you can choose another format if needed).
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // Ensure the default workbook font is considered during rendering.
        saveOptions.CheckWorkbookDefaultFont = true;

        // Convert the workbook using the custom load options.
        ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

        Console.WriteLine("Conversion completed successfully.");
    }
}
