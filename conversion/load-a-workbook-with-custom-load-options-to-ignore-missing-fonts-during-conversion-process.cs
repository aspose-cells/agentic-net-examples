// Title: Convert Excel to PDF with Aspose.Cells .NET – ignore missing fonts & useless shapes using LoadOptions
// Description: Demonstrates how to load an XLSX workbook with LoadOptions that substitute missing fonts (e.g., replace any missing font with Arial) and ignore useless shapes, then convert it to PDF and optionally save it back to Excel using Aspose.Cells for .NET.
// Keywords: Aspose.Cells LoadOptions | missing font substitution | IgnoreUselessShapes | Excel to PDF conversion .NET | IndividualFontConfigs example | C# Aspose.Cells PDF conversion | font substitute Aspose.Cells
// Common Searches: Aspose.Cells replace missing fonts during conversion | ignore useless shapes when loading Excel workbook Aspose | LoadOptions FontConfigs C# Aspose.Cells | convert XLSX to PDF with custom font mapping Aspose | how to use ConversionUtility with LoadOptions Aspose.Cells
// Developer Intent: Load an Excel workbook with custom LoadOptions that map missing fonts to a substitute and skip unnecessary shapes, then convert the workbook to PDF (and optionally re‑save as XLSX).
// Use Cases: Prevent rendering errors by automatically substituting any unavailable font with a standard font like Arial during PDF generation. | Improve loading performance for large spreadsheets by enabling IgnoreUselessShapes, which skips non‑essential drawing objects. | Maintain consistent font handling by reusing the same LoadOptions when loading the workbook for further processing or saving back to Excel.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions to map several specific missing fonts to different substitute fonts and then converts the workbook to PDF. | Explain the impact of IgnoreUselessShapes on memory usage and conversion speed in Aspose.Cells, and show how to enable it. | Show how to apply a single LoadOptions instance for both PDF conversion via ConversionUtility and saving the workbook back to XLSX format.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to load an XLSX workbook with LoadOptions that substitute missing fonts (e.g., replace any missing font with Arial) and ignore useless shapes, then convert it to PDF and optionally save it back to Excel using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Paths for the source workbook and the converted file
        string sourcePath = "input.xlsx";
        string pdfPath = "output.pdf";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Initialize IndividualFontConfigs to handle missing fonts
        loadOptions.FontConfigs = new IndividualFontConfigs();

        // Define a generic substitute font for any missing font.
        // Here "MissingFont" is a placeholder; Aspose.Cells will use the substitute
        // when the original font cannot be found.
        loadOptions.FontConfigs.SetFontSubstitutes("MissingFont", new string[] { "Arial" });

        // Optional: ignore useless shapes to speed up loading
        loadOptions.IgnoreUselessShapes = true;

        // Convert the workbook to PDF using the custom LoadOptions.
        // This utilizes the provided ConversionUtility.Convert method that accepts
        // both LoadOptions and SaveOptions.
        ConversionUtility.Convert(
            sourcePath,
            loadOptions,
            pdfPath,
            new PdfSaveOptions()
        );

        // Additionally, demonstrate loading the workbook with the same options
        // and saving it back to Excel format.
        Workbook workbook = new Workbook(sourcePath, loadOptions);
        workbook.Save("output_converted.xlsx", SaveFormat.Xlsx);
    }
}
