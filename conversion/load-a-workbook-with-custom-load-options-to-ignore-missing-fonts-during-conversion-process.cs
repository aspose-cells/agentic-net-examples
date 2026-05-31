using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Source Excel file to be loaded and converted
        string sourcePath = "input.xlsx";

        // Destination file after conversion (e.g., PDF)
        string destPath = "output.pdf";

        // ---------- Create custom LoadOptions ----------
        LoadOptions loadOptions = new LoadOptions();

        // Initialize individual font configurations
        loadOptions.FontConfigs = new IndividualFontConfigs();

        // Define font substitution rules for missing fonts.
        // Example: if the workbook references "Calibri" which is not available,
        // substitute it with "Arial". Add more rules as needed.
        loadOptions.FontConfigs.SetFontSubstitutes("Calibri", new string[] { "Arial" });
        loadOptions.FontConfigs.SetFontSubstitutes("Times New Roman", new string[] { "Arial" });

        // ---------- Load the workbook with the custom options ----------
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // (Optional) Perform any workbook manipulations here
        // e.g., workbook.Worksheets[0].Cells["A1"].PutValue("Loaded with custom font handling");

        // ---------- Convert the workbook to another format ----------
        // Create save options for the target format (PDF in this example)
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Use ConversionUtility with the same LoadOptions to ensure font handling during conversion
        ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

        Console.WriteLine("Conversion completed successfully.");
    }
}