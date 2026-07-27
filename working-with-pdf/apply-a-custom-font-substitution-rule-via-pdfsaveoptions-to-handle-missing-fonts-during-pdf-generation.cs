using System;
using Aspose.Cells;

public class FontSubstitutionPdfDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample text that uses a font which might be missing on the target system
        sheet.Cells["A1"].PutValue("Sample text with missing font");

        // Apply the original font (e.g., Arial) to the cell style
        Style style = workbook.CreateStyle();
        style.Font.Name = "Arial"; // Assume Arial may not be installed
        sheet.Cells["A1"].SetStyle(style);

        // Define substitute fonts for the original font
        string originalFont = "Arial";
        string[] substituteFonts = new string[] { "Liberation Sans", "Helvetica", "Verdana" };
        FontConfigs.SetFontSubstitutes(originalFont, substituteFonts);

        // Configure PDF save options with a default fallback font
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DefaultFont = "Liberation Sans"
        };

        // Save the workbook as PDF using the configured options
        workbook.Save("FontSubstitutionDemo.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example.