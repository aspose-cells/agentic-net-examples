using System;
using Aspose.Cells;

namespace AsposeCellsFontSubstitutionDemo
{
    class Program
    {
        static void Main()
        {
            // Enable system font substitutes when a requested font is missing.
            FontConfigs.PreferSystemFontSubstitutes = true;

            // Create or load a workbook.
            // Here we create a sample workbook; replace with new Workbook("input.xlsx") to load an existing file.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample text with a possibly unavailable font");
            Style style = sheet.Cells["A1"].GetStyle();
            style.Font.Name = "NonExistentFont"; // Simulate missing font
            sheet.Cells["A1"].SetStyle(style);

            // Configure PDF save options to use the workbook's default font checking (system default will be used).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                CheckWorkbookDefaultFont = true   // Default value; kept for clarity
                // Do not set DefaultFont so the system default font is used.
            };

            // Save the workbook as PDF.
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("Workbook converted to PDF with system default font substitution.");
        }
    }
}