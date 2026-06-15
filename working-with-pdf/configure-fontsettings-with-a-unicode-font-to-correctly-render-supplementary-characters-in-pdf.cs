using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsUnicodePdfDemo
{
    class Program
    {
        static void Main()
        {
            // Set a Unicode capable default font globally (optional, can also be set per save options)
            // Example font that supports many Unicode ranges, including supplementary characters
            FontConfigs.DefaultFontName = "Arial Unicode MS";

            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add regular text and supplementary Unicode characters (e.g., emoji)
            sheet.Cells["A1"].PutValue("Regular text");
            sheet.Cells["A2"].PutValue("Supplementary characters: 😀 🎉 𠜎"); // Emoji and a CJK Extension B character

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use the Unicode font when a cell does not specify a compatible font
                DefaultFont = "Arial Unicode MS",
                // Try to use the workbook's default font first
                CheckWorkbookDefaultFont = true,
                // Ensure font compatibility checking is enabled (default true) so fallback fonts are used if needed
                CheckFontCompatibility = true
            };

            // Save the workbook as PDF with the configured options
            workbook.Save("UnicodeOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved successfully with Unicode font settings.");
        }
    }
}