// Title: Load a Unicode Font with FontConfigs and Convert an Aspose.Cells Workbook to PDF (C#)
// Description: Demonstrates how to point Aspose.Cells to a folder containing a Unicode‑capable font (e.g., Arial Unicode MS or Noto Sans), set it as the default, insert text with supplementary characters such as emoji, and configure PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, CheckFontCompatibility) so the PDF renders those glyphs correctly.
// Keywords: Aspose.Cells FontConfigs | custom font folder C# | Unicode supplementary characters PDF | emoji rendering Aspose.Cells | PdfSaveOptions DefaultFont | CheckWorkbookDefaultFont | CheckFontCompatibility | load Unicode font Aspose.Cells | C# PDF conversion Aspose.Cells
// Common Searches: how to load a Unicode font in Aspose.Cells before PDF export | set default font for emoji in Aspose.Cells PDF conversion | Aspose.Cells FontConfigs SetFontFolder example | render supplementary Unicode characters in PDF with Aspose.Cells | C# Aspose.Cells PDFSaveOptions Unicode support
// Developer Intent: Configure Aspose.Cells to use a Unicode‑compatible font so that supplementary characters are displayed correctly when saving a workbook as PDF.
// Use Cases: Create multilingual PDF reports that include CJK, mathematical symbols, or emoji by loading a font directory with a comprehensive Unicode font. | Generate financial statements where rare currency symbols or technical glyphs must appear intact in the PDF output. | Automate document pipelines that require accurate rendering of non‑BMP characters across global locales.
// AI Prompts: Write C# code that uses FontConfigs.SetFontFolder to load a custom Unicode font directory and saves a workbook to PDF with proper emoji rendering. | Explain the role of CheckWorkbookDefaultFont and CheckFontCompatibility in preserving supplementary Unicode characters during PDF conversion. | Provide a step‑by‑step troubleshooting guide for missing glyphs after exporting an Aspose.Cells workbook to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsUnicodeFontDemo
{
    // Demonstrates how to point Aspose.Cells to a folder containing a Unicode‑capable font (e.g., Arial Unicode MS or Noto Sans), set it as the default, insert text with supplementary characters such as emoji, and configure PdfSaveOptions (DefaultFont, CheckWorkbookDefaultFont, CheckFontCompatibility) so the PDF renders those glyphs correctly.
    class Program
    {
        static void Main()
        {
            // Specify the folder that contains the Unicode supporting font (e.g., "Arial Unicode MS" or "NotoSansCJK").
            // The folder can contain subfolders with additional fonts.
            FontConfigs.SetFontFolder(@"C:\UnicodeFonts", true);

            // Optionally set the default font name that Aspose.Cells will use when a cell does not specify a font.
            FontConfigs.DefaultFontName = "Arial Unicode MS";

            // Create a new workbook and access the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert text that includes Unicode supplementary characters (e.g., emoji or characters outside the BMP).
            // Example: U+1F600 GRINNING FACE 😀 (represented as a surrogate pair in .NET strings).
            sheet.Cells["A1"].PutValue("Unicode test: 😀 𝔘𝔫𝔦𝔠𝔬𝔡𝔢");

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use the font that supports the supplementary characters.
                DefaultFont = "Arial Unicode MS",
                // Ensure the workbook's default font is considered first.
                CheckWorkbookDefaultFont = true,
                // Keep font compatibility checking enabled to allow substitution if needed.
                CheckFontCompatibility = true
            };

            // Save the workbook as PDF using the configured options.
            workbook.Save("UnicodeOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with Unicode supplementary characters rendered correctly.");
        }
    }
}
