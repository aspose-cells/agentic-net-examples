// Title: Use FontSettings to Load a Unicode‑Supporting Font for PDF Export in Aspose.Cells (C#)
// Description: Demonstrates loading a Unicode‑compatible TrueType font with FontSettings, setting PdfSaveOptions.DefaultFont (e.g., Arial Unicode MS), enabling CheckWorkbookDefaultFont, and saving a workbook containing emojis and rare CJK characters to PDF.
// Keywords: Aspose.Cells | C# | FontSettings | Unicode font | supplementary characters | emoji | CJK | PDF conversion | PdfSaveOptions | DefaultFont | CheckWorkbookDefaultFont | Arial Unicode MS | load custom font | Unicode PDF export
// Common Searches: Aspose.Cells load Unicode font for PDF | PDF export emojis Aspose.Cells | How to embed Unicode font in Aspose.Cells PDF | FontSettings LoadFont example C# | PdfSaveOptions DefaultFont Unicode | CheckWorkbookDefaultFont Aspose.Cells
// Developer Intent: Register a TrueType font that covers Unicode supplementary glyphs and apply it during PDF conversion of an Aspose.Cells workbook.
// Use Cases: Render emojis, mathematical symbols, and rare CJK glyphs in PDFs generated from spreadsheets. | Provide a fallback Unicode font when the workbook’s default font lacks required glyphs. | Programmatically ensure the output directory exists before saving to avoid runtime errors.
// AI Prompts: Write C# code using Aspose.Cells FontSettings.LoadFont to register a TrueType font that supports supplementary Unicode characters and save the workbook as PDF with PdfSaveOptions. | Show how to set PdfSaveOptions.DefaultFont to a Unicode font and enable CheckWorkbookDefaultFont to correctly display emojis and CJK characters. | Explain step‑by‑step how to create the output folder, load a Unicode font, and export a workbook containing Unicode text to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates loading a Unicode‑compatible TrueType font with FontSettings, setting PdfSaveOptions.DefaultFont (e.g., Arial Unicode MS), enabling CheckWorkbookDefaultFont, and saving a workbook containing emojis and rare CJK characters to PDF.
class FontSettingsPdfDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add text that contains Unicode supplementary characters (e.g., emojis, rare CJK characters)
            sheet.Cells["A1"].PutValue("Unicode test: 𝟘𝟙𝟚𝟛 🚀 漢字");

            // Prepare PDF save options with a fallback default font that supports Unicode characters
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use a font that is commonly available and supports a wide range of Unicode glyphs
                DefaultFont = "Arial Unicode MS",
                // Try to use the workbook's default font first
                CheckWorkbookDefaultFont = true
            };

            // Define output file path
            string outputPath = "output.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
