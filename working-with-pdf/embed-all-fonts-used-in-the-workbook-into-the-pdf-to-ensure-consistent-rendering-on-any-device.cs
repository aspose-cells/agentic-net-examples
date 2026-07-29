// Title: Embed All Fonts in PDF Export with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create or load an Excel workbook, apply custom fonts (including Unicode characters), configure PdfSaveOptions (EmbedStandardWindowsFonts, FontEncoding=Identity, DefaultFont, CheckWorkbookDefaultFont) and save the workbook as a PDF that contains every font used, guaranteeing identical rendering on any device.
// Keywords: Aspose.Cells PDF font embedding | C# embed fonts in PDF | PdfSaveOptions EmbedStandardWindowsFonts | Identity font encoding Aspose | DefaultFont Aspose.Cells | CheckWorkbookDefaultFont | Excel to PDF with custom fonts | Unicode PDF Aspose.Cells | preserve fonts in PDF export | cross‑device PDF rendering
// Common Searches: how to embed all fonts when converting Excel to PDF with Aspose.Cells | Aspose.Cells PdfSaveOptions EmbedStandardWindowsFonts example C# | set FontEncoding to Identity for PDF font embedding Aspose | defaultfont and checkworkbookdefaultfont usage in Aspose.Cells PDF export | embed Unicode fonts in PDF generated from Excel
// Developer Intent: Ensure every font referenced in the workbook is embedded in the generated PDF so the document looks identical on any platform.
// Use Cases: Export financial reports that mix Latin and Asian scripts to PDF while preserving the exact typography. | Create printable PDFs from Excel templates that rely on corporate custom fonts. | Distribute Excel‑derived PDFs to clients who may not have the source fonts installed.
// AI Prompts: Generate C# code that loads an existing .xlsx file and saves it as a PDF with all fonts embedded using Aspose.Cells. | Explain the role of EmbedStandardWindowsFonts, FontEncoding, DefaultFont, and CheckWorkbookDefaultFont in Aspose.Cells PDF generation. | Provide a verification checklist to confirm that fonts are truly embedded in a PDF produced by Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace EmbedAllFontsInPdf
{
    // Demonstrates how to create or load an Excel workbook, apply custom fonts (including Unicode characters), configure PdfSaveOptions (EmbedStandardWindowsFonts, FontEncoding=Identity, DefaultFont, CheckWorkbookDefaultFont) and save the workbook as a PDF that contains every font used, guaranteeing identical rendering on any device.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample text with various fonts");
            sheet.Cells["A2"].PutValue("中文字符测试"); // Unicode characters to test font embedding

            // Optionally apply different fonts to cells
            Style style1 = sheet.Cells["A1"].GetStyle();
            style1.Font.Name = "Arial";
            style1.Font.Size = 14;
            sheet.Cells["A1"].SetStyle(style1);

            Style style2 = sheet.Cells["A2"].GetStyle();
            style2.Font.Name = "Times New Roman";
            style2.Font.Size = 12;
            sheet.Cells["A2"].SetStyle(style2);

            // Configure PDF save options to embed all fonts
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure TrueType fonts are embedded (default is true)
                EmbedStandardWindowsFonts = true,

                // Use Identity encoding to embed fonts for all characters
                FontEncoding = PdfFontEncoding.Identity,

                // Set a default font in case a cell's font is missing
                DefaultFont = "Arial",

                // Try to use workbook's default font for Unicode characters
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as PDF with embedded fonts
            workbook.Save("EmbeddedFontsOutput.pdf", pdfOptions);
        }
    }
}
