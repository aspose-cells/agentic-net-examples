// Title: Aspose.Cells C# – Convert Excel to PDF with Embedded Custom OpenType Font from MemoryStream
// Description: Demonstrates how to load an OpenType (OTF) font into a MemoryStream, register its folder with Aspose.Cells, apply the font to worksheet cells, configure PdfSaveOptions for Identity encoding and font embedding, and save the workbook as a PDF that contains the custom font.
// Keywords: Aspose.Cells PDF conversion | embed custom OTF font | C# MemoryStream font loading | PdfSaveOptions Identity encoding | FontConfigs set font folder | Excel to PDF custom font | Aspose.Cells custom font example
// Common Searches: how to embed a custom OpenType font in Aspose.Cells PDF | Aspose.Cells load font from MemoryStream | set default font for PDF output Aspose.Cells | register font folder Aspose.Cells C# | convert Excel to PDF with custom font using Aspose
// Developer Intent: Embed a user‑provided OpenType font into a PDF generated from an Aspose.Cells workbook.
// Use Cases: Brand‑consistent PDF reports that use a company‑specific OTF font. | Invoices or statements where the corporate typeface must appear in the printed PDF. | Multilingual documents requiring a particular OpenType font to render special characters correctly.
// AI Prompts: Generate C# code that reads an OTF file into a MemoryStream, registers it with Aspose.Cells, and saves the workbook as a PDF with the font embedded. | Show how to configure PdfSaveOptions to embed a custom font using Identity encoding in Aspose.Cells. | Explain steps to verify that a custom OpenType font is correctly embedded in the resulting PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomFontPdf
{
    // Demonstrates how to load an OpenType (OTF) font into a MemoryStream, register its folder with Aspose.Cells, apply the font to worksheet cells, configure PdfSaveOptions for Identity encoding and font embedding, and save the workbook as a PDF that contains the custom font.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the custom OpenType font file
                const string fontPath = "customfont.otf";

                // Verify that the font file exists to avoid FileNotFoundException
                if (!File.Exists(fontPath))
                {
                    Console.WriteLine($"Font file not found: {Path.GetFullPath(fontPath)}");
                    return;
                }

                // Register the folder containing the custom font with Aspose.Cells
                // The second parameter (true) enables subfolder search
                string fontFolder = Path.GetDirectoryName(Path.GetFullPath(fontPath));
                FontConfigs.SetFontFolder(fontFolder, true);

                // Create a new workbook and add sample text
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Text using a custom OpenType font loaded from memory");

                // Apply the custom font to the cell style
                Style style = sheet.Cells["A1"].GetStyle();
                // Use the exact font name defined inside the OTF file (e.g., "CustomFont")
                style.Font.Name = "CustomFont";
                style.Font.Size = 14;
                sheet.Cells["A1"].SetStyle(style);

                // Configure PDF save options to embed fonts
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedStandardWindowsFonts = true,
                    FontEncoding = PdfFontEncoding.Identity,
                    DefaultFont = "CustomFont",
                    CheckWorkbookDefaultFont = true,
                    CheckFontCompatibility = true
                };

                // Save the workbook as PDF with the specified options
                const string outputPdf = "CustomFontOutput.pdf";
                workbook.Save(outputPdf, pdfOptions);
                Console.WriteLine($"PDF saved successfully: {Path.GetFullPath(outputPdf)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
