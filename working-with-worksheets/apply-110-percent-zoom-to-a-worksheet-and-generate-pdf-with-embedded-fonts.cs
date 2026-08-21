// Title: Set Worksheet Zoom to 110% and Export PDF with Embedded Fonts using Aspose.Cells for .NET
// Description: Creates a workbook, applies a 110% view zoom to the first worksheet, configures PdfSaveOptions to embed standard Windows TrueType fonts, and saves the result as a PDF file.
// Keywords: Aspose.Cells worksheet zoom | 110% zoom Aspose.Cells | PDF font embedding Aspose.Cells | PdfSaveOptions embed fonts | C# export worksheet to PDF | Aspose.Cells view settings | embed Windows fonts PDF
// Common Searches: how to set worksheet zoom in Aspose.Cells .NET | embed fonts when saving PDF with Aspose.Cells | Aspose.Cells PdfSaveOptions example | export Excel to PDF with embedded fonts C# | increase worksheet view zoom Aspose.Cells
// Developer Intent: Apply a 110% view zoom to a worksheet and generate a PDF that includes embedded TrueType fonts.
// Use Cases: Produce PDFs that match the on‑screen layout for precise visual reporting. | Ensure PDF documents render correctly on machines lacking the original fonts. | Create printable reports where standard Windows fonts are guaranteed to appear.
// AI Prompts: Show code to set any custom zoom level (e.g., 150%) on a worksheet and embed all used fonts when exporting to PDF with Aspose.Cells for .NET. | Provide a method to verify that fonts are embedded in the generated PDF and handle embedding errors gracefully. | Explain how to embed non‑standard custom fonts in a PDF using Aspose.Cells, including required license settings and font folder configuration.

using System;
using Aspose.Cells;

// Creates a workbook, applies a 110% view zoom to the first worksheet, configures PdfSaveOptions to embed standard Windows TrueType fonts, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Apply 110% zoom to the worksheet (view zoom)
            worksheet.Zoom = 110; // Valid range is 10 to 400

            // Configure PDF save options to embed fonts in the output PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed standard Windows TrueType fonts (required for font embedding)
                EmbedStandardWindowsFonts = true
                // FontEncoding property omitted because PdfFontEncoding is not available in this version
            };

            // Save the workbook as a PDF file with the specified options
            workbook.Save("Worksheet_Zoom_110.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
